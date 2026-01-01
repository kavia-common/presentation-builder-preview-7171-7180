import React, { useEffect, useMemo, useRef, useState } from 'react';
import './App.css';

const LAYOUTS = [
  { value: 'title', label: 'Title slide' },
  { value: 'titleBody', label: 'Title + Body' },
  { value: 'twoColumn', label: 'Two column' },
];

/**
 * Creates a stable-ish unique id for slides without adding a UUID dependency.
 */
function createId() {
  return `s_${Math.random().toString(16).slice(2)}_${Date.now().toString(16)}`;
}

/**
 * Minimal text sanitization for rendering into the in-app preview.
 */
function safeText(value) {
  if (value == null) return '';
  return String(value);
}

/**
 * Convenience clamp.
 */
function clamp(n, min, max) {
  return Math.max(min, Math.min(max, n));
}

/**
 * Very small XML escape helper.
 */
function escapeXml(str) {
  return safeText(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

/**
 * Convert a JS string to UTF-8 bytes.
 */
function utf8Bytes(str) {
  return new TextEncoder().encode(str);
}

/**
 * Compute CRC32 for bytes. Needed by ZIP format even when using "store" (no compression).
 */
function crc32(bytes) {
  // Precompute table once.
  if (!crc32._table) {
    const table = new Uint32Array(256);
    for (let i = 0; i < 256; i += 1) {
      let c = i;
      for (let k = 0; k < 8; k += 1) {
        c = (c & 1) ? (0xedb88320 ^ (c >>> 1)) : (c >>> 1);
      }
      table[i] = c >>> 0;
    }
    crc32._table = table;
  }
  let crc = 0xffffffff;
  for (let i = 0; i < bytes.length; i += 1) {
    crc = crc32._table[(crc ^ bytes[i]) & 0xff] ^ (crc >>> 8);
  }
  return (crc ^ 0xffffffff) >>> 0;
}

/**
 * Minimal ZIP builder using STORE (no compression).
 * Enough to produce a valid .pptx (Open XML package is a ZIP).
 */
function zipStore(files) {
  const fileEntries = [];
  let offset = 0;

  const localParts = [];
  const centralParts = [];

  const pushU16 = (arr, v) => {
    arr.push(v & 0xff, (v >>> 8) & 0xff);
  };
  const pushU32 = (arr, v) => {
    arr.push(v & 0xff, (v >>> 8) & 0xff, (v >>> 16) & 0xff, (v >>> 24) & 0xff);
  };

  for (const f of files) {
    const nameBytes = utf8Bytes(f.name);
    const data = f.data instanceof Uint8Array ? f.data : new Uint8Array(f.data);
    const crc = crc32(data);
    const compSize = data.length;
    const uncompSize = data.length;

    // Local file header
    const localHeader = [];
    pushU32(localHeader, 0x04034b50);
    pushU16(localHeader, 20); // version needed
    pushU16(localHeader, 0); // flags
    pushU16(localHeader, 0); // method = store
    pushU16(localHeader, 0); // mod time
    pushU16(localHeader, 0); // mod date
    pushU32(localHeader, crc);
    pushU32(localHeader, compSize);
    pushU32(localHeader, uncompSize);
    pushU16(localHeader, nameBytes.length);
    pushU16(localHeader, 0); // extra length

    localParts.push(new Uint8Array(localHeader));
    localParts.push(nameBytes);
    localParts.push(data);

    const localHeaderOffset = offset;
    offset += localHeader.length + nameBytes.length + data.length;

    // Central directory header
    const centralHeader = [];
    pushU32(centralHeader, 0x02014b50);
    pushU16(centralHeader, 20); // version made by
    pushU16(centralHeader, 20); // version needed
    pushU16(centralHeader, 0); // flags
    pushU16(centralHeader, 0); // method
    pushU16(centralHeader, 0); // time
    pushU16(centralHeader, 0); // date
    pushU32(centralHeader, crc);
    pushU32(centralHeader, compSize);
    pushU32(centralHeader, uncompSize);
    pushU16(centralHeader, nameBytes.length);
    pushU16(centralHeader, 0); // extra
    pushU16(centralHeader, 0); // comment
    pushU16(centralHeader, 0); // disk
    pushU16(centralHeader, 0); // int attrs
    pushU32(centralHeader, 0); // ext attrs
    pushU32(centralHeader, localHeaderOffset);

    centralParts.push(new Uint8Array(centralHeader));
    centralParts.push(nameBytes);

    fileEntries.push({ name: f.name });
  }

  const centralStart = offset;
  for (const part of centralParts) offset += part.length;
  const centralSize = offset - centralStart;

  // End of central directory record
  const eocd = [];
  const totalEntries = fileEntries.length;
  const commentLen = 0;
  pushU32(eocd, 0x06054b50);
  pushU16(eocd, 0); // disk
  pushU16(eocd, 0); // disk where central starts
  pushU16(eocd, totalEntries);
  pushU16(eocd, totalEntries);
  pushU32(eocd, centralSize);
  pushU32(eocd, centralStart);
  pushU16(eocd, commentLen);

  const parts = [...localParts, ...centralParts, new Uint8Array(eocd)];
  const totalLen = parts.reduce((sum, p) => sum + p.length, 0);
  const out = new Uint8Array(totalLen);
  let pos = 0;
  for (const p of parts) {
    out.set(p, pos);
    pos += p.length;
  }
  return out;
}

/**
 * Read a File object as an ArrayBuffer.
 */
async function readFileAsArrayBuffer(file) {
  if (!file) throw new Error('No file provided');
  return await new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onerror = () => reject(new Error('Failed to read file'));
    reader.onload = () => resolve(reader.result);
    reader.readAsArrayBuffer(file);
  });
}

/**
 * Convert a uint8 buffer to base64 (for embedding in PPTX).
 */
function u8ToBase64(bytes) {
  let binary = '';
  for (let i = 0; i < bytes.length; i += 1) {
    binary += String.fromCharCode(bytes[i]);
  }
  return btoa(binary);
}

/**
 * Creates minimal PPTX XML parts.
 * Note: This is a lightweight, standards-based PPTX generator (no Node APIs).
 * It supports:
 * - A locked cover slide (slide1) with background image + name/date overlay
 * - Titles + body text and a simple 2-column approximation for subsequent slides
 */
function buildPptxBytes({ slides, title, cover }) {
  const slideCount = slides.length;

  const contentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  <Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
  <Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
  <Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>
  <Override PartName="/ppt/media/cover.png" ContentType="image/png"/>
  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
  ${Array.from({ length: slideCount - 1 }, (_, i) => {
    const n = i + 2;
    return `<Override PartName="/ppt/slides/slide${n}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`;
  }).join('\n  ')}
</Types>`;

  const rootRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
</Relationships>`;

  const presRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  ${Array.from({ length: slideCount }, (_, i) => {
    const n = i + 1;
    return `<Relationship Id="rId${n}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${n}.xml"/>`;
  }).join('\n  ')}
  <Relationship Id="rId${slideCount + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>
  <Relationship Id="rId${slideCount + 2}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
</Relationships>`;

  const slideLayoutRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
</Relationships>`;

  const slideMasterRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>
</Relationships>`;

  // Presentation: references slides and a single slide master.
  const presentationXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:sldMasterIdLst>
    <p:sldMasterId id="2147483648" r:id="rId${slideCount + 1}"/>
  </p:sldMasterIdLst>
  <p:sldIdLst>
    ${Array.from({ length: slideCount }, (_, i) => {
      const n = i + 1;
      // ids must be unique 32-bit int. We can use 256 + n.
      return `<p:sldId id="${256 + n}" r:id="rId${n}"/>`;
    }).join('\n    ')}
  </p:sldIdLst>
  <p:sldSz cx="12192000" cy="6858000" type="screen16x9"/>
  <p:notesSz cx="6858000" cy="9144000"/>
</p:presentation>`;

  // Minimal theme: defines basic colors; Office is fine with a minimal theme.
  const themeXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Ocean Professional">
  <a:themeElements>
    <a:clrScheme name="Ocean">
      <a:dk1><a:srgbClr val="111827"/></a:dk1>
      <a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
      <a:dk2><a:srgbClr val="1F2937"/></a:dk2>
      <a:lt2><a:srgbClr val="F9FAFB"/></a:lt2>
      <a:accent1><a:srgbClr val="2563EB"/></a:accent1>
      <a:accent2><a:srgbClr val="F59E0B"/></a:accent2>
      <a:accent3><a:srgbClr val="10B981"/></a:accent3>
      <a:accent4><a:srgbClr val="EF4444"/></a:accent4>
      <a:accent5><a:srgbClr val="A855F7"/></a:accent5>
      <a:accent6><a:srgbClr val="06B6D4"/></a:accent6>
      <a:hlink><a:srgbClr val="2563EB"/></a:hlink>
      <a:folHlink><a:srgbClr val="1D4ED8"/></a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="Ocean Fonts">
      <a:majorFont><a:latin typeface="Aptos Display"/></a:majorFont>
      <a:minorFont><a:latin typeface="Aptos"/></a:minorFont>
    </a:fontScheme>
    <a:fmtScheme name="Ocean Format"/>
  </a:themeElements>
</a:theme>`;

  // Slide master + layout: minimal scaffolding.
  const slideMasterXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:bg><p:bgRef idx="1001"><a:schemeClr val="lt2"/></p:bgRef></p:bg>
    <p:spTree>
      <p:nvGrpSpPr>
        <p:cNvPr id="1" name=""/>
        <p:cNvGrpSpPr/>
        <p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr/>
    </p:spTree>
  </p:cSld>
  <p:sldLayoutIdLst>
    <p:sldLayoutId id="1" r:id="rId1"/>
  </p:sldLayoutIdLst>
</p:sldMaster>`;

  const slideLayoutXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="title">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr>
        <p:cNvPr id="1" name=""/>
        <p:cNvGrpSpPr/>
        <p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr/>
    </p:spTree>
  </p:cSld>
</p:sldLayout>`;

  // Slide relationships:
  // - cover slide also relates to cover.png
  const coverSlideRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/cover.png"/>
</Relationships>`;

  // Subsequent slides: relate to slideLayout1
  const slideRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
</Relationships>`;

  const slideXmlCover = (s0) => {
    const coverTitle = escapeXml(safeText(s0.title).trim() || safeText(title).trim() || 'Cover');
    const nameText = escapeXml(safeText(cover?.name).trim() || '');
    const dateText = escapeXml(safeText(cover?.date).trim() || '');

    // Full slide image
    const slideW = 13.333 * 914400;
    const slideH = 7.5 * 914400;

    // Overlay texts (approx position matching the provided image)
    const titleX = 0.8 * 914400;
    const titleY = 2.1 * 914400;
    const titleW = 6.4 * 914400;
    const titleH = 0.8 * 914400;

    const metaX = 0.8 * 914400;
    const nameY = 2.95 * 914400;
    const dateY = 3.45 * 914400;
    const metaW = 6.2 * 914400;
    const metaH = 0.45 * 914400;

    const mkTextBox = ({ id, name, x, y, w, h, text, bold, size, colorHex }) => `      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="${id}" name="${name}"/>
          <p:cNvSpPr txBox="1"/>
          <p:nvPr/>
        </p:nvSpPr>
        <p:spPr>
          <a:xfrm>
            <a:off x="${x}" y="${y}"/>
            <a:ext cx="${w}" cy="${h}"/>
          </a:xfrm>
        </p:spPr>
        <p:txBody>
          <a:bodyPr wrap="square"/>
          <a:lstStyle/>
          <a:p>
            <a:r>
              <a:rPr lang="en-US" b="${bold ? 1 : 0}" sz="${size}">
                <a:solidFill><a:srgbClr val="${colorHex}"/></a:solidFill>
                <a:latin typeface="Aptos Display"/>
              </a:rPr>
              <a:t>${text}</a:t>
            </a:r>
          </a:p>
        </p:txBody>
      </p:sp>`;

    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr>
        <p:cNvPr id="1" name=""/>
        <p:cNvGrpSpPr/>
        <p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr/>

      <!-- Background image -->
      <p:pic>
        <p:nvPicPr>
          <p:cNvPr id="2" name="Cover Image"/>
          <p:cNvPicPr/>
          <p:nvPr/>
        </p:nvPicPr>
        <p:blipFill>
          <a:blip r:embed="rId2"/>
          <a:stretch><a:fillRect/></a:stretch>
        </p:blipFill>
        <p:spPr>
          <a:xfrm>
            <a:off x="0" y="0"/>
            <a:ext cx="${slideW}" cy="${slideH}"/>
          </a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
        </p:spPr>
      </p:pic>

${mkTextBox({
  id: 10,
  name: 'Cover Title',
  x: titleX,
  y: titleY,
  w: titleW,
  h: titleH,
  text: coverTitle,
  bold: true,
  size: 3200,
  colorHex: 'FFFFFF',
})}

${mkTextBox({
  id: 11,
  name: 'Cover Name',
  x: metaX,
  y: nameY,
  w: metaW,
  h: metaH,
  text: nameText ? `Name : ${nameText}` : 'Name :',
  bold: true,
  size: 1800,
  colorHex: 'FFFFFF',
})}

${mkTextBox({
  id: 12,
  name: 'Cover Date',
  x: metaX,
  y: dateY,
  w: metaW,
  h: metaH,
  text: dateText ? `Date : ${dateText}` : 'Date :',
  bold: true,
  size: 1800,
  colorHex: 'FFFFFF',
})}

    </p:spTree>
  </p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sld>`;
  };

  // Simple slide XML with 2 text boxes (title + body).
  const slideXmlFor = (slide, idx) => {
    const titleText = escapeXml(safeText(slide.title).trim() || `Slide ${idx + 1}`);
    const bodyText = escapeXml(safeText(slide.body).trim());

    const bodyLines = safeText(slide.body)
      .split('\n')
      .map(l => l.trim())
      .filter(Boolean);

    const mkParagraphs = (lines) =>
      lines.length
        ? lines
            .map(
              (l) => `<a:p><a:r><a:rPr lang="en-US" dirty="0"/><a:t>${escapeXml(l)}</a:t></a:r></a:p>`
            )
            .join('')
        : `<a:p><a:r><a:rPr lang="en-US" dirty="0"/><a:t></a:t></a:r></a:p>`;

    // Position in EMUs (1 inch = 914400). Slide size is 13.333" x 7.5". We keep it simple.
    const titleX = 0.7 * 914400;
    const titleY = 0.6 * 914400;
    const titleW = 11.9 * 914400;
    const titleH = 0.9 * 914400;

    const bodyY = 1.7 * 914400;
    const bodyH = 5.2 * 914400;

    const isTitleOnly = slide.layout === 'title';
    const isTwoCol = slide.layout === 'twoColumn';

    const bodyBlocks = (() => {
      if (isTitleOnly) {
        // Subtitle-style single body box (optional)
        return bodyText
          ? `<p:sp>
  <p:nvSpPr>
    <p:cNvPr id="3" name="Subtitle ${idx + 1}"/>
    <p:cNvSpPr txBox="1"/>
    <p:nvPr/>
  </p:nvSpPr>
  <p:spPr>
    <a:xfrm>
      <a:off x="${0.7 * 914400}" y="${1.6 * 914400}"/>
      <a:ext cx="${11.9 * 914400}" cy="${0.9 * 914400}"/>
    </a:xfrm>
  </p:spPr>
  <p:txBody>
    <a:bodyPr wrap="square"/>
    <a:lstStyle/>
    <a:p>
      <a:r>
        <a:rPr lang="en-US" sz="2400" />
        <a:t>${bodyText}</a:t>
      </a:r>
    </a:p>
  </p:txBody>
</p:sp>`
          : '';
      }

      if (isTwoCol) {
        const left = bodyLines.slice(0, Math.ceil(bodyLines.length / 2));
        const right = bodyLines.slice(Math.ceil(bodyLines.length / 2));

        return `<p:sp>
  <p:nvSpPr>
    <p:cNvPr id="3" name="BodyLeft ${idx + 1}"/>
    <p:cNvSpPr txBox="1"/>
    <p:nvPr/>
  </p:nvSpPr>
  <p:spPr>
    <a:xfrm>
      <a:off x="${0.7 * 914400}" y="${bodyY}"/>
      <a:ext cx="${5.9 * 914400}" cy="${bodyH}"/>
    </a:xfrm>
  </p:spPr>
  <p:txBody>
    <a:bodyPr wrap="square"/>
    <a:lstStyle/>
    ${mkParagraphs(left.length ? left : [''])}
  </p:txBody>
</p:sp>
<p:sp>
  <p:nvSpPr>
    <p:cNvPr id="4" name="BodyRight ${idx + 1}"/>
    <p:cNvSpPr txBox="1"/>
    <p:nvPr/>
  </p:nvSpPr>
  <p:spPr>
    <a:xfrm>
      <a:off x="${6.9 * 914400}" y="${bodyY}"/>
      <a:ext cx="${5.7 * 914400}" cy="${bodyH}"/>
    </a:xfrm>
  </p:spPr>
  <p:txBody>
    <a:bodyPr wrap="square"/>
    <a:lstStyle/>
    ${mkParagraphs(right.length ? right : [''])}
  </p:txBody>
</p:sp>`;
      }

      // Default title+body
      return `<p:sp>
  <p:nvSpPr>
    <p:cNvPr id="3" name="Body ${idx + 1}"/>
    <p:cNvSpPr txBox="1"/>
    <p:nvPr/>
  </p:nvSpPr>
  <p:spPr>
    <a:xfrm>
      <a:off x="${0.9 * 914400}" y="${bodyY}"/>
      <a:ext cx="${11.3 * 914400}" cy="${bodyH}"/>
    </a:xfrm>
  </p:spPr>
  <p:txBody>
    <a:bodyPr wrap="square"/>
    <a:lstStyle/>
    ${mkParagraphs(bodyLines.length ? bodyLines : bodyText ? [bodyText] : [''])}
  </p:txBody>
</p:sp>`;
    })();

    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:bg>
      <p:bgPr>
        <a:solidFill><a:srgbClr val="F9FAFB"/></a:solidFill>
      </p:bgPr>
    </p:bg>
    <p:spTree>
      <p:nvGrpSpPr>
        <p:cNvPr id="1" name=""/>
        <p:cNvGrpSpPr/>
        <p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr/>

      <!-- Top accent line -->
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="2" name="TopBar ${idx + 1}"/>
          <p:cNvSpPr/>
          <p:nvPr/>
        </p:nvSpPr>
        <p:spPr>
          <a:xfrm>
            <a:off x="0" y="0"/>
            <a:ext cx="${13.333 * 914400}" cy="${0.15 * 914400}"/>
          </a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
          <a:solidFill><a:srgbClr val="2563EB"/></a:solidFill>
          <a:ln><a:solidFill><a:srgbClr val="2563EB"/></a:solidFill></a:ln>
        </p:spPr>
      </p:sp>

      <!-- Title -->
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="10" name="Title ${idx + 1}"/>
          <p:cNvSpPr txBox="1"/>
          <p:nvPr/>
        </p:nvSpPr>
        <p:spPr>
          <a:xfrm>
            <a:off x="${titleX}" y="${titleY}"/>
            <a:ext cx="${titleW}" cy="${titleH}"/>
          </a:xfrm>
        </p:spPr>
        <p:txBody>
          <a:bodyPr wrap="square"/>
          <a:lstStyle/>
          <a:p>
            <a:r>
              <a:rPr lang="en-US" b="1" sz="3800">
                <a:solidFill><a:srgbClr val="111827"/></a:solidFill>
                <a:latin typeface="Aptos Display"/>
              </a:rPr>
              <a:t>${titleText}</a:t>
            </a:r>
          </a:p>
        </p:txBody>
      </p:sp>

      ${bodyBlocks}
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sld>`;
  };

  const coverPngBytes =
    cover?.imageBytes instanceof Uint8Array ? cover.imageBytes : new Uint8Array(cover?.imageBytes || []);

  const files = [
    { name: '[Content_Types].xml', data: utf8Bytes(contentTypes) },
    { name: '_rels/.rels', data: utf8Bytes(rootRels) },
    { name: 'ppt/presentation.xml', data: utf8Bytes(presentationXml) },
    { name: 'ppt/_rels/presentation.xml.rels', data: utf8Bytes(presRels) },
    { name: 'ppt/theme/theme1.xml', data: utf8Bytes(themeXml) },
    { name: 'ppt/slideMasters/slideMaster1.xml', data: utf8Bytes(slideMasterXml) },
    { name: 'ppt/slideMasters/_rels/slideMaster1.xml.rels', data: utf8Bytes(slideMasterRels) },
    { name: 'ppt/slideLayouts/slideLayout1.xml', data: utf8Bytes(slideLayoutXml) },
    { name: 'ppt/slideLayouts/_rels/slideLayout1.xml.rels', data: utf8Bytes(slideLayoutRels) },
    { name: 'ppt/media/cover.png', data: coverPngBytes },
    // Slides
    { name: 'ppt/slides/slide1.xml', data: utf8Bytes(slideXmlCover(slides[0])) },
    { name: 'ppt/slides/_rels/slide1.xml.rels', data: utf8Bytes(coverSlideRels) },
    ...slides.slice(1).flatMap((s, i) => {
      const slideNumber = i + 2; // because slide1 is cover
      return [
        { name: `ppt/slides/slide${slideNumber}.xml`, data: utf8Bytes(slideXmlFor(s, slideNumber - 1)) },
        { name: `ppt/slides/_rels/slide${slideNumber}.xml.rels`, data: utf8Bytes(slideRels) },
      ];
    }),
  ];

  return zipStore(files);
}

/**
 * Trigger a file download in the browser.
 */
function downloadBytes(bytes, fileName) {
  const blob = new Blob([bytes], {
    type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
  });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = fileName;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

function SlideThumbnail({ slide, index, isSelected, onSelect }) {
  const title = safeText(slide.title).trim() || `Slide ${index + 1}`;
  const body = safeText(slide.body).trim();

  return (
    <button
      type="button"
      className={`thumb ${isSelected ? 'thumb--selected' : ''}`}
      onClick={onSelect}
      aria-label={`Select slide ${index + 1}`}
    >
      <div className="thumb__frame" aria-hidden="true">
        <div className={`thumb__layout thumb__layout--${slide.layout}`}>
          <div className="thumb__title">{title}</div>
          {slide.layout !== 'title' ? (
            <div className="thumb__body">{body || 'Body…'}</div>
          ) : (
            <div className="thumb__subtitle">{body || 'Subtitle…'}</div>
          )}
        </div>
      </div>
      <div className="thumb__meta">
        <span className="thumb__index">{index + 1}</span>
        <span className="thumb__label">{LAYOUTS.find(l => l.value === slide.layout)?.label}</span>
      </div>
    </button>
  );
}

function CoverSlidePreview({ slide, cover }) {
  const title = safeText(slide.title).trim() || 'Cover';
  const name = safeText(cover?.name).trim();
  const date = safeText(cover?.date).trim();
  const imgSrc = cover?.imagePreviewUrl || '/assets/cover-slide.png';

  return (
    <div className="preview">
      <div className="preview__paper preview__paper--cover">
        <div
          className="preview__coverBg"
          aria-hidden="true"
          style={{ backgroundImage: `url("${imgSrc}")` }}
        />
        <div className="preview__coverShade" aria-hidden="true" />
        <div className="preview__coverOverlay">
          <div className="preview__coverBadgeRow">
            <span className="preview__coverBadge">Locked global cover • Slide 1</span>
          </div>

          <div className="preview__coverTitle">{title}</div>

          <div className="preview__coverMeta" role="group" aria-label="Cover fields">
            <div className="preview__coverMetaLine">
              <span className="preview__coverMetaKey">Name :</span>
              <span className="preview__coverMetaVal">{name || '—'}</span>
            </div>
            <div className="preview__coverMetaLine">
              <span className="preview__coverMetaKey">Date :</span>
              <span className="preview__coverMetaVal">{date || '—'}</span>
            </div>
          </div>
        </div>
      </div>
      <div className="preview__note">Cover slide is locked and always remains first.</div>
    </div>
  );
}

function SlidePreview({ slide, cover }) {
  // Cover slide is special and locked (index 0)
  if (slide.isCover) {
    return <CoverSlidePreview slide={slide} cover={cover} />;
  }

  const title = safeText(slide.title).trim() || 'Untitled Slide';
  const body = safeText(slide.body).trim();

  return (
    <div className="preview">
      <div className="preview__paper">
        <div className="preview__topbar" aria-hidden="true" />
        <div className="preview__content">
          <div className="preview__title">{title}</div>

          {slide.layout === 'title' ? (
            <div className="preview__subtitle">{body || 'Subtitle text…'}</div>
          ) : slide.layout === 'twoColumn' ? (
            <div className="preview__twoCol">
              {(() => {
                const parts = body.split('\n').filter(Boolean);
                const left = parts.slice(0, Math.ceil(parts.length / 2)).join('\n');
                const right = parts.slice(Math.ceil(parts.length / 2)).join('\n');
                return (
                  <>
                    <div className="preview__col">
                      <div className="preview__card">
                        <pre className="preview__pre">{left || '• Left column\n• Add points'}</pre>
                      </div>
                    </div>
                    <div className="preview__col">
                      <div className="preview__card">
                        <pre className="preview__pre">{right || '• Right column\n• Add points'}</pre>
                      </div>
                    </div>
                  </>
                );
              })()}
            </div>
          ) : (
            <div className="preview__card">
              <pre className="preview__pre">{body || 'Add slide body text…'}</pre>
            </div>
          )}
        </div>
      </div>
      <div className="preview__note">Preview is a lightweight approximation of the final slide.</div>
    </div>
  );
}

// PUBLIC_INTERFACE
function App() {
  const [theme, setTheme] = useState('light');

  const [presentationTitle, setPresentationTitle] = useState('Ocean Professional Deck');

  // Cover state (locked global slide 1)
  const [coverName, setCoverName] = useState('Subrata B');
  const [coverDate, setCoverDate] = useState('24 Dec 2023');
  const [coverImageBytes, setCoverImageBytes] = useState(null);
  const [coverImagePreviewUrl, setCoverImagePreviewUrl] = useState(null);

  // Create the initial cover bytes from the public asset, so PPTX always includes the image.
  useEffect(() => {
    let cancelled = false;

    async function loadDefaultCoverBytes() {
      try {
        const res = await fetch('/assets/cover-slide.png', { cache: 'no-cache' });
        const buf = await res.arrayBuffer();
        if (cancelled) return;
        setCoverImageBytes(new Uint8Array(buf));
      } catch (e) {
        // If it fails, PPTX will still generate; cover will just have no image bytes.
        // Keep silent to avoid disrupting UX.
      }
    }

    loadDefaultCoverBytes();

    return () => {
      cancelled = true;
    };
  }, []);

  const [slides, setSlides] = useState(() => [
    {
      id: 'cover', // stable id
      title: 'Weekly Metrics',
      body: '',
      layout: 'title',
      isCover: true,
      locked: true,
    },
    {
      id: createId(),
      title: 'Agenda',
      body: '• Performance highlights\n• Key initiatives\n• Risks & mitigations\n• Next steps',
      layout: 'titleBody',
    },
    {
      id: createId(),
      title: 'KPIs',
      body: 'Revenue: +18%\nRetention: 96%\nNPS: 54\nPipeline: $3.2M',
      layout: 'twoColumn',
    },
  ]);

  const [selectedSlideId, setSelectedSlideId] = useState(slides[0]?.id || null);
  const [status, setStatus] = useState({ type: 'idle', message: '' }); // idle | error | success | working
  const [isGenerating, setIsGenerating] = useState(false);

  const selectedSlideIndex = useMemo(() => slides.findIndex(s => s.id === selectedSlideId), [
    slides,
    selectedSlideId,
  ]);
  const selectedSlide = slides[clamp(selectedSlideIndex, 0, slides.length - 1)];

  const titleInputRef = useRef(null);

  useEffect(() => {
    document.documentElement.setAttribute('data-theme', theme);
  }, [theme]);

  // Ensure selected slide remains valid after deletions/reorders.
  useEffect(() => {
    if (!slides.length) {
      setSelectedSlideId(null);
      return;
    }
    if (!selectedSlideId || !slides.some(s => s.id === selectedSlideId)) {
      setSelectedSlideId(slides[0].id);
    }
  }, [slides, selectedSlideId]);

  // PUBLIC_INTERFACE
  const toggleTheme = () => {
    setTheme(prev => (prev === 'light' ? 'dark' : 'light'));
  };

  function setSlideField(slideId, patch) {
    setSlides(prev =>
      prev.map(s => {
        if (s.id !== slideId) return s;
        // Cover slide is locked in position, but title can be edited (to render on cover)
        if (s.isCover) {
          const allowed = { title: patch.title };
          return { ...s, ...allowed };
        }
        return { ...s, ...patch };
      })
    );
  }

  function addSlide() {
    const newSlide = {
      id: createId(),
      title: '',
      body: '',
      layout: 'titleBody',
    };
    setSlides(prev => [...prev, newSlide]);
    setSelectedSlideId(newSlide.id);
    setStatus({ type: 'idle', message: '' });

    // Focus title input after render.
    window.setTimeout(() => titleInputRef.current?.focus?.(), 0);
  }

  function duplicateSlide(slideId) {
    const src = slides.find(s => s.id === slideId);
    if (!src) return;
    if (src.isCover) {
      setStatus({ type: 'error', message: 'Cover slide is locked and cannot be duplicated.' });
      return;
    }

    const clone = { ...src, id: createId(), title: src.title ? `${src.title} (Copy)` : '' };
    const idx = slides.findIndex(s => s.id === slideId);
    const insertAt = idx >= 0 ? idx + 1 : slides.length;

    setSlides(prev => {
      const next = [...prev];
      next.splice(insertAt, 0, clone);
      return next;
    });
    setSelectedSlideId(clone.id);
    setStatus({ type: 'idle', message: '' });
  }

  function removeSlide(slideId) {
    const src = slides.find(s => s.id === slideId);
    if (src?.isCover) {
      setStatus({ type: 'error', message: 'Cover slide is locked and cannot be deleted.' });
      return;
    }

    if (slides.length <= 2) {
      // with a required cover, we still need at least 1 additional slide
      setStatus({ type: 'error', message: 'You must keep at least one slide after the cover.' });
      return;
    }

    const idx = slides.findIndex(s => s.id === slideId);
    setSlides(prev => prev.filter(s => s.id !== slideId));

    // Update selection to a neighbor slide.
    const nextIdx = clamp(idx, 0, slides.length - 2);
    const nextId = slides.filter(s => s.id !== slideId)[nextIdx]?.id;
    if (nextId) setSelectedSlideId(nextId);

    setStatus({ type: 'idle', message: '' });
  }

  function moveSlide(slideId, direction) {
    const idx = slides.findIndex(s => s.id === slideId);
    if (idx < 0) return;

    const src = slides[idx];
    if (src?.isCover) return; // locked
    // Also prevent moving another slide above cover
    if (direction === 'up' && idx === 1) return;

    const target = direction === 'up' ? idx - 1 : idx + 1;
    if (target < 0 || target >= slides.length) return;

    setSlides(prev => {
      const next = [...prev];
      const [item] = next.splice(idx, 1);
      next.splice(target, 0, item);
      return next;
    });

    setStatus({ type: 'idle', message: '' });
  }

  function validate() {
    if (!presentationTitle.trim()) {
      return 'Presentation title is required.';
    }
    if (slides.length < 2) {
      return 'Add at least one slide after the cover.';
    }
    // Cover name/date can be empty (optional), but encourage user by not erroring.

    for (let i = 0; i < slides.length; i += 1) {
      const s = slides[i];
      if (!safeText(s.title).trim()) {
        return `Slide ${i + 1} title is required.`;
      }
      if (!s.isCover && safeText(s.layout) !== 'title' && !safeText(s.body).trim()) {
        return `Slide ${i + 1} body text is required for this layout.`;
      }
    }
    return null;
  }

  async function handleCoverImageReplace(file) {
    try {
      if (!file) return;
      if (!/image\/png/i.test(file.type) && !/image\/(jpeg|jpg)/i.test(file.type)) {
        setStatus({ type: 'error', message: 'Please choose a PNG or JPG image.' });
        return;
      }

      const buf = await readFileAsArrayBuffer(file);
      const bytes = new Uint8Array(buf);

      // PPTX part expects PNG. If JPG is provided we still embed as cover.png; Office may reject.
      // For now, we accept PNG primarily; JPG is allowed for preview but may not render in PPT.
      // To keep behavior predictable, hard-reject non-PNG for PPTX embedding.
      if (!/image\/png/i.test(file.type)) {
        setStatus({
          type: 'error',
          message: 'Only PNG is supported for PPTX embedding right now. Please provide a PNG image.',
        });
        return;
      }

      // Update state for both preview and PPTX generation.
      setCoverImageBytes(bytes);

      // Create a local preview URL
      const blobUrl = URL.createObjectURL(new Blob([bytes], { type: 'image/png' }));
      // Revoke old url to avoid leaks
      setCoverImagePreviewUrl(prev => {
        if (prev) URL.revokeObjectURL(prev);
        return blobUrl;
      });

      setStatus({ type: 'success', message: 'Cover image replaced for this session.' });
    } catch (e) {
      setStatus({
        type: 'error',
        message: `Failed to replace cover image. ${e?.message ? `(${e.message})` : ''}`.trim(),
      });
    }
  }

  async function handleGenerateAndDownload() {
    setStatus({ type: 'idle', message: '' });

    const error = validate();
    if (error) {
      setStatus({ type: 'error', message: error });
      return;
    }

    try {
      setIsGenerating(true);
      setStatus({ type: 'working', message: 'Generating PPTX…' });

      const bytes = buildPptxBytes({
        slides,
        title: presentationTitle,
        cover: {
          name: coverName,
          date: coverDate,
          imageBytes: coverImageBytes || new Uint8Array(),
        },
      });

      const fileNameBase = (presentationTitle || 'presentation')
        .trim()
        .slice(0, 80)
        .replace(/[^\w\- ]+/g, '')
        .replace(/\s+/g, '-')
        .toLowerCase();

      downloadBytes(bytes, `${fileNameBase || 'presentation'}.pptx`);

      setStatus({ type: 'success', message: 'PPTX generated. Your download should start automatically.' });
    } catch (e) {
      setStatus({
        type: 'error',
        message: `Failed to generate PPTX. ${e?.message ? `(${e.message})` : ''}`.trim(),
      });
    } finally {
      setIsGenerating(false);
    }
  }

  const cover = useMemo(
    () => ({
      name: coverName,
      date: coverDate,
      imageBytes: coverImageBytes,
      imagePreviewUrl: coverImagePreviewUrl,
      imageDefaultAssetPath: '/assets/cover-slide.png',
    }),
    [coverName, coverDate, coverImageBytes, coverImagePreviewUrl]
  );

  return (
    <div className="appShell">
      <header className="topbar">
        <div className="topbar__left">
          <div className="brandMark" aria-hidden="true">
            <span className="brandMark__dot" />
          </div>
          <div>
            <div className="topbar__title">PPT Generator</div>
            <div className="topbar__subtitle">Create, preview, and download .pptx — 100% client-side.</div>
          </div>
        </div>

        <div className="topbar__right">
          <button
            type="button"
            className="btn btn--ghost"
            onClick={toggleTheme}
            aria-label={`Switch to ${theme === 'light' ? 'dark' : 'light'} mode`}
            title="Toggle theme"
          >
            {theme === 'light' ? 'Dark mode' : 'Light mode'}
          </button>

          <button
            type="button"
            className="btn btn--primary"
            onClick={handleGenerateAndDownload}
            disabled={isGenerating}
          >
            {isGenerating ? 'Generating…' : 'Generate & Download'}
          </button>
        </div>
      </header>

      <main className="layout">
        <section className="panel panel--editor" aria-label="Slide editor">
          <div className="panel__header">
            <div className="panel__heading">
              <h2 className="panel__title">Editor</h2>
              <p className="panel__desc">Manage slides, edit content, and choose a simple layout.</p>
            </div>
            <button type="button" className="btn btn--accent" onClick={addSlide}>
              + Add slide
            </button>
          </div>

          <div className="panel__content">
            <div className="field">
              <label className="field__label" htmlFor="deckTitle">
                Presentation title <span className="field__req">*</span>
              </label>
              <input
                id="deckTitle"
                className="input"
                value={presentationTitle}
                onChange={e => setPresentationTitle(e.target.value)}
                placeholder="e.g., Q4 Strategy Review"
              />
            </div>

            <div className="divider" />

            <div className="editorGrid">
              <div className="slideList">
                <div className="slideList__header">
                  <div className="slideList__title">Slides</div>
                  <div className="slideList__count">{slides.length}</div>
                </div>

                <div className="slideList__items" role="list">
                  {slides.map((s, idx) => {
                    const isSelected = s.id === selectedSlideId;
                    const isLockedCover = Boolean(s.isCover);

                    return (
                      <div key={s.id} className={`slideRow ${isSelected ? 'slideRow--selected' : ''}`} role="listitem">
                        <button
                          type="button"
                          className="slideRow__select"
                          onClick={() => setSelectedSlideId(s.id)}
                          aria-label={`Select slide ${idx + 1}`}
                        >
                          <span className="slideRow__index">{idx + 1}</span>
                          <span className="slideRow__text">
                            <span className="slideRow__title">
                              {safeText(s.title).trim() ? safeText(s.title).trim() : 'Untitled'}
                              {isLockedCover ? ' (Locked Cover)' : ''}
                            </span>
                            <span className="slideRow__meta">
                              {isLockedCover ? 'Global cover (image + name/date)' : LAYOUTS.find(l => l.value === s.layout)?.label || 'Layout'}
                            </span>
                          </span>
                        </button>

                        <div className="slideRow__actions">
                          <button
                            type="button"
                            className="iconBtn"
                            onClick={() => moveSlide(s.id, 'up')}
                            disabled={isLockedCover || idx <= 1}
                            aria-label="Move slide up"
                            title={isLockedCover ? 'Cover is locked' : idx <= 1 ? 'Cannot move above cover' : 'Move up'}
                          >
                            ↑
                          </button>
                          <button
                            type="button"
                            className="iconBtn"
                            onClick={() => moveSlide(s.id, 'down')}
                            disabled={isLockedCover || idx === slides.length - 1}
                            aria-label="Move slide down"
                            title={isLockedCover ? 'Cover is locked' : 'Move down'}
                          >
                            ↓
                          </button>
                          <button
                            type="button"
                            className="iconBtn"
                            onClick={() => duplicateSlide(s.id)}
                            disabled={isLockedCover}
                            aria-label="Duplicate slide"
                            title={isLockedCover ? 'Cover is locked' : 'Duplicate'}
                          >
                            ⧉
                          </button>
                          <button
                            type="button"
                            className="iconBtn iconBtn--danger"
                            onClick={() => removeSlide(s.id)}
                            disabled={isLockedCover}
                            aria-label="Delete slide"
                            title={isLockedCover ? 'Cover is locked' : 'Delete'}
                          >
                            ✕
                          </button>
                        </div>
                      </div>
                    );
                  })}
                </div>
              </div>

              <div className="slideEditor">
                <div className="slideEditor__header">
                  <div>
                    <div className="slideEditor__kicker">Selected</div>
                    <div className="slideEditor__title">
                      Slide {selectedSlideIndex + 1} of {slides.length}
                    </div>
                  </div>
                </div>

                {selectedSlide ? (
                  <div className="slideEditor__form">
                    {selectedSlide.isCover ? (
                      <>
                        <div className="notice notice--info" role="status" style={{ marginTop: 0 }}>
                          <div className="notice__title">Cover slide</div>
                          <div className="notice__message">
                            Slide 1 is a locked global cover. You can edit Name/Date and the cover title, and you can
                            replace the background image (position remains locked).
                          </div>
                        </div>

                        <div className="divider" />

                        <div className="field">
                          <label className="field__label" htmlFor="coverTitle">
                            Cover title <span className="field__req">*</span>
                          </label>
                          <input
                            id="coverTitle"
                            ref={titleInputRef}
                            className="input"
                            value={selectedSlide.title}
                            onChange={e => setSlideField(selectedSlide.id, { title: e.target.value })}
                            placeholder="e.g., Weekly Metrics"
                          />
                          {!safeText(selectedSlide.title).trim() ? (
                            <div className="field__help field__help--error">Title is required.</div>
                          ) : (
                            <div className="field__help">Shown on the cover overlay and embedded in the PPTX.</div>
                          )}
                        </div>

                        <div className="field">
                          <label className="field__label" htmlFor="coverName">
                            Name (editable)
                          </label>
                          <input
                            id="coverName"
                            className="input"
                            value={coverName}
                            onChange={e => setCoverName(e.target.value)}
                            placeholder="e.g., Jane Doe"
                          />
                          <div className="field__help">Renders on top of the cover image.</div>
                        </div>

                        <div className="field">
                          <label className="field__label" htmlFor="coverDate">
                            Date (editable)
                          </label>
                          <input
                            id="coverDate"
                            className="input"
                            value={coverDate}
                            onChange={e => setCoverDate(e.target.value)}
                            placeholder="e.g., 01 Jan 2026"
                          />
                          <div className="field__help">Renders on top of the cover image.</div>
                        </div>

                        <div className="field">
                          <label className="field__label" htmlFor="coverImage">
                            Cover image (background)
                          </label>
                          <input
                            id="coverImage"
                            className="input"
                            type="file"
                            accept="image/png"
                            onChange={e => handleCoverImageReplace(e.target.files?.[0] || null)}
                          />
                          <div className="field__help">
                            Loaded from <code>{cover.imageDefaultAssetPath}</code> by default. You can replace it with a PNG
                            (stored for this session). Position remains locked.
                          </div>
                        </div>
                      </>
                    ) : (
                      <>
                        <div className="field">
                          <label className="field__label" htmlFor="slideTitle">
                            Slide title <span className="field__req">*</span>
                          </label>
                          <input
                            id="slideTitle"
                            ref={titleInputRef}
                            className="input"
                            value={selectedSlide.title}
                            onChange={e => setSlideField(selectedSlide.id, { title: e.target.value })}
                            placeholder="e.g., Agenda"
                          />
                          {!safeText(selectedSlide.title).trim() ? (
                            <div className="field__help field__help--error">Title is required.</div>
                          ) : (
                            <div className="field__help">Used as the main slide heading.</div>
                          )}
                        </div>

                        <div className="field">
                          <label className="field__label" htmlFor="slideLayout">
                            Layout
                          </label>
                          <select
                            id="slideLayout"
                            className="select"
                            value={selectedSlide.layout}
                            onChange={e => setSlideField(selectedSlide.id, { layout: e.target.value })}
                          >
                            {LAYOUTS.map(l => (
                              <option key={l.value} value={l.value}>
                                {l.label}
                              </option>
                            ))}
                          </select>
                          <div className="field__help">A simple layout choice affects PPTX formatting.</div>
                        </div>

                        <div className="field">
                          <label className="field__label" htmlFor="slideBody">
                            Body text {selectedSlide.layout === 'title' ? '(optional)' : <span className="field__req">*</span>}
                          </label>
                          <textarea
                            id="slideBody"
                            className="textarea"
                            value={selectedSlide.body}
                            onChange={e => setSlideField(selectedSlide.id, { body: e.target.value })}
                            placeholder={
                              selectedSlide.layout === 'twoColumn'
                                ? 'Use new lines to create bullet points.\nWe will auto-split between columns.'
                                : 'Use new lines for bullet points.'
                            }
                            rows={selectedSlide.layout === 'title' ? 4 : 8}
                          />
                          {selectedSlide.layout !== 'title' && !safeText(selectedSlide.body).trim() ? (
                            <div className="field__help field__help--error">Body text is required for this layout.</div>
                          ) : (
                            <div className="field__help">Tip: use new lines to create bullet-like rows.</div>
                          )}
                        </div>
                      </>
                    )}
                  </div>
                ) : (
                  <div className="emptyState">
                    <div className="emptyState__title">No slide selected</div>
                    <div className="emptyState__desc">Add a slide to start editing.</div>
                  </div>
                )}
              </div>
            </div>

            {status.type !== 'idle' ? (
              <div
                className={`notice ${
                  status.type === 'error'
                    ? 'notice--error'
                    : status.type === 'success'
                      ? 'notice--success'
                      : 'notice--info'
                }`}
                role={status.type === 'error' ? 'alert' : 'status'}
              >
                <div className="notice__title">
                  {status.type === 'error'
                    ? 'Validation / Error'
                    : status.type === 'success'
                      ? 'Success'
                      : 'Working'}
                </div>
                <div className="notice__message">{status.message}</div>
              </div>
            ) : null}
          </div>
        </section>

        <section className="panel panel--preview" aria-label="Presentation preview">
          <div className="panel__header">
            <div className="panel__heading">
              <h2 className="panel__title">Preview</h2>
              <p className="panel__desc">Thumbnails and selected-slide preview.</p>
            </div>
          </div>

          <div className="panel__content panel__content--preview">
            <div className="thumbGrid" aria-label="Slide thumbnails">
              {slides.map((s, idx) => (
                <SlideThumbnail
                  key={s.id}
                  slide={s}
                  index={idx}
                  isSelected={s.id === selectedSlideId}
                  onSelect={() => setSelectedSlideId(s.id)}
                />
              ))}
            </div>

            {selectedSlide ? <SlidePreview slide={selectedSlide} cover={cover} /> : null}
          </div>
        </section>
      </main>

      <footer className="footer">
        <div className="footer__left">
          <span className="pill">Ocean Professional</span>
          <span className="footer__text">No backend. Everything runs in your browser.</span>
        </div>
        <div className="footer__right">
          <span className="footer__text">Format: PPTX (Open XML)</span>
        </div>
      </footer>
    </div>
  );
}

export default App;
