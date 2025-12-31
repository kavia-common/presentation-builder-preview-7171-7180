import { render, screen } from '@testing-library/react';
import App from './App';

test('renders PPT Generator header', () => {
  render(<App />);
  expect(screen.getByText(/PPT Generator/i)).toBeInTheDocument();
});
