import React from 'react';
import { render, type RenderOptions } from '@testing-library/react';

/**
 * Custom render function that wraps components in any required providers.
 */
export function renderWithProviders(
  ui: React.ReactElement,
  options?: Omit<RenderOptions, 'wrapper'>
) {
  const Wrapper: React.FC<{ children: React.ReactNode }> = ({ children }) => <>{children}</>;

  return render(ui, { wrapper: Wrapper, ...options });
}
