/**
 * Unit test for ChatErrorBoundary.
 *
 * Verifies:
 *   - Renders children normally when no error occurs
 *   - Shows fallback UI when a child throws during render
 *   - Displays the error message in the fallback
 *   - "Try again" button recovers (resets error state and re-renders children)
 *   - Logs the error to console.error
 */

import React, { useState } from 'react';
import { describe, it, expect, vi, beforeEach } from 'vitest';
import { screen } from '@testing-library/react';
import userEvent from '@testing-library/user-event';
import { renderWithProviders } from '../test-utils';
import { ChatErrorBoundary } from '@/components/ChatErrorBoundary';

// ─── Test helpers ───

/** A component that throws on render when `shouldThrow` is true. */
const ThrowingChild: React.FC<{ shouldThrow: boolean }> = ({ shouldThrow }) => {
  if (shouldThrow) {
    throw new Error('Simulated render crash');
  }
  return <div data-testid="child-content">Chat works fine</div>;
};

/** A wrapper that lets us toggle the throw state via a button outside the boundary. */
const TestHarness: React.FC<{ initialThrow?: boolean }> = ({ initialThrow = false }) => {
  const [shouldThrow, setShouldThrow] = useState(initialThrow);
  return (
    <div>
      <button data-testid="toggle-throw" onClick={() => setShouldThrow(prev => !prev)}>
        Toggle
      </button>
      <ChatErrorBoundary>
        <ThrowingChild shouldThrow={shouldThrow} />
      </ChatErrorBoundary>
    </div>
  );
};

// ─── Tests ───

describe('ChatErrorBoundary', () => {
  beforeEach(() => {
    // Suppress React's console.error for error boundary — it's expected
    vi.spyOn(console, 'error').mockImplementation(() => {});
  });

  it('renders children normally when no error occurs', () => {
    renderWithProviders(
      <ChatErrorBoundary>
        <div data-testid="child">Hello</div>
      </ChatErrorBoundary>
    );

    expect(screen.getByTestId('child')).toBeInTheDocument();
    expect(screen.getByText('Hello')).toBeInTheDocument();
  });

  it('shows fallback UI when a child throws during render', () => {
    renderWithProviders(
      <ChatErrorBoundary>
        <ThrowingChild shouldThrow={true} />
      </ChatErrorBoundary>
    );

    expect(screen.queryByTestId('child-content')).not.toBeInTheDocument();
    expect(screen.getByText('Something went wrong')).toBeInTheDocument();
    expect(screen.getByText('Simulated render crash')).toBeInTheDocument();
    expect(screen.getByText('Try again')).toBeInTheDocument();
  });

  it('logs the error via console.error', () => {
    renderWithProviders(
      <ChatErrorBoundary>
        <ThrowingChild shouldThrow={true} />
      </ChatErrorBoundary>
    );

    expect(console.error).toHaveBeenCalled();
  });

  it('"Try again" button resets the error boundary', async () => {
    const user = userEvent.setup();

    // Use TestHarness — starts throwing, but we can toggle it off
    renderWithProviders(<TestHarness initialThrow={true} />);

    // Fallback is shown
    expect(screen.getByText('Something went wrong')).toBeInTheDocument();

    // Toggle the throw state OFF (so re-render succeeds)
    await user.click(screen.getByTestId('toggle-throw'));

    // Click "Try again" in the error boundary
    await user.click(screen.getByText('Try again'));

    // Child should render successfully now
    expect(screen.getByTestId('child-content')).toBeInTheDocument();
    expect(screen.getByText('Chat works fine')).toBeInTheDocument();
    expect(screen.queryByText('Something went wrong')).not.toBeInTheDocument();
  });

  it('shows generic message when error has no message', () => {
    const BadChild = () => {
      throw new Error('');
    };

    renderWithProviders(
      <ChatErrorBoundary>
        <BadChild />
      </ChatErrorBoundary>
    );

    expect(screen.getByText('Something went wrong')).toBeInTheDocument();
    // Should show fallback text for empty error message
    expect(screen.getByText('An unexpected error occurred.')).toBeInTheDocument();
  });
});
