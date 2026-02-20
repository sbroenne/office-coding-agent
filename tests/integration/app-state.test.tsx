/**
 * Integration test for App component state transitions and theme detection.
 */
import React from 'react';
import { describe, it, expect, beforeEach, vi, afterEach } from 'vitest';
import { screen, waitFor } from '@testing-library/react';
import { render } from '@testing-library/react';
import { useSettingsStore } from '@/stores/settingsStore';

vi.mock('@/components/ChatHeader', () => ({
  ChatHeader: ({
    onClearMessages,
  }: {
    onClearMessages: () => void;
  }) =>
    React.createElement(
      'div',
      { 'data-testid': 'chat-header' },
      React.createElement(
        'button',
        { 'data-testid': 'clear-btn', onClick: onClearMessages },
        'Clear'
      )
    ),
}));

vi.mock('@/components/ChatPanel', () => ({
  ChatPanel: () =>
    React.createElement('div', {
      'data-testid': 'chat-panel',
    }),
}));

const { App } = await import('@/taskpane/App');

describe('App — state transitions and theme', () => {
  let originalMatchMedia: typeof window.matchMedia;

  beforeEach(() => {
    useSettingsStore.getState().reset();
    originalMatchMedia = window.matchMedia;
  });

  afterEach(() => {
    window.matchMedia = originalMatchMedia;
    document.documentElement.classList.remove('dark');
  });

  function mockMatchMedia(prefersDark: boolean) {
    window.matchMedia = vi.fn().mockImplementation((query: string) => ({
      matches: query === '(prefers-color-scheme: dark)' ? prefersDark : false,
      media: query,
      onchange: null,
      addListener: vi.fn(),
      removeListener: vi.fn(),
      addEventListener: vi.fn(),
      removeEventListener: vi.fn(),
      dispatchEvent: vi.fn(() => false),
    }));
  }

  it('renders without crashing', () => {
    render(<App />);
    expect(document.body.querySelector('div')).not.toBeNull();
  });

  it('shows chat UI after hydration', async () => {
    render(<App />);
    await waitFor(() => {
      expect(screen.getByTestId('chat-header')).toBeInTheDocument();
      expect(screen.getByTestId('chat-panel')).toBeInTheDocument();
    });
  });

  it('adds .dark class to documentElement when dark mode is preferred', async () => {
    mockMatchMedia(true);
    render(<App />);
    await waitFor(() => {
      expect(document.documentElement.classList.contains('dark')).toBe(true);
    });
  });

  it('does not add .dark class when light mode is preferred', async () => {
    mockMatchMedia(false);
    render(<App />);
    await waitFor(() => {
      expect(document.documentElement.classList.contains('dark')).toBe(false);
    });
  });
});
