/**
 * Integration test: App error boundary keeps header alive on chat crash.
 */

import React from 'react';
import { describe, it, expect, beforeEach, vi } from 'vitest';
import { screen, waitFor } from '@testing-library/react';
import userEvent from '@testing-library/user-event';
import { render } from '@testing-library/react';
import { useSettingsStore } from '@/stores/settingsStore';

let chatPanelShouldCrash = false;

vi.mock('@/components/ChatHeader', () => ({
  ChatHeader: ({
    onClearMessages,
  }: {
    onClearMessages: () => void;
  }) =>
    React.createElement(
      'div',
      { 'data-testid': 'chat-header' },
      React.createElement('button', { 'data-testid': 'clear-btn', onClick: onClearMessages }, 'New conversation')
    ),
}));

vi.mock('@/components/ChatPanel', () => ({
  ChatPanel: () => {
    if (chatPanelShouldCrash) {
      throw new Error('Thread render failed');
    }
    return React.createElement('div', { 'data-testid': 'chat-panel' }, 'Chat works');
  },
}));

const { App } = await import('@/taskpane/App');

describe('App — error boundary integration', () => {
  beforeEach(() => {
    useSettingsStore.getState().reset();
    chatPanelShouldCrash = false;
    vi.spyOn(console, 'error').mockImplementation(() => {});
  });

  it('shows chat normally when ChatPanel does not crash', async () => {
    render(<App />);

    await waitFor(() => {
      expect(screen.getByTestId('chat-header')).toBeInTheDocument();
      expect(screen.getByTestId('chat-panel')).toBeInTheDocument();
      expect(screen.getByText('Chat works')).toBeInTheDocument();
    });
  });

  it('keeps ChatHeader alive when ChatPanel crashes', async () => {
    chatPanelShouldCrash = true;
    render(<App />);

    await waitFor(() => {
      expect(screen.getByTestId('chat-header')).toBeInTheDocument();
      expect(screen.getByText('New conversation')).toBeInTheDocument();
      expect(screen.queryByTestId('chat-panel')).not.toBeInTheDocument();
      expect(screen.getByText('Something went wrong')).toBeInTheDocument();
      expect(screen.getByText('Thread render failed')).toBeInTheDocument();
    });
  });

  it('"Try again" recovers when the crash is resolved', async () => {
    const user = userEvent.setup();
    chatPanelShouldCrash = true;
    render(<App />);

    await waitFor(() => {
      expect(screen.getByText('Something went wrong')).toBeInTheDocument();
    });

    chatPanelShouldCrash = false;
    await user.click(screen.getByText('Try again'));

    await waitFor(() => {
      expect(screen.getByTestId('chat-panel')).toBeInTheDocument();
      expect(screen.getByText('Chat works')).toBeInTheDocument();
      expect(screen.queryByText('Something went wrong')).not.toBeInTheDocument();
    });
  });
});
