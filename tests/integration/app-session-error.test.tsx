/**
 * Integration tests for SessionErrorBanner rendering in App.
 *
 * These tests mock createWebSocketClient to always reject, ensuring the
 * SessionErrorBanner is rendered deterministically without relying on test
 * environment networking behaviour.
 *
 * This is the Vitest-level regression guard for the fix in App.tsx that was
 * missing: useOfficeChat returned sessionError but App never rendered it.
 */

import React from 'react';
import { describe, it, expect, vi, beforeEach } from 'vitest';
import { render, screen, waitFor } from '@testing-library/react';
import { useSettingsStore } from '@/stores/settingsStore';

// Must be at module level (hoisted) — mock websocket before App is imported
vi.mock('@/lib/websocket-client', () => ({
  createWebSocketClient: vi.fn().mockRejectedValue(new Error('server unavailable')),
}));

vi.mock('@/components/ChatHeader', () => ({
  ChatHeader: () => React.createElement('div', { 'data-testid': 'chat-header' }),
}));

vi.mock('@/components/ChatPanel', () => ({
  ChatPanel: () => React.createElement('div', { 'data-testid': 'chat-panel' }),
}));

const { App } = await import('@/taskpane/App');

describe('App — SessionErrorBanner', () => {
  beforeEach(() => {
    useSettingsStore.getState().reset();
  });

  it('renders the session error banner when the WebSocket connection fails', async () => {
    render(<App />);
    await waitFor(() => {
      expect(screen.getByText(/Connection failed:/)).toBeInTheDocument();
    });
  });

  it('renders the Retry button inside the session error banner', async () => {
    render(<App />);
    await waitFor(() => {
      expect(screen.getByRole('button', { name: 'Retry' })).toBeInTheDocument();
    });
  });

  it('still renders ChatHeader and ChatPanel alongside the error banner', async () => {
    render(<App />);
    await waitFor(() => {
      expect(screen.getByTestId('chat-header')).toBeInTheDocument();
      expect(screen.getByTestId('chat-panel')).toBeInTheDocument();
      expect(screen.getByText(/Connection failed:/)).toBeInTheDocument();
    });
  });
});
