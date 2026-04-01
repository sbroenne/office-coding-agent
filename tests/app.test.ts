import React from 'react';
import { describe, it, expect, beforeEach, vi } from 'vitest';
import { screen, waitFor } from '@testing-library/react';
import { renderWithProviders } from './test-utils';
import { useSettingsStore } from '@/stores/settingsStore';

vi.mock('@/components/ChatHeader', () => ({
  ChatHeader: () => React.createElement('div', { 'data-testid': 'chat-header' }, 'ChatHeader'),
}));
vi.mock('@/components/ChatPanel', () => ({
  ChatPanel: () => React.createElement('div', { 'data-testid': 'chat-panel' }, 'ChatPanel'),
}));

// Prevent useOfficeChat from opening a real WebSocket connection in tests.
// Without this mock the hook tries to connect to ws://localhost:3000 which
// fails, producing console logs that arrive after test teardown and cause
// EnvironmentTeardownError: Closing rpc while "onUserConsoleLog" was pending.
vi.mock('@/hooks/useOfficeChat', () => ({
  useOfficeChat: () => ({
    messages: [],
    isRunning: false,
    send: vi.fn(),
    cancel: vi.fn(),
    sessionError: null,
    isConnecting: false,
    clearMessages: vi.fn(),
    restoreSession: vi.fn(),
    deleteSession: vi.fn(),
    sessions: [],
    activeSessionId: null,
    pendingPermission: null,
    allowAllPermissions: vi.fn(),
    approvePermission: vi.fn(),
    denyPermission: vi.fn(),
    allowPermissionAlways: vi.fn(),
    compactSession: vi.fn(),
    switchModel: vi.fn(),
    enqueue: vi.fn(),
    queuedPrompts: [],
    dequeue: vi.fn(),
    clearQueue: vi.fn(),
  }),
}));

const { App } = await import('@/taskpane/App');

describe('App', () => {
  beforeEach(() => {
    useSettingsStore.getState().reset();
  });

  it('renders without crashing', () => {
    renderWithProviders(React.createElement(App));
    expect(document.body.querySelector('div')).not.toBeNull();
  });

  it('shows chat UI after hydration', async () => {
    renderWithProviders(React.createElement(App));
    await waitFor(() => {
      expect(screen.getByTestId('chat-header')).toBeInTheDocument();
      expect(screen.getByTestId('chat-panel')).toBeInTheDocument();
    });
  });
});
