/**
 * Integration test: App error boundary keeps header alive on chat crash.
 *
 * Verifies that when ChatPanel (or its children) throw during render,
 * the ChatErrorBoundary catches the error and:
 *   - Shows the fallback UI
 *   - Keeps the ChatHeader and settings accessible
 *   - Allows recovery via "Try again"
 */

import React from 'react';
import { describe, it, expect, beforeEach, vi } from 'vitest';
import { screen, waitFor } from '@testing-library/react';
import userEvent from '@testing-library/user-event';
import { render } from '@testing-library/react';
import { useSettingsStore } from '@/stores/settingsStore';

// ─── Track whether ChatPanel should crash ───
let chatPanelShouldCrash = false;

// ─── Mock components ───

vi.mock('@/components/ChatHeader', () => ({
  ChatHeader: ({
    onClearMessages,
  }: {
    onClearMessages: () => void;
    settingsOpen: boolean;
    onSettingsOpenChange: (open: boolean) => void;
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

vi.mock('@/components/SetupWizard', () => ({
  SetupWizard: ({ onComplete }: { onComplete: () => void }) =>
    React.createElement('div', { 'data-testid': 'setup-wizard' }, [
      React.createElement('button', { key: 'btn', onClick: onComplete }, 'Complete'),
    ]),
}));

vi.mock('@assistant-ui/react', () => ({
  AssistantRuntimeProvider: ({ children }: { children: React.ReactNode }) =>
    React.createElement('div', null, children),
}));

vi.mock('@assistant-ui/react-ai-sdk', () => ({
  useAISDKRuntime: () => ({}),
}));

vi.mock('@/services/ai/aiClientFactory', () => ({
  getProviderModel: vi.fn(() => ({})),
}));

vi.mock('@/hooks/useOfficeChat', () => ({
  useOfficeChat: () => ({
    messages: [],
    sendMessage: vi.fn(),
    stop: vi.fn(),
    status: 'ready',
    setMessages: vi.fn(),
    error: undefined,
    clearError: vi.fn(),
    id: 'test',
  }),
}));

const { App } = await import('@/taskpane/App');

// ─── Helpers ───

function configureReadyState() {
  const epId = useSettingsStore.getState().addEndpoint({
    displayName: 'Test',
    resourceUrl: 'https://test.openai.azure.com',
    authMethod: 'apiKey',
    apiKey: 'key',
  });
  useSettingsStore
    .getState()
    .setModelsForEndpoint(epId, [
      { id: 'gpt-4.1', name: 'gpt-4.1', ownedBy: 'user', provider: 'OpenAI' },
    ]);
}

// ─── Tests ───

describe('App — error boundary integration', () => {
  beforeEach(() => {
    useSettingsStore.getState().reset();
    chatPanelShouldCrash = false;
    vi.spyOn(console, 'error').mockImplementation(() => {});
  });

  it('shows chat normally when ChatPanel does not crash', async () => {
    configureReadyState();
    render(<App />);

    await waitFor(() => {
      expect(screen.getByTestId('chat-header')).toBeInTheDocument();
      expect(screen.getByTestId('chat-panel')).toBeInTheDocument();
      expect(screen.getByText('Chat works')).toBeInTheDocument();
    });
  });

  it('keeps ChatHeader alive when ChatPanel crashes', async () => {
    configureReadyState();
    chatPanelShouldCrash = true;
    render(<App />);

    await waitFor(() => {
      // Header should still be visible
      expect(screen.getByTestId('chat-header')).toBeInTheDocument();
      expect(screen.getByText('New conversation')).toBeInTheDocument();

      // Chat panel should NOT be visible
      expect(screen.queryByTestId('chat-panel')).not.toBeInTheDocument();

      // Error boundary fallback should be shown
      expect(screen.getByText('Something went wrong')).toBeInTheDocument();
      expect(screen.getByText('Thread render failed')).toBeInTheDocument();
    });
  });

  it('"Try again" recovers when the crash is resolved', async () => {
    const user = userEvent.setup();
    configureReadyState();
    chatPanelShouldCrash = true;
    render(<App />);

    await waitFor(() => {
      expect(screen.getByText('Something went wrong')).toBeInTheDocument();
    });

    // Fix the crash
    chatPanelShouldCrash = false;

    // Click "Try again"
    await user.click(screen.getByText('Try again'));

    // Chat should recover
    await waitFor(() => {
      expect(screen.getByTestId('chat-panel')).toBeInTheDocument();
      expect(screen.getByText('Chat works')).toBeInTheDocument();
      expect(screen.queryByText('Something went wrong')).not.toBeInTheDocument();
    });
  });
});
