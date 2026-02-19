/**
 * Integration test for App component state transitions and theme detection.
 *
 * Tests the App component's routing between loading → setup → ready states,
 * theme detection (dark/light class toggling), and configuration-driven
 * state transitions.
 */
import React from 'react';
import { describe, it, expect, beforeEach, vi, afterEach } from 'vitest';
import { screen, waitFor } from '@testing-library/react';
import { render } from '@testing-library/react';
import { useSettingsStore } from '@/stores/settingsStore';

// ─── Mock heavy child components ───
// Keeps tests fast and avoids needing a full assistant-ui runtime.

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
      React.createElement(
        'button',
        { 'data-testid': 'clear-btn', onClick: onClearMessages },
        'Clear'
      )
    ),
}));

vi.mock('@/components/ChatPanel', () => ({
  ChatPanel: ({ isConfigured }: { isConfigured: boolean }) =>
    React.createElement('div', {
      'data-testid': 'chat-panel',
      'data-configured': String(isConfigured),
    }),
}));

vi.mock('@/components/SetupWizard', () => ({
  SetupWizard: ({ onComplete }: { onComplete: () => void }) =>
    React.createElement('div', { 'data-testid': 'setup-wizard' }, [
      React.createElement('span', { key: 'title' }, 'Connect to Azure AI Foundry'),
      React.createElement('button', { key: 'btn', onClick: onComplete }, 'Complete'),
    ]),
}));

// ─── Mock assistant-ui runtime ───

vi.mock('@assistant-ui/react', () => ({
  AssistantRuntimeProvider: ({ children }: { children: React.ReactNode }) =>
    React.createElement('div', { 'data-testid': 'runtime-provider' }, children),
}));

vi.mock('@assistant-ui/react-ai-sdk', () => ({
  useAISDKRuntime: () => ({}),
}));

// ─── Mock AI services ───

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

// ─── Tests ───

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

  it('shows setup wizard when no endpoints are configured', async () => {
    render(<App />);
    await waitFor(() => {
      expect(screen.getByText('Connect to Azure AI Foundry')).toBeInTheDocument();
    });
  });

  it('shows chat UI when endpoints and models are configured', async () => {
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

  it('reverts to setup wizard if config is wiped while in ready state', async () => {
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
    render(<App />);
    await waitFor(() => {
      expect(screen.getByTestId('chat-header')).toBeInTheDocument();
    });
    useSettingsStore.getState().reset();
    await waitFor(() => {
      expect(screen.getByText('Connect to Azure AI Foundry')).toBeInTheDocument();
    });
  });
});
