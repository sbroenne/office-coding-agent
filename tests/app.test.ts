import React from 'react';
import { describe, it, expect, beforeEach, vi } from 'vitest';
import { screen, waitFor, act } from '@testing-library/react';
import userEvent from '@testing-library/user-event';
import { renderWithProviders } from './test-utils';
import { useSettingsStore } from '@/stores/settingsStore';

// ─── Mock child components to isolate App logic ───
vi.mock('@/components/SetupWizard', () => ({
  SetupWizard: ({ onComplete }: { onComplete: () => void }) =>
    React.createElement(
      'div',
      { 'data-testid': 'setup-wizard' },
      React.createElement(
        'button',
        { 'data-testid': 'setup-complete-btn', onClick: onComplete },
        'Complete Setup'
      )
    ),
}));
vi.mock('@/components/ChatHeader', () => ({
  ChatHeader: () => React.createElement('div', { 'data-testid': 'chat-header' }, 'ChatHeader'),
}));
vi.mock('@/components/ChatPanel', () => ({
  ChatPanel: () => React.createElement('div', { 'data-testid': 'chat-panel' }, 'ChatPanel'),
}));

// Import App AFTER mocks are registered
const { App } = await import('@/taskpane/App');

describe('App', () => {
  beforeEach(() => {
    useSettingsStore.getState().reset();
  });

  it('shows setup wizard when no endpoints exist', async () => {
    renderWithProviders(React.createElement(App));

    await waitFor(() => {
      expect(screen.getByTestId('setup-wizard')).toBeInTheDocument();
    });
  });

  it('shows setup wizard when endpoint exists but has no models', async () => {
    // Add an endpoint WITHOUT any models (the exact bug scenario)
    useSettingsStore.getState().addEndpoint({
      displayName: 'Test',
      resourceUrl: 'https://test.openai.azure.com',
      authMethod: 'apiKey',
      apiKey: 'key',
    });

    renderWithProviders(React.createElement(App));

    await waitFor(() => {
      expect(screen.getByTestId('setup-wizard')).toBeInTheDocument();
    });
  });

  it('shows chat UI when endpoint exists with models', async () => {
    const epId = useSettingsStore.getState().addEndpoint({
      displayName: 'Test',
      resourceUrl: 'https://test.openai.azure.com',
      authMethod: 'apiKey',
      apiKey: 'key',
    });
    useSettingsStore
      .getState()
      .setModelsForEndpoint(epId, [
        { id: 'gpt-5.2-chat', name: 'gpt-5.2-chat', ownedBy: 'user', provider: 'OpenAI' },
      ]);

    renderWithProviders(React.createElement(App));

    await waitFor(() => {
      expect(screen.getByTestId('chat-header')).toBeInTheDocument();
      expect(screen.getByTestId('chat-panel')).toBeInTheDocument();
    });
  });

  it('transitions from ready → setup when all endpoints are removed', async () => {
    // Start with a working configuration
    const epId = useSettingsStore.getState().addEndpoint({
      displayName: 'Test',
      resourceUrl: 'https://test.openai.azure.com',
      authMethod: 'apiKey',
      apiKey: 'key',
    });
    useSettingsStore
      .getState()
      .setModelsForEndpoint(epId, [
        { id: 'gpt-5.2-chat', name: 'gpt-5.2-chat', ownedBy: 'user', provider: 'OpenAI' },
      ]);

    renderWithProviders(React.createElement(App));

    // Verify chat UI is showing
    await waitFor(() => {
      expect(screen.getByTestId('chat-header')).toBeInTheDocument();
    });

    // Remove all endpoints — should trigger transition to setup
    act(() => {
      useSettingsStore.getState().removeEndpoint(epId);
    });

    await waitFor(() => {
      expect(screen.getByTestId('setup-wizard')).toBeInTheDocument();
    });
  });

  it('transitions from setup → ready when wizard completes', async () => {
    const user = userEvent.setup();

    // Start with empty store → shows wizard
    renderWithProviders(React.createElement(App));

    await waitFor(() => {
      expect(screen.getByTestId('setup-wizard')).toBeInTheDocument();
    });

    // Simulate completing setup by adding valid config before clicking Complete
    act(() => {
      const epId = useSettingsStore.getState().addEndpoint({
        displayName: 'Test',
        resourceUrl: 'https://test.openai.azure.com',
        authMethod: 'apiKey',
        apiKey: 'key',
      });
      useSettingsStore
        .getState()
        .setModelsForEndpoint(epId, [
          { id: 'gpt-5.2-chat', name: 'gpt-5.2-chat', ownedBy: 'user', provider: 'OpenAI' },
        ]);
    });

    // Click the complete button on the mocked wizard
    await user.click(screen.getByTestId('setup-complete-btn'));

    await waitFor(() => {
      expect(screen.getByTestId('chat-header')).toBeInTheDocument();
      expect(screen.getByTestId('chat-panel')).toBeInTheDocument();
    });
  });
});
