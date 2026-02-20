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
