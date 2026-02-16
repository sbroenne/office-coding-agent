import React from 'react';
import { describe, it, expect, vi, beforeEach } from 'vitest';
import { screen } from '@testing-library/react';
import userEvent from '@testing-library/user-event';
import { renderWithProviders } from '../test-utils';
import { ChatPanel } from '@/components/ChatPanel';

// ─── Mock child components ───
// ChatPanel is now a thin wrapper around Thread + AgentPicker + ModelPicker.
// We mock children to isolate ChatPanel's own rendering logic.

vi.mock('@/components/assistant-ui/thread', () => ({
  Thread: () => React.createElement('div', { 'data-testid': 'thread' }, 'Thread'),
}));

vi.mock('@/components/AgentPicker', () => ({
  AgentPicker: () => React.createElement('button', { 'data-testid': 'agent-picker' }, 'Excel'),
}));

vi.mock('@/components/ModelPicker', () => ({
  ModelPicker: ({ onOpenSettings }: { onOpenSettings?: () => void }) =>
    React.createElement(
      'button',
      { 'data-testid': 'model-picker', onClick: onOpenSettings },
      'gpt-4.1'
    ),
}));

// ─── Tests ───

describe('ChatPanel', () => {
  beforeEach(() => {
    vi.clearAllMocks();
  });

  it('renders the Thread component', () => {
    renderWithProviders(<ChatPanel isConfigured={true} />);
    expect(screen.getByTestId('thread')).toBeInTheDocument();
  });

  it('renders AgentPicker and ModelPicker in the toolbar', () => {
    renderWithProviders(<ChatPanel isConfigured={true} />);
    expect(screen.getByTestId('agent-picker')).toBeInTheDocument();
    expect(screen.getByTestId('model-picker')).toBeInTheDocument();
  });

  it('shows not-configured warning when isConfigured is false', () => {
    renderWithProviders(<ChatPanel isConfigured={false} />);
    expect(screen.getByText(/No model configured/)).toBeInTheDocument();
  });

  it('hides not-configured warning when isConfigured is true', () => {
    renderWithProviders(<ChatPanel isConfigured={true} />);
    expect(screen.queryByText(/No model configured/)).not.toBeInTheDocument();
  });

  it('shows Open Settings link when onOpenSettings is provided', () => {
    renderWithProviders(<ChatPanel isConfigured={false} onOpenSettings={() => {}} />);
    expect(screen.getByText('Open Settings')).toBeInTheDocument();
  });

  it('calls onOpenSettings when the settings link is clicked', async () => {
    const user = userEvent.setup();
    const onOpenSettings = vi.fn();
    renderWithProviders(<ChatPanel isConfigured={false} onOpenSettings={onOpenSettings} />);
    await user.click(screen.getByText('Open Settings'));
    expect(onOpenSettings).toHaveBeenCalledOnce();
  });

  it('shows fallback text when no onOpenSettings handler is provided', () => {
    renderWithProviders(<ChatPanel isConfigured={false} />);
    expect(screen.getByText(/Open Settings to get started/)).toBeInTheDocument();
  });

  it('passes onOpenSettings to ModelPicker', async () => {
    const user = userEvent.setup();
    const onOpenSettings = vi.fn();
    renderWithProviders(<ChatPanel isConfigured={true} onOpenSettings={onOpenSettings} />);
    await user.click(screen.getByTestId('model-picker'));
    expect(onOpenSettings).toHaveBeenCalledOnce();
  });
});
