import React from 'react';
import { describe, it, expect, vi, beforeEach } from 'vitest';
import { screen } from '@testing-library/react';
import { renderWithProviders } from '../test-utils';
import { ChatPanel } from '@/components/ChatPanel';

// ─── Mock child components ───
// ChatPanel is a thin wrapper around Thread + AgentPicker + ModelPicker.
// We mock children to isolate ChatPanel's own rendering logic.

vi.mock('@/components/assistant-ui/thread', () => ({
  Thread: () => React.createElement('div', { 'data-testid': 'thread' }, 'Thread'),
}));

vi.mock('@/components/AgentPicker', () => ({
  AgentPicker: () => React.createElement('button', { 'data-testid': 'agent-picker' }, 'Excel'),
}));

vi.mock('@/components/ModelPicker', () => ({
  ModelPicker: () =>
    React.createElement('button', { 'data-testid': 'model-picker' }, 'gpt-4.1'),
}));

// ─── Tests ───

describe('ChatPanel', () => {
  beforeEach(() => {
    vi.clearAllMocks();
  });

  it('renders the Thread component', () => {
    renderWithProviders(<ChatPanel />);
    expect(screen.getByTestId('thread')).toBeInTheDocument();
  });

  it('renders AgentPicker and ModelPicker in the toolbar', () => {
    renderWithProviders(<ChatPanel />);
    expect(screen.getByTestId('agent-picker')).toBeInTheDocument();
    expect(screen.getByTestId('model-picker')).toBeInTheDocument();
  });

  it('renders toolbar border between thread and pickers', () => {
    const { container } = renderWithProviders(<ChatPanel />);
    expect(container.querySelector('.border-t')).toBeInTheDocument();
  });
});
