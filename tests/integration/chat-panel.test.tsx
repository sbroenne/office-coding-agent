/**
 * Integration test for the ChatPanel component.
 *
 * Renders ChatPanel with real AgentPicker and ModelPicker (inside FluentProvider)
 * but mocks the Thread component since it requires an AssistantRuntimeProvider
 * with a real runtime. Tests verify:
 *   - ChatPanel composition (thread + toolbar)
 *   - Not-configured warning banner behaviour
 *   - Settings callback wiring
 *   - Toolbar renders agent and model pickers
 */

import React from 'react';
import { describe, it, expect, beforeEach, vi } from 'vitest';
import { screen } from '@testing-library/react';
import userEvent from '@testing-library/user-event';
import { renderWithProviders } from '../test-utils';
import { ChatPanel } from '@/components/ChatPanel';
import { useSettingsStore } from '@/stores/settingsStore';

// Mock Thread — it requires AssistantRuntimeProvider context
vi.mock('@/components/assistant-ui/thread', () => ({
  Thread: () => React.createElement('div', { 'data-testid': 'thread' }, 'Thread'),
}));

// ─── Tests ───

describe('ChatPanel — integration', () => {
  beforeEach(() => {
    vi.clearAllMocks();
    useSettingsStore.getState().reset();
  });

  it('renders Thread, AgentPicker, and ModelPicker together', () => {
    renderWithProviders(<ChatPanel isConfigured={true} />);
    expect(screen.getByTestId('thread')).toBeInTheDocument();
    // AgentPicker renders agent name (default agent "Excel"); it uses Fluent Menu
    expect(screen.getByText('Excel')).toBeInTheDocument();
  });

  it('shows warning banner when not configured', () => {
    renderWithProviders(<ChatPanel isConfigured={false} />);
    expect(screen.getByText(/No model configured/)).toBeInTheDocument();
  });

  it('hides warning banner when configured', () => {
    renderWithProviders(<ChatPanel isConfigured={true} />);
    expect(screen.queryByText(/No model configured/)).not.toBeInTheDocument();
  });

  it('renders Open Settings link in warning when handler is provided', () => {
    const handler = vi.fn();
    renderWithProviders(<ChatPanel isConfigured={false} onOpenSettings={handler} />);
    const link = screen.getByText('Open Settings');
    expect(link).toBeInTheDocument();
    expect(link.tagName).toBe('BUTTON');
  });

  it('calls onOpenSettings when the warning link is clicked', async () => {
    const user = userEvent.setup();
    const handler = vi.fn();
    renderWithProviders(<ChatPanel isConfigured={false} onOpenSettings={handler} />);
    await user.click(screen.getByText('Open Settings'));
    expect(handler).toHaveBeenCalledOnce();
  });

  it('shows static fallback when no onOpenSettings is provided', () => {
    renderWithProviders(<ChatPanel isConfigured={false} />);
    expect(screen.getByText(/Open Settings to get started/)).toBeInTheDocument();
    expect(screen.queryByRole('button', { name: 'Open Settings' })).not.toBeInTheDocument();
  });

  it('renders toolbar border between thread and pickers', () => {
    const { container } = renderWithProviders(<ChatPanel isConfigured={true} />);
    const toolbar = container.querySelector('.border-t');
    expect(toolbar).toBeInTheDocument();
  });
});
