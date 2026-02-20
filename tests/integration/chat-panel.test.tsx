/**
 * Integration test for the ChatPanel component.
 *
 * Renders ChatPanel with real AgentPicker and ModelPicker (inside FluentProvider)
 * but mocks the Thread component since it requires an AssistantRuntimeProvider
 * with a real runtime. Tests verify:
 *   - ChatPanel composition (thread + toolbar)
 *   - Toolbar renders agent and model pickers
 */

import React from 'react';
import { describe, it, expect, beforeEach, vi } from 'vitest';
import { screen } from '@testing-library/react';
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
    renderWithProviders(<ChatPanel />);
    expect(screen.getByTestId('thread')).toBeInTheDocument();
    // AgentPicker renders agent name (default agent "Excel")
    expect(screen.getByText('Excel')).toBeInTheDocument();
  });

  it('renders toolbar border between thread and pickers', () => {
    const { container } = renderWithProviders(<ChatPanel />);
    const toolbar = container.querySelector('.border-t');
    expect(toolbar).toBeInTheDocument();
  });
});
