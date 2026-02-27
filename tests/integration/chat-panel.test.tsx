/**
 * Integration test for the ChatPanel component.
 *
 * Renders ChatPanel with real AgentPicker, ModelPicker, and McpPill
 * but mocks the Thread component since it requires an AssistantRuntimeProvider
 * with a real runtime. Tests verify:
 *   - ChatPanel passes leftToolbar and rightToolbar to Thread
 *   - Agent picker, model picker, and tools pill are rendered
 */

import React from 'react';
import { describe, it, expect, beforeEach, vi } from 'vitest';
import { screen } from '@testing-library/react';
import { renderWithProviders } from '../test-utils';
import { ChatPanel } from '@/components/ChatPanel';
import { useSettingsStore } from '@/stores/settingsStore';

// Mock Thread — it requires AssistantRuntimeProvider context.
// Render the toolbar slots so AgentPicker, ModelPicker, McpPill are reachable.
vi.mock('@/components/assistant-ui/thread', () => ({
  Thread: ({
    leftToolbar,
    rightToolbar,
  }: {
    leftToolbar?: React.ReactNode;
    rightToolbar?: React.ReactNode;
  }) =>
    React.createElement(
      'div',
      { 'data-testid': 'thread' },
      React.createElement('div', { 'data-testid': 'left-toolbar' }, leftToolbar),
      React.createElement('div', { 'data-testid': 'right-toolbar' }, rightToolbar)
    ),
}));

// ─── Tests ───

describe('ChatPanel — integration', () => {
  beforeEach(() => {
    vi.clearAllMocks();
    useSettingsStore.getState().reset();
  });

  it('renders Thread with agent picker, model picker, tools pill, and skill picker', () => {
    renderWithProviders(<ChatPanel />);
    expect(screen.getByTestId('thread')).toBeInTheDocument();
    // AgentPicker renders the default agent name in the left toolbar
    expect(screen.getByText('Excel')).toBeInTheDocument();
    // ModelPicker renders with its aria-label in the right toolbar
    expect(screen.getByLabelText('Select model')).toBeInTheDocument();
    // McpPill renders a "Tools" button
    expect(screen.getByLabelText('MCP Tools')).toBeInTheDocument();
    // SkillPicker renders in the left toolbar
    expect(screen.getByLabelText('Agent skills')).toBeInTheDocument();
  });

  it('passes leftToolbar and rightToolbar slots to Thread', () => {
    renderWithProviders(<ChatPanel />);
    // Left slot contains agent + skills + tools pill
    const leftSlot = screen.getByTestId('left-toolbar');
    expect(leftSlot).toContainElement(screen.getByLabelText('Select agent'));
    expect(leftSlot).toContainElement(screen.getByLabelText('Agent skills'));
    expect(leftSlot).toContainElement(screen.getByLabelText('MCP Tools'));
    // Right slot contains the model picker
    const rightSlot = screen.getByTestId('right-toolbar');
    expect(rightSlot).toContainElement(screen.getByLabelText('Select model'));
  });
});
