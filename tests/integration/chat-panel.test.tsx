/**
 * Integration test for the ChatPanel component.
 *
 * Renders ChatPanel with real AgentPicker, ModelPicker, and McpPill
 * but mocks the Thread component since it requires an AssistantRuntimeProvider
 * with a real runtime. Tests verify:
 *   - ChatPanel passes leftToolbar and rightToolbar to Thread
 *   - Agent picker, model picker, and tools pill are rendered
 *
 * Note: SkillPicker lives in ChatHeader, not here.
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

  it('renders Thread with agent picker, model picker, and tools icon', () => {
    renderWithProviders(<ChatPanel />);
    expect(screen.getByTestId('thread')).toBeInTheDocument();
    // AgentPicker renders as an icon button in the left toolbar
    expect(screen.getByLabelText('Select agent')).toBeInTheDocument();
    // ModelPicker renders with its aria-label in the left toolbar
    expect(screen.getByLabelText('Select model')).toBeInTheDocument();
    // McpPill renders as an icon button in the right toolbar
    expect(screen.getByLabelText('MCP Tools')).toBeInTheDocument();
  });

  it('passes leftToolbar and rightToolbar slots to Thread', () => {
    renderWithProviders(<ChatPanel />);
    // Left slot contains agent picker + model picker
    const leftSlot = screen.getByTestId('left-toolbar');
    expect(leftSlot).toContainElement(screen.getByLabelText('Select agent'));
    expect(leftSlot).toContainElement(screen.getByLabelText('Select model'));
    // Right slot contains the tools icon
    const rightSlot = screen.getByTestId('right-toolbar');
    expect(rightSlot).toContainElement(screen.getByLabelText('MCP Tools'));
  });
});
