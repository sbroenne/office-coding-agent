/**
 * Integration test for the ChatPanel component.
 *
 * Renders ChatPanel with real AgentPicker, ModelPicker, and McpPill
 * but mocks MessageList since it requires full chat state.
 * Tests verify AgentPicker, ModelPicker, and McpPill are rendered in their slots.
 *
 * Note: SkillPicker lives in ChatHeader, not here.
 */

import React from 'react';
import { describe, it, expect, beforeEach, vi } from 'vitest';
import { screen } from '@testing-library/react';
import { renderWithProviders } from '../test-utils';
import { ChatPanel } from '@/components/ChatPanel';
import { useSettingsStore } from '@/stores/settingsStore';

// Mock MessageList — render the toolbar slots so AgentPicker, ModelPicker, McpPill are reachable.
vi.mock('@/components/chat/MessageList', () => ({
  MessageList: ({
    leftToolbar,
    rightToolbar,
  }: {
    leftToolbar?: React.ReactNode;
    rightToolbar?: React.ReactNode;
  }) =>
    React.createElement(
      'div',
      { 'data-testid': 'message-list' },
      React.createElement('div', { 'data-testid': 'left-toolbar' }, leftToolbar),
      React.createElement('div', { 'data-testid': 'right-toolbar' }, rightToolbar)
    ),
}));

const DEFAULT_PROPS = {
  messages: [],
  isRunning: false,
  onSend: vi.fn(),
  onCancel: vi.fn(),
};

// ─── Tests ───

describe('ChatPanel — integration', () => {
  beforeEach(() => {
    vi.clearAllMocks();
    useSettingsStore.getState().reset();
  });

  it('renders MessageList with agent picker, model picker, and plugins icon', () => {
    renderWithProviders(<ChatPanel {...DEFAULT_PROPS} />);
    expect(screen.getByTestId('message-list')).toBeInTheDocument();
    // AgentPicker renders as an icon button in the left toolbar
    expect(screen.getByLabelText('Select agent')).toBeInTheDocument();
    // ModelPicker renders with its aria-label in the left toolbar
    expect(screen.getByLabelText('Select model')).toBeInTheDocument();
    // Plugins pill renders as an icon button in the right toolbar
    expect(screen.getByLabelText('Plugins')).toBeInTheDocument();
  });

  it('passes leftToolbar and rightToolbar slots to MessageList', () => {
    renderWithProviders(<ChatPanel {...DEFAULT_PROPS} />);
    // Left slot contains agent picker + model picker
    const leftSlot = screen.getByTestId('left-toolbar');
    expect(leftSlot).toContainElement(screen.getByLabelText('Select agent'));
    expect(leftSlot).toContainElement(screen.getByLabelText('Select model'));
    // Right slot contains the plugins icon
    const rightSlot = screen.getByTestId('right-toolbar');
    expect(rightSlot).toContainElement(screen.getByLabelText('Plugins'));
  });
});

