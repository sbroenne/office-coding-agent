/**
 * Integration test for the ChatPanel component.
 *
 * Renders ChatPanel with real ModelPicker and McpPicker
 * but mocks MessageList since it requires full chat state.
 * Tests verify ModelPicker in left toolbar and McpPicker (server icon) in right toolbar.
 *
 * Note: plugin help lives in ChatHeader, not here.
 */

import React from 'react';
import { describe, it, expect, beforeEach, afterEach, vi } from 'vitest';
import { screen, waitFor } from '@testing-library/react';
import { renderWithProviders } from '../test-utils';
import { ChatPanel } from '@/components/ChatPanel';
import { useSettingsStore } from '@/stores/settingsStore';

// Mock MessageList — render the toolbar slots so ModelPicker and McpPicker are reachable.
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
    vi.spyOn(globalThis, 'fetch').mockResolvedValue({
      ok: true,
      json: async () => ({ servers: [] }),
    } as Response);
    useSettingsStore.getState().reset();
  });

  afterEach(() => {
    vi.restoreAllMocks();
  });

  it('renders MessageList with model picker and MCP server button', async () => {
    renderWithProviders(<ChatPanel {...DEFAULT_PROPS} />);
    await waitFor(() => expect(globalThis.fetch).toHaveBeenCalled());
    expect(screen.getByTestId('message-list')).toBeInTheDocument();
    // ModelPicker renders with its aria-label in the left toolbar
    expect(screen.getByLabelText('Select model')).toBeInTheDocument();
    // McpPicker renders with server icon in the right toolbar (NOT the Plugins button)
    expect(screen.getByLabelText('MCP servers')).toBeInTheDocument();
    // No duplicate Plugins button in the toolbar (it lives in the header only)
    expect(screen.queryByLabelText('Plugins')).not.toBeInTheDocument();
  });

  it('passes leftToolbar and rightToolbar slots to MessageList', async () => {
    renderWithProviders(<ChatPanel {...DEFAULT_PROPS} />);
    await waitFor(() => expect(globalThis.fetch).toHaveBeenCalled());
    // Left slot contains model picker
    const leftSlot = screen.getByTestId('left-toolbar');
    expect(leftSlot).toContainElement(screen.getByLabelText('Select model'));
    // Right slot contains the MCP server picker button
    const rightSlot = screen.getByTestId('right-toolbar');
    expect(rightSlot).toContainElement(screen.getByLabelText('MCP servers'));
  });
});


