/**
 * Integration test for the ChatPanel component.
 *
 * Renders ChatPanel with its real message list, composer, and pickers.
 * Tests verify AgentPicker + ModelPicker in left toolbar and McpPicker in right toolbar.
 *
 * Note: plugin help lives in ChatHeader, not here.
 */

import { describe, it, expect, beforeEach, afterEach, vi } from 'vitest';
import { screen, waitFor } from '@testing-library/react';
import userEvent from '@testing-library/user-event';
import { renderWithProviders } from '../test-utils';
import { ChatPanel } from '@/components/ChatPanel';
import { useSettingsStore } from '@/stores/settingsStore';

const DEFAULT_PROPS = {
  messages: [],
  isRunning: false,
  onSend: vi.fn(),
  onCancel: vi.fn(),
  sessionMode: 'interactive' as const,
  onSwitchSessionMode: vi.fn(),
  onOpenPlan: vi.fn(),
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

  it('renders MessageList with CLI agent picker, model picker, and MCP server button', async () => {
    renderWithProviders(<ChatPanel {...DEFAULT_PROPS} />);
    await waitFor(() => expect(globalThis.fetch).toHaveBeenCalled());
    expect(screen.getByRole('textbox', { name: 'Message input' })).toBeInTheDocument();
    expect(screen.getByLabelText('Select agent')).toBeInTheDocument();
    // ModelPicker renders with its aria-label in the left toolbar
    expect(screen.getByLabelText('Select model')).toBeInTheDocument();
    expect(screen.getByLabelText('Select conversation mode')).toBeInTheDocument();
    // McpPicker renders with server icon in the right toolbar (NOT the Plugins button)
    expect(screen.getByLabelText('MCP servers')).toBeInTheDocument();
    // No duplicate Plugins button in the toolbar (it lives in the header only)
    expect(screen.queryByLabelText('Plugins')).not.toBeInTheDocument();
  });

  it('renders both toolbar groups inside the real composer', async () => {
    renderWithProviders(<ChatPanel {...DEFAULT_PROPS} />);
    await waitFor(() => expect(globalThis.fetch).toHaveBeenCalled());
    // Left slot contains CLI agent picker and model picker
    const composer = screen.getByRole('textbox', { name: 'Message input' }).parentElement;
    expect(composer).toContainElement(screen.getByLabelText('Select agent'));
    expect(composer).toContainElement(screen.getByLabelText('Select model'));
    expect(composer).toContainElement(screen.getByLabelText('MCP servers'));
  });

  it.each([
    ['excel', 'office-excel:excel', 'Excel'],
    ['powerpoint', 'office-powerpoint:powerpoint', 'PowerPoint'],
    ['word', 'office-word:word', 'Word'],
  ] as const)('shows the %s host agent as selected by default', async (host, name, displayName) => {
    useSettingsStore.getState().setAvailableAgents([
      { name, displayName, description: `Work with ${displayName}.` },
      { name: 'custom-reviewer', displayName: 'Reviewer', description: 'Review documents.' },
    ]);
    const onSwitchAgent = vi.fn().mockResolvedValue(undefined);
    renderWithProviders(<ChatPanel {...DEFAULT_PROPS} host={host} onSwitchAgent={onSwitchAgent} />);
    await waitFor(() => expect(globalThis.fetch).toHaveBeenCalled());
    expect(screen.getByLabelText('Select agent')).toHaveTextContent(displayName);
    await userEvent.click(screen.getByLabelText('Select agent'));
    const hostOption = screen.getByRole('button', { name: /\(default\)/ });
    expect(hostOption.querySelector('.codicon-check')).toHaveClass('opacity-100');
    expect(screen.queryByText('Office default')).not.toBeInTheDocument();
    await userEvent.click(hostOption);
    expect(onSwitchAgent).toHaveBeenCalledWith(null);
  });

  it('preserves and displays an explicitly selected CLI agent', async () => {
    useSettingsStore
      .getState()
      .setAvailableAgents([
        { name: 'custom-reviewer', displayName: 'Reviewer', description: 'Review documents.' },
      ]);
    useSettingsStore.getState().setActiveAgent('custom-reviewer');
    renderWithProviders(<ChatPanel {...DEFAULT_PROPS} host="excel" />);
    await waitFor(() => expect(globalThis.fetch).toHaveBeenCalled());
    expect(screen.getByLabelText('Select agent')).toHaveTextContent('Reviewer');
  });

  it.each(['Plan', 'Autopilot', 'Interactive'])(
    'requests %s through the mode picker',
    async label => {
      const onSwitchSessionMode = vi.fn().mockResolvedValue(undefined);
      renderWithProviders(
        <ChatPanel {...DEFAULT_PROPS} onSwitchSessionMode={onSwitchSessionMode} />
      );
      const modeButton = screen.getByLabelText('Select conversation mode');
      expect(modeButton).toHaveTextContent('Interactive');
      await userEvent.click(modeButton);
      expect(      screen.getByRole('button', { name: 'Interactive' })).toHaveAttribute(
        'aria-pressed',
        'true'
      );
      await userEvent.click(screen.getByRole('button', { name: label }));
      expect(onSwitchSessionMode).toHaveBeenCalledWith(label.toLowerCase());
    }
  );

  it('shows mode-switch failures instead of ignoring them', async () => {
    const onSwitchSessionMode = vi.fn().mockRejectedValue(new Error('No active session'));
    renderWithProviders(<ChatPanel {...DEFAULT_PROPS} onSwitchSessionMode={onSwitchSessionMode} />);
    await userEvent.click(screen.getByLabelText('Select conversation mode'));
    await userEvent.click(screen.getByRole('button', { name: 'Autopilot' }));
    expect(await screen.findByRole('alert')).toHaveTextContent('No active session');
    expect(screen.getByLabelText('Select conversation mode')).toBeEnabled();
    expect(    screen.getByRole('button', { name: 'Interactive' })).toHaveAttribute(
      'aria-pressed',
      'true'
    );
  });
});
