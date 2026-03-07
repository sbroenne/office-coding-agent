/**
 * Integration tests: ChatComposer slash command menu.
 *
 * Tests the "/" trigger that opens a floating prompt menu populated
 * by pluginPrompts from the settingsStore.
 */
import { describe, it, expect, beforeEach, vi } from 'vitest';
import { render, screen, fireEvent } from '@testing-library/react';
import userEvent from '@testing-library/user-event';
import { ChatComposer } from '@/components/chat/ChatComposer';
import type { PluginPrompt } from '@/types/plugin';

const MOCK_PROMPTS: PluginPrompt[] = [
  {
    name: 'preflight',
    description: 'Run preflight account assessment',
    agent: 'SPT IQ Preflight',
    argumentHint: 'TPID (e.g. 12345678)',
    body: 'Please run a preflight for TPID ${input:tpid}',
  },
  {
    name: 'baseline',
    description: 'ACR baseline analysis',
    agent: 'SPT IQ Consumption',
    argumentHint: 'TPID',
    body: 'Show ACR baseline for ${input:tpid}',
  },
  {
    name: 'tracker',
    description: 'Check deal tracker',
    agent: 'SPT IQ Tracker',
    argumentHint: '',
    body: 'Show deals in tracker',
  },
];

const noop = () => {};

describe('Integration: ChatComposer slash commands', () => {
  let onSend: (text: string) => void;
  let onAgentSelect: (agentName: string) => void;

  beforeEach(() => {
    onSend = vi.fn();
    onAgentSelect = vi.fn();
  });

  it('shows slash menu when user types "/"', async () => {
    render(
      <ChatComposer
        onSend={onSend}
        onCancel={noop}
        isRunning={false}
        slashCommands={MOCK_PROMPTS}
        onAgentSelect={onAgentSelect}
      />
    );

    const input = screen.getByLabelText('Message input');
    await userEvent.type(input, '/');

    expect(screen.getByRole('listbox', { name: /slash commands/i })).toBeInTheDocument();
    expect(screen.getByText('/preflight')).toBeInTheDocument();
    expect(screen.getByText('/baseline')).toBeInTheDocument();
    expect(screen.getByText('/tracker')).toBeInTheDocument();
  });

  it('filters slash menu as user types after "/"', async () => {
    render(
      <ChatComposer
        onSend={onSend}
        onCancel={noop}
        isRunning={false}
        slashCommands={MOCK_PROMPTS}
        onAgentSelect={onAgentSelect}
      />
    );

    const input = screen.getByLabelText('Message input');
    await userEvent.type(input, '/base');

    expect(screen.getByText('/baseline')).toBeInTheDocument();
    expect(screen.queryByText('/preflight')).not.toBeInTheDocument();
  });

  it('closes slash menu when user clears input', async () => {
    render(
      <ChatComposer
        onSend={onSend}
        onCancel={noop}
        isRunning={false}
        slashCommands={MOCK_PROMPTS}
        onAgentSelect={onAgentSelect}
      />
    );

    const input = screen.getByLabelText('Message input');
    await userEvent.type(input, '/');
    expect(screen.getByRole('listbox')).toBeInTheDocument();

    await userEvent.clear(input);
    expect(screen.queryByRole('listbox')).not.toBeInTheDocument();
  });

  it('pressing Escape closes the slash menu', async () => {
    render(
      <ChatComposer
        onSend={onSend}
        onCancel={noop}
        isRunning={false}
        slashCommands={MOCK_PROMPTS}
        onAgentSelect={onAgentSelect}
      />
    );

    const input = screen.getByLabelText('Message input');
    await userEvent.type(input, '/');
    expect(screen.getByRole('listbox')).toBeInTheDocument();

    await userEvent.keyboard('{Escape}');
    expect(screen.queryByRole('listbox')).not.toBeInTheDocument();
  });

  it('clicking a slash command fills the textarea and calls onAgentSelect', async () => {
    render(
      <ChatComposer
        onSend={onSend}
        onCancel={noop}
        isRunning={false}
        slashCommands={MOCK_PROMPTS}
        onAgentSelect={onAgentSelect}
      />
    );

    const input = screen.getByLabelText('Message input') as HTMLTextAreaElement;
    await userEvent.type(input, '/');

    const preflightItem = screen.getByText('/preflight').closest('[role="option"]')!;
    fireEvent.mouseDown(preflightItem);

    // Body should be filled with ${input:tpid} replaced by <tpid>
    expect(input.value).toBe('Please run a preflight for TPID <tpid>');
    expect(onAgentSelect).toHaveBeenCalledWith('SPT IQ Preflight');
    expect(screen.queryByRole('listbox')).not.toBeInTheDocument();
  });

  it('pressing Enter on selected item fills composer and calls onAgentSelect', async () => {
    render(
      <ChatComposer
        onSend={onSend}
        onCancel={noop}
        isRunning={false}
        slashCommands={MOCK_PROMPTS}
        onAgentSelect={onAgentSelect}
      />
    );

    const input = screen.getByLabelText('Message input') as HTMLTextAreaElement;
    await userEvent.type(input, '/');
    // ArrowDown to select second item (baseline), then Enter
    await userEvent.keyboard('{ArrowDown}{Enter}');

    expect(input.value).toBe('Show ACR baseline for <tpid>');
    expect(onAgentSelect).toHaveBeenCalledWith('SPT IQ Consumption');
  });

  it('does not show slash menu when slashCommands is empty', async () => {
    render(
      <ChatComposer
        onSend={onSend}
        onCancel={noop}
        isRunning={false}
        slashCommands={[]}
        onAgentSelect={onAgentSelect}
      />
    );

    const input = screen.getByLabelText('Message input');
    await userEvent.type(input, '/');
    expect(screen.queryByRole('listbox')).not.toBeInTheDocument();
  });

  it('shows description alongside command name in menu', async () => {
    render(
      <ChatComposer
        onSend={onSend}
        onCancel={noop}
        isRunning={false}
        slashCommands={MOCK_PROMPTS}
        onAgentSelect={onAgentSelect}
      />
    );

    const input = screen.getByLabelText('Message input');
    await userEvent.type(input, '/');

    expect(screen.getByText('Run preflight account assessment')).toBeInTheDocument();
  });
});