/**
 * Regression tests: Queued prompt display in the message list.
 *
 * When the user queues follow-up prompts while the agent is running, each
 * queued prompt must be visible in the thread as a "pending" user message —
 * identical to VS Code Copilot Chat behaviour.
 *
 * Failure modes covered:
 * - Queued prompts are NOT shown at all (only a count badge is displayed)
 * - Queued prompts cannot be individually cancelled
 */
import { describe, it, expect, vi } from 'vitest';
import { screen } from '@testing-library/react';
import userEvent from '@testing-library/user-event';
import { renderWithProviders } from '../test-utils';
import { MessageList } from '@/components/chat/MessageList';
import type { ChatMessage } from '@/types';

// ─── Helpers ──────────────────────────────────────────────────────────────────

const noop = () => {};

function makeUserMsg(text: string): ChatMessage {
  return {
    id: `u-${Math.random()}`,
    role: 'user' as const,
    content: [{ type: 'text' as const, text }],
    createdAt: new Date(),
  };
}

function makeAssistantMsg(text: string): ChatMessage {
  return {
    id: `a-${Math.random()}`,
    role: 'assistant' as const,
    content: [{ type: 'text' as const, text }],
    status: { type: 'complete', reason: 'stop' },
    createdAt: new Date(),
  };
}

// ─── Tests ────────────────────────────────────────────────────────────────────

describe('Integration: Queued prompt display', () => {
  it('shows each queued prompt as visible text in the thread', () => {
    renderWithProviders(
      <MessageList
        messages={[makeUserMsg('First message'), makeAssistantMsg('Response')]}
        isRunning
        onSend={noop}
        onCancel={noop}
        queuedPrompts={['Fix the formula', 'Now add a chart']}
      />
    );

    // Both queued prompts must be visible — not just a count badge
    expect(screen.getByText('Fix the formula')).toBeInTheDocument();
    expect(screen.getByText('Now add a chart')).toBeInTheDocument();
  });

  it('queued prompt items have a visual "queued" indicator', () => {
    renderWithProviders(
      <MessageList
        messages={[]}
        isRunning
        onSend={noop}
        onCancel={noop}
        queuedPrompts={['Do something']}
      />
    );

    const item = screen.getByText('Do something').closest('[data-queued="true"]');
    expect(item).toBeTruthy();
  });

  it('renders queued prompts in order (first in, first out)', () => {
    renderWithProviders(
      <MessageList
        messages={[]}
        isRunning
        onSend={noop}
        onCancel={noop}
        queuedPrompts={['First queued', 'Second queued', 'Third queued']}
      />
    );

    const all = screen.getAllByTestId('queued-prompt');
    expect(all[0]).toHaveTextContent('First queued');
    expect(all[1]).toHaveTextContent('Second queued');
    expect(all[2]).toHaveTextContent('Third queued');
  });

  it('each queued prompt has a cancel button', () => {
    renderWithProviders(
      <MessageList
        messages={[]}
        isRunning
        onSend={noop}
        onCancel={noop}
        queuedPrompts={['Fix the formula', 'Now add a chart']}
      />
    );

    const cancelBtns = screen.getAllByRole('button', { name: /remove from queue/i });
    expect(cancelBtns).toHaveLength(2);
  });

  it('clicking cancel on a queued prompt calls onDequeue with the correct index', async () => {
    const user = userEvent.setup();
    const onDequeue = vi.fn();

    renderWithProviders(
      <MessageList
        messages={[]}
        isRunning
        onSend={noop}
        onCancel={noop}
        queuedPrompts={['First', 'Second', 'Third']}
        onDequeue={onDequeue}
      />
    );

    // Cancel the second queued prompt (index 1)
    const cancelBtns = screen.getAllByRole('button', { name: /remove from queue/i });
    await user.click(cancelBtns[1]);

    expect(onDequeue).toHaveBeenCalledOnce();
    expect(onDequeue).toHaveBeenCalledWith(1);
  });

  it('shows no queued prompts when queuedPrompts is empty', () => {
    renderWithProviders(
      <MessageList
        messages={[makeUserMsg('Hello')]}
        isRunning
        onSend={noop}
        onCancel={noop}
        queuedPrompts={[]}
      />
    );

    expect(screen.queryAllByTestId('queued-prompt')).toHaveLength(0);
  });

  it('shows no queued prompts when queuedPrompts prop is omitted', () => {
    renderWithProviders(
      <MessageList
        messages={[makeUserMsg('Hello')]}
        isRunning
        onSend={noop}
        onCancel={noop}
      />
    );

    expect(screen.queryAllByTestId('queued-prompt')).toHaveLength(0);
  });

  it('queued prompts appear after the last assistant message', () => {
    renderWithProviders(
      <MessageList
        messages={[makeUserMsg('Hello'), makeAssistantMsg('World')]}
        isRunning
        onSend={noop}
        onCancel={noop}
        queuedPrompts={['Queued follow-up']}
      />
    );

    const world = screen.getByText('World');
    const queued = screen.getByText('Queued follow-up');

    // queued message must appear AFTER the last assistant message in the DOM
    expect(
      world.compareDocumentPosition(queued) & Node.DOCUMENT_POSITION_FOLLOWING
    ).toBeTruthy();
  });
});
