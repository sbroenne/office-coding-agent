/**
 * Integration tests for Thread / AssistantMessage rendering behaviour.
 *
 * Specifically covers:
 *  1. Action bar (copy / thumbsup / thumbsdown) is present in the DOM after
 *     a completed assistant text response.
 *  2. The "MessagePartText can only be used inside text or reasoning message
 *     parts" crash does NOT occur when the assistant response ends with a
 *     tool-call part (i.e. no trailing text part).
 *
 * Uses the same fake session/client pattern as use-office-chat.test.tsx.
 * No real WebSocket or Copilot API is required.
 */

import { describe, it, expect, vi, beforeEach } from 'vitest';
import { render, screen, waitFor, act, fireEvent } from '@testing-library/react';
import type { SessionEvent } from '@github/copilot-sdk';
import { MessageList } from '@/components/chat/MessageList';
import { ChatActionsContext } from '@/contexts/ChatActionsContext';
import { useOfficeChat } from '@/hooks/useOfficeChat';
import { useSettingsStore } from '@/stores/settingsStore';
import { useSessionHistoryStore } from '@/stores/sessionHistoryStore';

// ─── Fake session / client ────────────────────────────────────────────────────

type EventEmitter = (event: SessionEvent) => void;

function makeFakeSession(events: SessionEvent[]) {
  return {
    sessionId: 'test-session-id',
    async *query() {
      for (const event of events) {
        yield event;
        if (event.type === 'session.idle') return;
      }
    },
    on: vi.fn(),
    onPermissionRequest: vi.fn(() => () => undefined),
    destroy: vi.fn().mockResolvedValue(undefined),
    send: vi.fn().mockResolvedValue('msg-id'),
    registerTools: vi.fn(),
    getToolHandler: vi.fn(),
    respondPermission: vi.fn().mockResolvedValue(undefined),
    _dispatchEvent: vi.fn() as EventEmitter,
  };
}

function makeFakeClient(session: ReturnType<typeof makeFakeSession>) {
  return {
    start: vi.fn().mockResolvedValue(undefined),
    createSession: vi.fn().mockResolvedValue(session),
    listModels: vi.fn().mockResolvedValue([]),
    stop: vi.fn().mockResolvedValue(undefined),
    onMcpStatus: vi.fn(() => () => undefined),
    onMcpLog: vi.fn(() => () => undefined),
    onMcpTools: vi.fn(() => () => undefined),
    onPluginAgents: vi.fn(() => () => undefined),
    onPluginSkills: vi.fn(() => () => undefined),
    onPluginPrompts: vi.fn(() => () => undefined),
  };
}

function makeEvent<T extends SessionEvent['type']>(
  type: T,
  data: Extract<SessionEvent, { type: T }>['data']
): SessionEvent {
  return {
    id: 'e1',
    timestamp: new Date().toISOString(),
    parentId: null,
    type,
    data,
  } as SessionEvent;
}

const IDLE_EVENT = makeEvent('session.idle', {});

vi.mock('@/lib/websocket-client', () => ({
  createWebSocketClient: vi.fn(),
}));

import { createWebSocketClient } from '@/lib/websocket-client';
const mockCreate = vi.mocked(createWebSocketClient);

// ─── Shared test wrapper ──────────────────────────────────────────────────────

/**
 * Renders MessageList with a real (faked) useOfficeChat and returns
 * the hook result so tests can drive messages through `send`.
 */
function renderThreadWithHook(host: 'excel' = 'excel') {
  let hookRef: ReturnType<typeof useOfficeChat> | undefined;

  function TestComponent() {
    const chat = useOfficeChat(host);
    hookRef = chat;
    return (
      <ChatActionsContext.Provider value={{ send: chat.send, enqueue: () => {} }}>
        <MessageList
          messages={chat.messages}
          isRunning={chat.isRunning}
          onSend={chat.send}
          onCancel={chat.cancel}
          onRegenerate={vi.fn()}
          onFeedback={vi.fn()}
          onEdit={vi.fn()}
        />
      </ChatActionsContext.Provider>
    );
  }

  render(<TestComponent />);
  return { getHook: () => hookRef! };
}

// ─── Tests ────────────────────────────────────────────────────────────────────

describe('Thread – AssistantMessage rendering', () => {
  beforeEach(() => {
    vi.clearAllMocks();
    useSettingsStore.getState().reset();
    useSessionHistoryStore.setState({ sessions: [], activeSessionId: null });
  });

  it('renders the action bar (copy + thumbsup + thumbsdown) in a completed text response', async () => {
    const session = makeFakeSession([
      makeEvent('assistant.message', { messageId: 'msg1', content: 'Hello!' }),
      IDLE_EVENT,
    ]);
    const client = makeFakeClient(session);
    mockCreate.mockResolvedValue(client as never);

    const { getHook } = renderThreadWithHook();

    await act(async () => {
      await new Promise(r => setTimeout(r, 100));
    });

    await act(async () => {
      void getHook().send('test');
      await new Promise(r => setTimeout(r, 200));
    });

    // Wait for the assistant message bubble to appear
    await waitFor(() => {
      expect(document.querySelector('[data-role="assistant"]')).toBeInTheDocument();
    });

    // Action bar must exist in the DOM (opacity-0 when not hovered, but present)
    const actionBar = document.querySelector('.aui-assistant-action-bar');
    expect(actionBar).toBeInTheDocument();

    // All three buttons must be present.
    // TooltipIconButton renders the tooltip text as an sr-only span inside the button,
    // making it the accessible name — so getByRole('button', { name }) is the right query.
    expect(screen.getByRole('button', { name: 'Copy' })).toBeInTheDocument();
    expect(screen.getByRole('button', { name: 'Good response' })).toBeInTheDocument();
    expect(screen.getByRole('button', { name: 'Bad response' })).toBeInTheDocument();

    // Initially hidden (no hover) — CSS opacity-0 class is applied
    expect(actionBar).toHaveClass('opacity-0');
  });

  it('action bar is positioned inside the assistant message container', async () => {
    const session = makeFakeSession([
      makeEvent('assistant.message', { messageId: 'msg1', content: 'Done!' }),
      IDLE_EVENT,
    ]);
    const client = makeFakeClient(session);
    mockCreate.mockResolvedValue(client as never);

    const { getHook } = renderThreadWithHook();

    await act(async () => {
      await new Promise(r => setTimeout(r, 100));
    });
    await act(async () => {
      void getHook().send('test');
      await new Promise(r => setTimeout(r, 200));
    });

    await waitFor(() => {
      expect(document.querySelector('[data-role="assistant"]')).toBeInTheDocument();
    });

    const msg = document.querySelector('[data-role="assistant"]');
    const actionBar = msg?.querySelector('.aui-assistant-action-bar');
    expect(actionBar).toBeTruthy();
  });

  it('does NOT crash when the assistant response contains only tool-call parts (no text)', async () => {
    // This is the regression test for "MessagePartText can only be used inside
    // text or reasoning message parts".
    //
    // When a response ends with a tool-call part and has no text part,
    // rendering the message must handle the empty text section gracefully
    // without throwing.
    const session = makeFakeSession([
      makeEvent('tool.execution_start', {
        toolCallId: 'tc1',
        toolName: 'manage_plugins',
        arguments: { action: 'list' },
      }),
      makeEvent('tool.execution_complete', {
        toolCallId: 'tc1',
        success: true,
        result: { content: '[]' },
      }),
      // ⚠️  No assistant.message event — response ends at the tool call
      IDLE_EVENT,
    ]);
    const client = makeFakeClient(session);
    mockCreate.mockResolvedValue(client as never);

    const { getHook } = renderThreadWithHook();

    await act(async () => {
      await new Promise(r => setTimeout(r, 100));
    });
    await act(async () => {
      void getHook().send('test');
      await new Promise(r => setTimeout(r, 200));
    });

    await waitFor(() => {
      expect(document.querySelector('[data-role="assistant"]')).toBeInTheDocument();
    });

    // The critical assertion: no error boundary and no crash text
    expect(screen.queryByText(/something went wrong/i)).not.toBeInTheDocument();
    expect(screen.queryByText(/MessagePartText can only/i)).not.toBeInTheDocument();

    // The tool card should render normally
    expect(document.querySelector('[data-slot="tool-fallback-root"]')).toBeInTheDocument();
  });

  it('renders the message text content alongside the action bar', async () => {
    const session = makeFakeSession([
      makeEvent('assistant.message', { messageId: 'msg1', content: 'All good here!' }),
      IDLE_EVENT,
    ]);
    const client = makeFakeClient(session);
    mockCreate.mockResolvedValue(client as never);

    const { getHook } = renderThreadWithHook();

    await act(async () => {
      await new Promise(r => setTimeout(r, 100));
    });
    await act(async () => {
      void getHook().send('test');
      await new Promise(r => setTimeout(r, 200));
    });

    await waitFor(() => {
      expect(screen.getByText('All good here!')).toBeInTheDocument();
    });

    // Action bar is still present alongside the text
    expect(document.querySelector('.aui-assistant-action-bar')).toBeInTheDocument();
  });

  it('thinking indicator appears with "Thinking…" during streaming', async () => {
    // Use a session that yields a tool start + complete + text so the stream
    // takes some time, then idle.
    const session = makeFakeSession([
      makeEvent('tool.execution_start', {
        toolCallId: 'tc1',
        toolName: 'get_range_values',
        arguments: { address: 'A1:C3' },
      }),
      makeEvent('tool.execution_complete', {
        toolCallId: 'tc1',
        success: true,
        result: { content: '[[1,2],[3,4]]' },
      }),
      makeEvent('assistant.message', { messageId: 'msg1', content: 'Here is the data.' }),
      IDLE_EVENT,
    ]);
    const client = makeFakeClient(session);
    mockCreate.mockResolvedValue(client as never);

    const { getHook } = renderThreadWithHook();

    await act(async () => {
      await new Promise(r => setTimeout(r, 100));
    });
    await act(async () => {
      void getHook().send('test');
      await new Promise(r => setTimeout(r, 200));
    });

    // After the stream completes, the shimmer thinking progress should be gone
    await waitFor(() => {
      expect(document.querySelector('.inline-working-progress')).not.toBeInTheDocument();
    });
  });

  it('thinking indicator is rendered inside the assistant message, not at thread level', async () => {
    const session = makeFakeSession([
      makeEvent('assistant.message', { messageId: 'msg1', content: 'Done' }),
      IDLE_EVENT,
    ]);
    const client = makeFakeClient(session);
    mockCreate.mockResolvedValue(client as never);

    const { getHook } = renderThreadWithHook();

    await act(async () => {
      await new Promise(r => setTimeout(r, 100));
    });
    await act(async () => {
      void getHook().send('test');
      await new Promise(r => setTimeout(r, 200));
    });

    await waitFor(() => {
      expect(document.querySelector('[data-role="assistant"]')).toBeInTheDocument();
    });

    // The thinking indicator is now rendered INSIDE the assistant message (as InlineWorkingProgress)
    // After completion, it should not be visible
    const assistantMsg = document.querySelector('[data-role="assistant"]');
    const indicatorInsideMsg = assistantMsg?.querySelector('.inline-working-progress');
    // After completion (text arrived), the shimmer progress should be gone
    expect(indicatorInsideMsg).toBeNull();
  });

  it('does NOT crash when message has only tool-call parts with empty streamText', async () => {
    // Regression: updateAssistant() used to always include { type: 'text', text: '' }
    // which fromThreadMessageLike() would strip, producing 0-part or non-text-ending
    // messages that crashed MarkdownTextPrimitive.
    const session = makeFakeSession([
      makeEvent('tool.execution_start', {
        toolCallId: 'tc1',
        toolName: 'set_range_values',
        arguments: { address: 'A1', values: [[1]] },
      }),
      makeEvent('tool.execution_complete', {
        toolCallId: 'tc1',
        success: true,
        result: { content: 'OK' },
      }),
      makeEvent('tool.execution_start', {
        toolCallId: 'tc2',
        toolName: 'get_range_values',
        arguments: { address: 'A1:B2' },
      }),
      makeEvent('tool.execution_complete', {
        toolCallId: 'tc2',
        success: true,
        result: { content: '[[1,2]]' },
      }),
      // assistant.message with text arrives only AFTER all tool calls
      makeEvent('assistant.message', { messageId: 'msg1', content: 'Done!' }),
      IDLE_EVENT,
    ]);
    const client = makeFakeClient(session);
    mockCreate.mockResolvedValue(client as never);

    const { getHook } = renderThreadWithHook();

    await act(async () => {
      await new Promise(r => setTimeout(r, 100));
    });
    await act(async () => {
      void getHook().send('test');
      await new Promise(r => setTimeout(r, 300));
    });

    // Must not crash
    expect(screen.queryByText(/something went wrong/i)).not.toBeInTheDocument();
    expect(screen.queryByText(/MessagePartText can only/i)).not.toBeInTheDocument();

    // Final text should render
    await waitFor(() => {
      expect(screen.getByText('Done!')).toBeInTheDocument();
    });
  });

  it('does NOT crash when Thread renders mid-stream with tool parts and empty streamText', async () => {
    // KEY REGRESSION TEST: The crash only happened when React rendered the
    // Thread while the assistant message was in-progress with tool-call parts
    // and streamText = ''.  Fast-session tests miss this because by the time
    // React paints, the stream is already done.  This test pauses mid-stream
    // so the Thread renders the intermediate state.
    let resolveStream!: () => void;
    const streamGate = new Promise<void>(r => {
      resolveStream = r;
    });

    const pausingSession = {
      sessionId: 'test-session-id',
      async *query() {
        yield makeEvent('tool.execution_start', {
          toolCallId: 'tc1',
          toolName: 'set_range_values',
          arguments: { address: 'A1', values: [[42]] },
        });
        yield makeEvent('tool.execution_complete', {
          toolCallId: 'tc1',
          success: true,
          result: { content: 'OK' },
        });
        // ── PAUSE ── React will render the in-progress message here.
        // At this point: toolParts has tc1, streamText is '' (no text yet).
        await streamGate;
        yield makeEvent('assistant.message', { messageId: 'msg1', content: 'All set.' });
        yield IDLE_EVENT;
      },
      on: vi.fn(),
      onPermissionRequest: vi.fn(() => () => undefined),
      destroy: vi.fn().mockResolvedValue(undefined),
      send: vi.fn().mockResolvedValue('msg-id'),
      registerTools: vi.fn(),
      getToolHandler: vi.fn(),
      respondPermission: vi.fn().mockResolvedValue(undefined),
      _dispatchEvent: vi.fn(),
    };
    const client = makeFakeClient(pausingSession as ReturnType<typeof makeFakeSession>);
    mockCreate.mockResolvedValue(client as never);

    const { getHook } = renderThreadWithHook();

    await act(async () => {
      await new Promise(r => setTimeout(r, 100));
    });

    // Start the stream — it will pause after the tool completes
    await act(async () => {
      void getHook().send('test');
      await new Promise(r => setTimeout(r, 150));
    });

    // The Thread has now rendered the in-progress assistant message with
    // only a tool-call part and no text.  This is where the old code crashed
    // because updateAssistant() sent { text: '' } which got stripped by
    // fromThreadMessageLike(), leaving only the tool-call part.
    expect(screen.queryByText(/MessagePartText can only/i)).not.toBeInTheDocument();
    expect(screen.queryByText(/something went wrong/i)).not.toBeInTheDocument();

    // The tool card should be visible
    expect(document.querySelector('[data-slot="tool-fallback-root"]')).toBeInTheDocument();

    // Release the stream and let it finish
    await act(async () => {
      resolveStream();
      await new Promise(r => setTimeout(r, 200));
    });

    // Final text should render
    await waitFor(() => {
      expect(screen.getByText('All set.')).toBeInTheDocument();
    });
  });

  it('thinking indicator stays as Thinking during tool execution (VS Code behavior)', async () => {
    // In VS Code, the "Thinking" label stays constant. Tool names are shown
    // in the tool cards, not in the thinking indicator text.
    let resolveStream!: () => void;
    const streamGate = new Promise<void>(r => {
      resolveStream = r;
    });

    const pausingSession = {
      sessionId: 'test-session-id',
      async *query() {
        yield makeEvent('tool.execution_start', {
          toolCallId: 'tc1',
          toolName: 'get_range_values',
          arguments: { address: 'A1:C3' },
        });
        // ── PAUSE ── thinking text should stay as "Thinking…"
        await streamGate;
        yield makeEvent('tool.execution_complete', {
          toolCallId: 'tc1',
          success: true,
          result: { content: '[[1]]' },
        });
        yield IDLE_EVENT;
      },
      on: vi.fn(),
      onPermissionRequest: vi.fn(() => () => undefined),
      destroy: vi.fn().mockResolvedValue(undefined),
      send: vi.fn().mockResolvedValue('msg-id'),
      registerTools: vi.fn(),
      getToolHandler: vi.fn(),
      respondPermission: vi.fn().mockResolvedValue(undefined),
      setModel: vi.fn().mockResolvedValue(undefined),
      compact: vi.fn().mockResolvedValue(undefined),
      _dispatchEvent: vi.fn(),
    };
    const client = makeFakeClient(pausingSession as ReturnType<typeof makeFakeSession>);
    mockCreate.mockResolvedValue(client as never);

    const { getHook } = renderThreadWithHook();

    await act(async () => {
      await new Promise(r => setTimeout(r, 100));
    });

    // Start the stream — it will pause after tool.execution_start
    await act(async () => {
      void getHook().send('test');
      await new Promise(r => setTimeout(r, 150));
    });

    // The shimmer progress hides once a tool card appears (VS Code behavior)
    // Tool cards handle their own shimmer on the tool name while running
    const indicator = document.querySelector('.inline-working-progress');
    const toolCard = document.querySelector('[data-slot="tool-fallback-root"]');

    // Either: shimmer is gone (tool card appeared) or shimmer is visible (tool hasn't rendered yet)
    // Once the tool card is in the DOM, shimmer must be gone
    if (toolCard) {
      expect(indicator).not.toBeInTheDocument();
    }
    // At minimum, a tool card should exist since tool.execution_start was emitted
    expect(toolCard).toBeInTheDocument();

    // Release and finish
    await act(async () => {
      resolveStream();
      await new Promise(r => setTimeout(r, 200));
    });

    // Indicator should be gone after completion
    expect(document.querySelector('.inline-working-progress')).not.toBeInTheDocument();
  });

  it('thinking indicator shows report_intent text in the rendered DOM', async () => {
    let resolveStream!: () => void;
    const streamGate = new Promise<void>(r => {
      resolveStream = r;
    });

    const pausingSession = {
      sessionId: 'test-session-id',
      async *query() {
        yield makeEvent('tool.execution_start', {
          toolCallId: 'ri1',
          toolName: 'report_intent',
          arguments: { intent: 'Analyzing your data' },
        });
        // ── PAUSE ── thinking text should now say "Analyzing your data"
        await streamGate;
        yield makeEvent('assistant.message', { messageId: 'msg1', content: 'Analysis done.' });
        yield IDLE_EVENT;
      },
      on: vi.fn(),
      onPermissionRequest: vi.fn(() => () => undefined),
      destroy: vi.fn().mockResolvedValue(undefined),
      send: vi.fn().mockResolvedValue('msg-id'),
      registerTools: vi.fn(),
      getToolHandler: vi.fn(),
      respondPermission: vi.fn().mockResolvedValue(undefined),
      _dispatchEvent: vi.fn(),
    };
    const client = makeFakeClient(pausingSession as ReturnType<typeof makeFakeSession>);
    mockCreate.mockResolvedValue(client as never);

    const { getHook } = renderThreadWithHook();

    await act(async () => {
      await new Promise(r => setTimeout(r, 100));
    });

    await act(async () => {
      void getHook().send('test');
      await new Promise(r => setTimeout(r, 150));
    });

    // The thinking indicator should show the intent text inside the assistant message
    const indicator = document.querySelector('.inline-working-progress');
    expect(indicator).toBeInTheDocument();
    expect(indicator!.textContent).toContain('Analyzing your data');

    await act(async () => {
      resolveStream();
      await new Promise(r => setTimeout(r, 200));
    });

    expect(document.querySelector('.inline-working-progress')).not.toBeInTheDocument();
  });
});

// ────────────────────────────────────────────────────────────────────────────
// Working box spinner between tool steps (visual regression)
// ────────────────────────────────────────────────────────────────────────────

describe('Working box spinner between tool steps (thinking gap fix)', () => {
  beforeEach(() => {
    vi.clearAllMocks();
    useSettingsStore.getState().reset();
    useSessionHistoryStore.setState({ sessions: [], activeSessionId: null });
  });

  it('[REGRESSION] report_intent text becomes the Working box header label (VS Code IChatTask.content)', async () => {
    // VS Code: IChatTask.content = the task label, shown in BOTH running and done states.
    // report_intent fires BEFORE tools → sets label for phase 0 Working box.
    const session = makeFakeSession([
      makeEvent('tool.execution_start', {
        toolCallId: 'ri1',
        toolName: 'report_intent',
        arguments: { intent: 'Reading your spreadsheet' },
      }),
      makeEvent('tool.execution_start', {
        toolCallId: 'tc1',
        toolName: 'get_range_values',
        arguments: { address: 'A1' },
      }),
      makeEvent('tool.execution_complete', {
        toolCallId: 'tc1',
        success: true,
        result: { content: '[[42]]' },
      }),
      makeEvent('assistant.message', { messageId: 'msg1', content: 'Value is 42.' }),
      IDLE_EVENT,
    ]);
    const client = makeFakeClient(session);
    mockCreate.mockResolvedValue(client as never);

    const { getHook } = renderThreadWithHook();
    await act(async () => { await new Promise(r => setTimeout(r, 100)); });
    await act(async () => {
      void getHook().send('test');
      await new Promise(r => setTimeout(r, 300));
    });

    await waitFor(() => { expect(screen.getByText('Value is 42.')).toBeInTheDocument(); });

    // The completed Working box header must show the report_intent text
    const doneTitle = document.querySelector('.chat-thinking-title-done');
    expect(doneTitle).toBeInTheDocument();
    expect(doneTitle!.textContent).toBe('Reading your spreadsheet');
  });

  it('[REGRESSION] report_intent after tools sets label on the SECOND Working box', async () => {
    // When report_intent fires AFTER the first tool, phase 0 box has no label ("Working"),
    // and phase 1 box gets the intent text as its label.
    const session = makeFakeSession([
      makeEvent('tool.execution_start', {
        toolCallId: 'tc1',
        toolName: 'get_range_values',
        arguments: { address: 'A1' },
      }),
      makeEvent('tool.execution_complete', {
        toolCallId: 'tc1',
        success: true,
        result: { content: '[[1]]' },
      }),
      makeEvent('tool.execution_start', {
        toolCallId: 'ri1',
        toolName: 'report_intent',
        arguments: { intent: 'Writing result' },
      }),
      makeEvent('tool.execution_start', {
        toolCallId: 'tc2',
        toolName: 'set_range_values',
        arguments: { address: 'B1', values: [[2]] },
      }),
      makeEvent('tool.execution_complete', {
        toolCallId: 'tc2',
        success: true,
        result: { content: 'OK' },
      }),
      makeEvent('assistant.message', { messageId: 'msg1', content: 'Done.' }),
      IDLE_EVENT,
    ]);
    const client = makeFakeClient(session);
    mockCreate.mockResolvedValue(client as never);

    const { getHook } = renderThreadWithHook();
    await act(async () => { await new Promise(r => setTimeout(r, 100)); });
    await act(async () => {
      void getHook().send('test');
      await new Promise(r => setTimeout(r, 300));
    });

    await waitFor(() => { expect(screen.getByText('Done.')).toBeInTheDocument(); });

    const boxes = document.querySelectorAll('.chat-thinking-box');
    expect(boxes).toHaveLength(2);

    // Phase 0 box: no label → "Working"
    const box0Title = boxes[0]!.querySelector('.chat-thinking-title-done');
    expect(box0Title!.textContent).toBe('Working');

    // Phase 1 box: label = "Writing result"
    const box1Title = boxes[1]!.querySelector('.chat-thinking-title-done');
    expect(box1Title!.textContent).toBe('Writing result');
  });

  it('shows spinner inside Working box after a tool completes (inter-step thinking)', async () => {
    // KEY regression test for the thinking gap bug.
    // After a tool completes and before the next tool/text starts,
    // the spinner inside the Working box must show "Thinking…"
    let resolveStream!: () => void;
    const streamGate = new Promise<void>(r => {
      resolveStream = r;
    });

    const pausingSession = {
      sessionId: 'test-session-id',
      async *query() {
        yield makeEvent('tool.execution_start', {
          toolCallId: 'tc1',
          toolName: 'get_range_values',
          arguments: { address: 'A1:C3' },
        });
        yield makeEvent('tool.execution_complete', {
          toolCallId: 'tc1',
          success: true,
          result: { content: '[[1,2,3]]' },
        });
        // ── PAUSE ── model is thinking between tools
        await streamGate;
        yield makeEvent('assistant.message', { messageId: 'msg1', content: 'Here is the data.' });
        yield IDLE_EVENT;
      },
      on: vi.fn(),
      onPermissionRequest: vi.fn(() => () => undefined),
      destroy: vi.fn().mockResolvedValue(undefined),
      send: vi.fn().mockResolvedValue('msg-id'),
      registerTools: vi.fn(),
      getToolHandler: vi.fn(),
      respondPermission: vi.fn().mockResolvedValue(undefined),
      setModel: vi.fn().mockResolvedValue(undefined),
      compact: vi.fn().mockResolvedValue(undefined),
      _dispatchEvent: vi.fn(),
    };
    const client = makeFakeClient(pausingSession as ReturnType<typeof makeFakeSession>);
    mockCreate.mockResolvedValue(client as never);

    const { getHook } = renderThreadWithHook();

    await act(async () => {
      await new Promise(r => setTimeout(r, 100));
    });

    await act(async () => {
      void getHook().send('test');
      await new Promise(r => setTimeout(r, 200));
    });

    // Inline shimmer should NOT show (tools exist now)
    expect(document.querySelector('.inline-working-progress')).not.toBeInTheDocument();

    // Working box must be present with "Working" header
    const workingHeader = document.querySelector('.chat-thinking-title-shimmer');
    expect(workingHeader).toBeInTheDocument();
    expect(workingHeader!.textContent).toBe('Working');

    // Spinner inside Working box must show "Thinking…"
    const spinner = document.querySelector('[data-testid="working-spinner"]');
    expect(spinner).toBeInTheDocument();
    expect(spinner!.textContent).toContain('Thinking');

    // The spinner label must have the shimmer CSS class
    const spinnerLabel = spinner!.querySelector('.chat-thinking-spinner-label');
    expect(spinnerLabel).toBeInTheDocument();

    // Release stream and finish
    await act(async () => {
      resolveStream();
      await new Promise(r => setTimeout(r, 200));
    });

    // After completion, spinner should be gone (Working box collapses)
    expect(document.querySelector('[data-testid="working-spinner"]')).not.toBeInTheDocument();
  });

  it('spinner hides while a tool is actively running', async () => {
    // When a tool is running, its card shows shimmer — no need for spinner
    let resolveStream!: () => void;
    const streamGate = new Promise<void>(r => {
      resolveStream = r;
    });

    const pausingSession = {
      sessionId: 'test-session-id',
      async *query() {
        yield makeEvent('tool.execution_start', {
          toolCallId: 'tc1',
          toolName: 'get_range_values',
          arguments: { address: 'A1:C3' },
        });
        // ── PAUSE ── tool is running
        await streamGate;
        yield makeEvent('tool.execution_complete', {
          toolCallId: 'tc1',
          success: true,
          result: { content: '[[1]]' },
        });
        yield IDLE_EVENT;
      },
      on: vi.fn(),
      onPermissionRequest: vi.fn(() => () => undefined),
      destroy: vi.fn().mockResolvedValue(undefined),
      send: vi.fn().mockResolvedValue('msg-id'),
      registerTools: vi.fn(),
      getToolHandler: vi.fn(),
      respondPermission: vi.fn().mockResolvedValue(undefined),
      setModel: vi.fn().mockResolvedValue(undefined),
      compact: vi.fn().mockResolvedValue(undefined),
      _dispatchEvent: vi.fn(),
    };
    const client = makeFakeClient(pausingSession as ReturnType<typeof makeFakeSession>);
    mockCreate.mockResolvedValue(client as never);

    const { getHook } = renderThreadWithHook();

    await act(async () => {
      await new Promise(r => setTimeout(r, 100));
    });
    await act(async () => {
      void getHook().send('test');
      await new Promise(r => setTimeout(r, 200));
    });

    // Tool card should exist (running)
    expect(document.querySelector('[data-slot="tool-fallback-root"]')).toBeInTheDocument();

    // Spinner must NOT show while a tool is actively running
    expect(document.querySelector('[data-testid="working-spinner"]')).not.toBeInTheDocument();

    await act(async () => {
      resolveStream();
      await new Promise(r => setTimeout(r, 200));
    });
  });

  it('spinner shows intent text from report_intent between tools', async () => {
    // When report_intent fires, its text should appear as the spinner label
    let resolveStream!: () => void;
    const streamGate = new Promise<void>(r => {
      resolveStream = r;
    });

    const pausingSession = {
      sessionId: 'test-session-id',
      async *query() {
        yield makeEvent('tool.execution_start', {
          toolCallId: 'tc1',
          toolName: 'get_range_values',
          arguments: { address: 'A1' },
        });
        yield makeEvent('tool.execution_complete', {
          toolCallId: 'tc1',
          success: true,
          result: { content: '[[42]]' },
        });
        yield makeEvent('tool.execution_start', {
          toolCallId: 'ri1',
          toolName: 'report_intent',
          arguments: { intent: 'Formatting the table' },
        });
        // ── PAUSE ── spinner should show "Formatting the table"
        await streamGate;
        yield makeEvent('tool.execution_start', {
          toolCallId: 'tc2',
          toolName: 'set_range_format',
          arguments: { address: 'A1:D10', format: { bold: true } },
        });
        yield makeEvent('tool.execution_complete', {
          toolCallId: 'tc2',
          success: true,
          result: { content: 'OK' },
        });
        yield makeEvent('assistant.message', { messageId: 'msg1', content: 'Formatted!' });
        yield IDLE_EVENT;
      },
      on: vi.fn(),
      onPermissionRequest: vi.fn(() => () => undefined),
      destroy: vi.fn().mockResolvedValue(undefined),
      send: vi.fn().mockResolvedValue('msg-id'),
      registerTools: vi.fn(),
      getToolHandler: vi.fn(),
      respondPermission: vi.fn().mockResolvedValue(undefined),
      setModel: vi.fn().mockResolvedValue(undefined),
      compact: vi.fn().mockResolvedValue(undefined),
      _dispatchEvent: vi.fn(),
    };
    const client = makeFakeClient(pausingSession as ReturnType<typeof makeFakeSession>);
    mockCreate.mockResolvedValue(client as never);

    const { getHook } = renderThreadWithHook();

    await act(async () => {
      await new Promise(r => setTimeout(r, 100));
    });
    await act(async () => {
      void getHook().send('test');
      await new Promise(r => setTimeout(r, 200));
    });

    // Spinner should show the intent text
    const spinner = document.querySelector('[data-testid="working-spinner"]');
    expect(spinner).toBeInTheDocument();
    expect(spinner!.textContent).toContain('Formatting the table');

    await act(async () => {
      resolveStream();
      await new Promise(r => setTimeout(r, 300));
    });
  });

  it('inter-step thinking works even after text has been streamed', async () => {
    // CRITICAL regression: once text streams, the old code permanently killed
    // the thinking indicator because of `!hasText`. Now the Working box
    // spinner takes over.
    let resolveStream!: () => void;
    const streamGate = new Promise<void>(r => {
      resolveStream = r;
    });

    const pausingSession = {
      sessionId: 'test-session-id',
      async *query() {
        // Model first outputs some text
        yield makeEvent('assistant.message_delta', {
          messageId: 'msg1',
          deltaContent: 'Let me check that. ',
        });
        // Then calls a tool
        yield makeEvent('tool.execution_start', {
          toolCallId: 'tc1',
          toolName: 'get_range_values',
          arguments: { address: 'A1:B2' },
        });
        yield makeEvent('tool.execution_complete', {
          toolCallId: 'tc1',
          success: true,
          result: { content: '[[1,2],[3,4]]' },
        });
        // ── PAUSE ── inter-step thinking after text + tool
        await streamGate;
        yield makeEvent('assistant.message', {
          messageId: 'msg1',
          content: 'Let me check that. The data shows values 1-4.',
        });
        yield IDLE_EVENT;
      },
      on: vi.fn(),
      onPermissionRequest: vi.fn(() => () => undefined),
      destroy: vi.fn().mockResolvedValue(undefined),
      send: vi.fn().mockResolvedValue('msg-id'),
      registerTools: vi.fn(),
      getToolHandler: vi.fn(),
      respondPermission: vi.fn().mockResolvedValue(undefined),
      setModel: vi.fn().mockResolvedValue(undefined),
      compact: vi.fn().mockResolvedValue(undefined),
      _dispatchEvent: vi.fn(),
    };
    const client = makeFakeClient(pausingSession as ReturnType<typeof makeFakeSession>);
    mockCreate.mockResolvedValue(client as never);

    const { getHook } = renderThreadWithHook();

    await act(async () => {
      await new Promise(r => setTimeout(r, 100));
    });
    await act(async () => {
      void getHook().send('test');
      await new Promise(r => setTimeout(r, 200));
    });

    // Text has been streamed — inline shimmer is correctly hidden
    expect(document.querySelector('.inline-working-progress')).not.toBeInTheDocument();

    // But the Working box spinner MUST show (this was the bug)
    const spinner = document.querySelector('[data-testid="working-spinner"]');
    expect(spinner).toBeInTheDocument();
    expect(spinner!.textContent).toContain('Thinking');

    // Working box should still be expanded and active
    const workingTitle = document.querySelector('.chat-thinking-title-shimmer');
    expect(workingTitle).toBeInTheDocument();

    await act(async () => {
      resolveStream();
      await new Promise(r => setTimeout(r, 200));
    });

    // After completion, spinner gone
    expect(document.querySelector('[data-testid="working-spinner"]')).not.toBeInTheDocument();
  });

  it('Working box shows correct completion title and collapses after all tools finish', async () => {
    const session = makeFakeSession([
      makeEvent('tool.execution_start', {
        toolCallId: 'tc1',
        toolName: 'get_range_values',
        arguments: { address: 'A1' },
      }),
      makeEvent('tool.execution_complete', {
        toolCallId: 'tc1',
        success: true,
        result: { content: '[[1]]' },
      }),
      makeEvent('tool.execution_start', {
        toolCallId: 'tc2',
        toolName: 'set_range_values',
        arguments: { address: 'B1', values: [[2]] },
      }),
      makeEvent('tool.execution_complete', {
        toolCallId: 'tc2',
        success: true,
        result: { content: 'OK' },
      }),
      makeEvent('assistant.message', { messageId: 'msg1', content: 'Done.' }),
      IDLE_EVENT,
    ]);
    const client = makeFakeClient(session);
    mockCreate.mockResolvedValue(client as never);

    const { getHook } = renderThreadWithHook();

    await act(async () => {
      await new Promise(r => setTimeout(r, 100));
    });
    await act(async () => {
      void getHook().send('test');
      await new Promise(r => setTimeout(r, 300));
    });

    await waitFor(() => {
      expect(screen.getByText('Done.')).toBeInTheDocument();
    });

    // After completion: Working box header shows the phase label.
    // No label was set (no report_intent) → falls back to "Working"
    const doneTitle = document.querySelector('.chat-thinking-title-done');
    expect(doneTitle).toBeInTheDocument();
    expect(doneTitle!.textContent).toBe('Working');

    // Working box should auto-collapse (collapsible content hidden)
    const collapsible = document.querySelector('.chat-thinking-collapsible');
    expect(collapsible).toBeInTheDocument();
    expect((collapsible as HTMLElement).style.display).toBe('none');
  });
});

// ────────────────────────────────────────────────────────────────────────────
// Tool-call visual ordering & tearing prevention
// ────────────────────────────────────────────────────────────────────────────

describe('Tool-call visual ordering (VS Code layout: tools above text)', () => {
  beforeEach(() => {
    vi.clearAllMocks();
    useSettingsStore.getState().reset();
    useSessionHistoryStore.setState({ sessions: [], activeSessionId: null });
  });

  it('tool cards are wrapped in a ToolGroup div with order: -1', async () => {
    const session = makeFakeSession([
      makeEvent('tool.execution_start', {
        toolCallId: 'tc1',
        toolName: 'get_range_values',
        arguments: { address: 'A1' },
      }),
      makeEvent('tool.execution_complete', {
        toolCallId: 'tc1',
        success: true,
        result: { content: '[[1]]' },
      }),
      makeEvent('assistant.message', { messageId: 'msg1', content: 'Got it.' }),
      IDLE_EVENT,
    ]);
    const client = makeFakeClient(session);
    mockCreate.mockResolvedValue(client as never);

    const { getHook } = renderThreadWithHook();

    await act(async () => {
      await new Promise(r => setTimeout(r, 100));
    });
    await act(async () => {
      void getHook().send('test');
      await new Promise(r => setTimeout(r, 300));
    });

    await waitFor(() => {
      expect(screen.getByText('Got it.')).toBeInTheDocument();
    });

    // Tool card must be inside a .chat-thinking-box wrapper (Working box (all tools use it now))
    const toolGroup = document.querySelector('.chat-thinking-box');
    expect(toolGroup).toBeInTheDocument();
    expect(toolGroup!.querySelector('[data-slot="tool-fallback-root"]')).toBeInTheDocument();

    // The wrapper must have order: -1 for CSS visual reordering
    expect((toolGroup as HTMLElement).style.order).toBe('-1');
  });

  it('message content area uses flex column layout', async () => {
    const session = makeFakeSession([
      makeEvent('assistant.message', { messageId: 'msg1', content: 'Hi' }),
      IDLE_EVENT,
    ]);
    const client = makeFakeClient(session);
    mockCreate.mockResolvedValue(client as never);

    const { getHook } = renderThreadWithHook();

    await act(async () => {
      await new Promise(r => setTimeout(r, 100));
    });
    await act(async () => {
      void getHook().send('test');
      await new Promise(r => setTimeout(r, 200));
    });

    await waitFor(() => {
      expect(document.querySelector('.aui-assistant-message-content')).toBeInTheDocument();
    });

    const content = document.querySelector('.aui-assistant-message-content');
    expect(content!.classList.contains('flex')).toBe(true);
    expect(content!.classList.contains('flex-col')).toBe(true);
  });

  it('ToolGroup has order:-1 so tools render visually above text', async () => {
    const session = makeFakeSession([
      makeEvent('tool.execution_start', {
        toolCallId: 'tc1',
        toolName: 'set_range_values',
        arguments: { address: 'B1', values: [[99]] },
      }),
      makeEvent('tool.execution_complete', {
        toolCallId: 'tc1',
        success: true,
        result: { content: 'OK' },
      }),
      makeEvent('assistant.message', { messageId: 'msg1', content: 'Updated B1.' }),
      IDLE_EVENT,
    ]);
    const client = makeFakeClient(session);
    mockCreate.mockResolvedValue(client as never);

    const { getHook } = renderThreadWithHook();

    await act(async () => {
      await new Promise(r => setTimeout(r, 100));
    });
    await act(async () => {
      void getHook().send('test');
      await new Promise(r => setTimeout(r, 300));
    });

    await waitFor(() => {
      expect(screen.getByText('Updated B1.')).toBeInTheDocument();
    });

    // Verify both text and tool wrapper exist inside message content
    const content = document.querySelector('.aui-assistant-message-content');
    const toolGroup = content!.querySelector('.chat-thinking-box');
    const toolCard = toolGroup!.querySelector('[data-slot="tool-fallback-root"]');
    expect(toolGroup).toBeInTheDocument();
    expect(toolCard).toBeInTheDocument();

    // Tool wrapper has order: -1 → visually above text (order: 0 default)
    expect((toolGroup as HTMLElement).style.order).toBe('-1');
  });

  it('multiple tool cards are all inside the same ToolGroup wrapper', async () => {
    const session = makeFakeSession([
      makeEvent('tool.execution_start', {
        toolCallId: 'tc1',
        toolName: 'get_range_values',
        arguments: { address: 'A1' },
      }),
      makeEvent('tool.execution_complete', {
        toolCallId: 'tc1',
        success: true,
        result: { content: '[[1]]' },
      }),
      makeEvent('tool.execution_start', {
        toolCallId: 'tc2',
        toolName: 'set_range_values',
        arguments: { address: 'B1', values: [[2]] },
      }),
      makeEvent('tool.execution_complete', {
        toolCallId: 'tc2',
        success: true,
        result: { content: 'OK' },
      }),
      makeEvent('assistant.message', { messageId: 'msg1', content: 'Both done.' }),
      IDLE_EVENT,
    ]);
    const client = makeFakeClient(session);
    mockCreate.mockResolvedValue(client as never);

    const { getHook } = renderThreadWithHook();

    await act(async () => {
      await new Promise(r => setTimeout(r, 100));
    });
    await act(async () => {
      void getHook().send('test');
      await new Promise(r => setTimeout(r, 300));
    });

    await waitFor(() => {
      expect(screen.getByText('Both done.')).toBeInTheDocument();
    });

    const toolGroups = document.querySelectorAll('.chat-thinking-box');
    expect(toolGroups).toHaveLength(1);

    const toolCards = toolGroups[0]!.querySelectorAll('[data-slot="tool-fallback-root"]');
    expect(toolCards).toHaveLength(2);
  });

  it('no crash when tools arrive before text (tools-first event sequence)', async () => {
    let resolveStream!: () => void;
    const streamGate = new Promise<void>(r => {
      resolveStream = r;
    });

    const pausingSession = {
      sessionId: 'test-session-id',
      async *query() {
        yield makeEvent('tool.execution_start', {
          toolCallId: 'tc1',
          toolName: 'get_range_values',
          arguments: { address: 'A1:C3' },
        });
        yield makeEvent('tool.execution_complete', {
          toolCallId: 'tc1',
          success: true,
          result: { content: '[[1,2,3]]' },
        });
        // ── PAUSE ── React renders: text part (empty) + tool card
        await streamGate;
        yield makeEvent('assistant.message', { messageId: 'msg1', content: 'Here is A1:C3.' });
        yield IDLE_EVENT;
      },
      on: vi.fn(),
      onPermissionRequest: vi.fn(() => () => undefined),
      destroy: vi.fn().mockResolvedValue(undefined),
      send: vi.fn().mockResolvedValue('msg-id'),
      registerTools: vi.fn(),
      getToolHandler: vi.fn(),
      respondPermission: vi.fn().mockResolvedValue(undefined),
      _dispatchEvent: vi.fn(),
    };
    const client = makeFakeClient(pausingSession as ReturnType<typeof makeFakeSession>);
    mockCreate.mockResolvedValue(client as never);

    const { getHook } = renderThreadWithHook();

    await act(async () => {
      await new Promise(r => setTimeout(r, 100));
    });

    await act(async () => {
      void getHook().send('test');
      await new Promise(r => setTimeout(r, 150));
    });

    // Mid-stream: tool card visible, no crash
    expect(screen.queryByText(/MessagePartText can only/i)).not.toBeInTheDocument();
    expect(screen.queryByText(/something went wrong/i)).not.toBeInTheDocument();
    expect(document.querySelector('[data-slot="tool-fallback-root"]')).toBeInTheDocument();

    await act(async () => {
      resolveStream();
      await new Promise(r => setTimeout(r, 200));
    });

    await waitFor(() => {
      expect(screen.getByText('Here is A1:C3.')).toBeInTheDocument();
    });
  });

  it('no crash when text streams first then a tool call arrives (text-first tearing scenario)', async () => {
    // THIS IS THE EXACT TEARING SCENARIO from #56:
    // Text starts streaming at index 0, then a tool call arrives.
    let resolveStream!: () => void;
    const streamGate = new Promise<void>(r => {
      resolveStream = r;
    });

    const pausingSession = {
      sessionId: 'test-session-id',
      async *query() {
        yield makeEvent('assistant.message_delta', {
          deltaContent: 'Let me check ',
        } as never);
        // ── PAUSE ── React renders: text part at index 0 ("Let me check ")
        await streamGate;
        yield makeEvent('tool.execution_start', {
          toolCallId: 'tc1',
          toolName: 'get_range_values',
          arguments: { address: 'Sheet1!A1' },
        });
        yield makeEvent('tool.execution_complete', {
          toolCallId: 'tc1',
          success: true,
          result: { content: '[[42]]' },
        });
        yield makeEvent('assistant.message_delta', {
          deltaContent: 'your data.',
        } as never);
        yield IDLE_EVENT;
      },
      on: vi.fn(),
      onPermissionRequest: vi.fn(() => () => undefined),
      destroy: vi.fn().mockResolvedValue(undefined),
      send: vi.fn().mockResolvedValue('msg-id'),
      registerTools: vi.fn(),
      getToolHandler: vi.fn(),
      respondPermission: vi.fn().mockResolvedValue(undefined),
      _dispatchEvent: vi.fn(),
    };
    const client = makeFakeClient(pausingSession as ReturnType<typeof makeFakeSession>);
    mockCreate.mockResolvedValue(client as never);

    const { getHook } = renderThreadWithHook();

    await act(async () => {
      await new Promise(r => setTimeout(r, 100));
    });

    await act(async () => {
      void getHook().send('test');
      await new Promise(r => setTimeout(r, 150));
    });

    // Mid-stream: text is visible, no crash
    expect(screen.queryByText(/MessagePartText can only/i)).not.toBeInTheDocument();
    expect(screen.queryByText(/something went wrong/i)).not.toBeInTheDocument();

    await act(async () => {
      resolveStream();
      await new Promise(r => setTimeout(r, 300));
    });

    // After completion: both text and tool card visible, no crash
    expect(screen.queryByText(/MessagePartText can only/i)).not.toBeInTheDocument();
    expect(document.querySelector('[data-slot="tool-fallback-root"]')).toBeInTheDocument();
  });

  it('no crash with interleaved text → tool → text sequence', async () => {
    const session = makeFakeSession([
      makeEvent('assistant.message_delta', { deltaContent: 'Checking...' } as never),
      makeEvent('tool.execution_start', {
        toolCallId: 'tc1',
        toolName: 'get_range_values',
        arguments: { address: 'A1' },
      }),
      makeEvent('tool.execution_complete', {
        toolCallId: 'tc1',
        success: true,
        result: { content: '[[5]]' },
      }),
      makeEvent('assistant.message_delta', { deltaContent: ' Value is 5.' } as never),
      IDLE_EVENT,
    ]);
    const client = makeFakeClient(session);
    mockCreate.mockResolvedValue(client as never);

    const { getHook } = renderThreadWithHook();

    await act(async () => {
      await new Promise(r => setTimeout(r, 100));
    });
    await act(async () => {
      void getHook().send('test');
      await new Promise(r => setTimeout(r, 300));
    });

    // No crash
    expect(screen.queryByText(/MessagePartText can only/i)).not.toBeInTheDocument();
    expect(screen.queryByText(/something went wrong/i)).not.toBeInTheDocument();

    // Tool card visible
    expect(document.querySelector('[data-slot="tool-fallback-root"]')).toBeInTheDocument();
  });

  it('no crash with multiple tool calls and no text at all', async () => {
    const session = makeFakeSession([
      makeEvent('tool.execution_start', {
        toolCallId: 'tc1',
        toolName: 'manage_plugins',
        arguments: { action: 'list' },
      }),
      makeEvent('tool.execution_complete', {
        toolCallId: 'tc1',
        success: true,
        result: { content: '["skill1"]' },
      }),
      makeEvent('tool.execution_start', {
        toolCallId: 'tc2',
        toolName: 'manage_plugins',
        arguments: { action: 'browse' },
      }),
      makeEvent('tool.execution_complete', {
        toolCallId: 'tc2',
        success: true,
        result: { content: '["agent1"]' },
      }),
      IDLE_EVENT,
    ]);
    const client = makeFakeClient(session);
    mockCreate.mockResolvedValue(client as never);

    const { getHook } = renderThreadWithHook();

    await act(async () => {
      await new Promise(r => setTimeout(r, 100));
    });
    await act(async () => {
      void getHook().send('test');
      await new Promise(r => setTimeout(r, 300));
    });

    await waitFor(() => {
      expect(document.querySelector('[data-role="assistant"]')).toBeInTheDocument();
    });

    // No crash
    expect(screen.queryByText(/MessagePartText can only/i)).not.toBeInTheDocument();
    expect(screen.queryByText(/something went wrong/i)).not.toBeInTheDocument();

    // Both tool cards rendered
    const toolCards = document.querySelectorAll('[data-slot="tool-fallback-root"]');
    expect(toolCards.length).toBe(2);
  });

  it('[REGRESSION] report_intent after first tool creates a NEW Working box (per-phase split)', async () => {
    // VS Code creates a new IChatTask (Working box) per intent phase.
    // When report_intent fires AFTER at least one tool, the next tools go into a new box.
    const session = makeFakeSession([
      makeEvent('tool.execution_start', {
        toolCallId: 'tc1',
        toolName: 'get_range_values',
        arguments: { address: 'A1' },
      }),
      makeEvent('tool.execution_complete', {
        toolCallId: 'tc1',
        success: true,
        result: { content: '[[1]]' },
      }),
      // report_intent after a tool → should start phase 2 → new Working box
      makeEvent('tool.execution_start', {
        toolCallId: 'ri1',
        toolName: 'report_intent',
        arguments: { intent: 'Writing result' },
      }),
      makeEvent('tool.execution_start', {
        toolCallId: 'tc2',
        toolName: 'set_range_values',
        arguments: { address: 'B1', values: [[2]] },
      }),
      makeEvent('tool.execution_complete', {
        toolCallId: 'tc2',
        success: true,
        result: { content: 'OK' },
      }),
      makeEvent('assistant.message', { messageId: 'msg1', content: 'Done in 2 phases.' }),
      IDLE_EVENT,
    ]);
    const client = makeFakeClient(session);
    mockCreate.mockResolvedValue(client as never);

    const { getHook } = renderThreadWithHook();

    await act(async () => {
      await new Promise(r => setTimeout(r, 100));
    });
    await act(async () => {
      void getHook().send('test');
      await new Promise(r => setTimeout(r, 300));
    });

    await waitFor(() => {
      expect(screen.getByText('Done in 2 phases.')).toBeInTheDocument();
    });

    // After a report_intent following a completed tool → TWO Working boxes
    const workingBoxes = document.querySelectorAll('.chat-thinking-box');
    expect(workingBoxes).toHaveLength(2);

    // Phase 1 box: has tc1 tool card
    const box1Cards = workingBoxes[0]!.querySelectorAll('[data-slot="tool-fallback-root"]');
    expect(box1Cards).toHaveLength(1);

    // Phase 2 box: has tc2 tool card
    const box2Cards = workingBoxes[1]!.querySelectorAll('[data-slot="tool-fallback-root"]');
    expect(box2Cards).toHaveLength(1);
  });

  it('[REGRESSION] without report_intent between tools, only ONE Working box is created', async () => {
    // Tools without a phase break → all in one box
    const session = makeFakeSession([
      makeEvent('tool.execution_start', {
        toolCallId: 'tc1',
        toolName: 'get_range_values',
        arguments: { address: 'A1' },
      }),
      makeEvent('tool.execution_complete', {
        toolCallId: 'tc1',
        success: true,
        result: { content: '[[1]]' },
      }),
      makeEvent('tool.execution_start', {
        toolCallId: 'tc2',
        toolName: 'set_range_values',
        arguments: { address: 'B1', values: [[2]] },
      }),
      makeEvent('tool.execution_complete', {
        toolCallId: 'tc2',
        success: true,
        result: { content: 'OK' },
      }),
      makeEvent('assistant.message', { messageId: 'msg1', content: 'Done.' }),
      IDLE_EVENT,
    ]);
    const client = makeFakeClient(session);
    mockCreate.mockResolvedValue(client as never);

    const { getHook } = renderThreadWithHook();

    await act(async () => {
      await new Promise(r => setTimeout(r, 100));
    });
    await act(async () => {
      void getHook().send('test');
      await new Promise(r => setTimeout(r, 300));
    });

    await waitFor(() => {
      expect(screen.getByText('Done.')).toBeInTheDocument();
    });

    // No phase break → exactly ONE Working box
    const workingBoxes = document.querySelectorAll('.chat-thinking-box');
    expect(workingBoxes).toHaveLength(1);
  });

  it('[REGRESSION] Copilot avatar/name header is NOT rendered in assistant messages', async () => {
    // VS Code hides the avatar+name for GitHub Copilot (the default assistant).
    // Our AssistantMessage must NOT render the "Copilot" label header.
    // The header contains a .codicon-copilot icon — checking for that is the
    // most reliable signal since the header is the only place it appears in a message.
    const session = makeFakeSession([
      makeEvent('assistant.message', { messageId: 'msg1', content: 'Hello!' }),
      IDLE_EVENT,
    ]);
    const client = makeFakeClient(session);
    mockCreate.mockResolvedValue(client as never);

    const { getHook } = renderThreadWithHook();

    await act(async () => {
      await new Promise(r => setTimeout(r, 100));
    });
    await act(async () => {
      void getHook().send('test');
      await new Promise(r => setTimeout(r, 200));
    });

    await waitFor(() => {
      expect(document.querySelector('[data-role="assistant"]')).toBeInTheDocument();
    });

    // The Copilot avatar icon must NOT be inside the assistant message
    // (VS Code hides avatar+name for the default Copilot assistant)
    const assistantMsg = document.querySelector('[data-role="assistant"]');
    const copilotIcon = assistantMsg?.querySelector('.codicon-copilot');
    expect(copilotIcon).toBeNull();
  });

  it('after completion, tool cards remain visible inside ToolGroup above text', async () => {
    const session = makeFakeSession([
      makeEvent('tool.execution_start', {
        toolCallId: 'tc1',
        toolName: 'get_range_values',
        arguments: { address: 'A1:A5' },
      }),
      makeEvent('tool.execution_complete', {
        toolCallId: 'tc1',
        success: true,
        result: { content: '[[1],[2],[3],[4],[5]]' },
      }),
      makeEvent('assistant.message', { messageId: 'msg1', content: 'Here are rows 1-5.' }),
      IDLE_EVENT,
    ]);
    const client = makeFakeClient(session);
    mockCreate.mockResolvedValue(client as never);

    const { getHook } = renderThreadWithHook();

    await act(async () => {
      await new Promise(r => setTimeout(r, 100));
    });
    await act(async () => {
      void getHook().send('test');
      await new Promise(r => setTimeout(r, 300));
    });

    await waitFor(() => {
      expect(screen.getByText('Here are rows 1-5.')).toBeInTheDocument();
    });

    // Tool cards must still be visible after completion
    const toolGroup = document.querySelector('.chat-thinking-box');
    expect(toolGroup).toBeInTheDocument();

    const toolCard = toolGroup!.querySelector('[data-slot="tool-fallback-root"]');
    expect(toolCard).toBeInTheDocument();

    // The completed tool card should show the checkmark icon (not spinner)
    const checkIcon = toolCard!.querySelector('.codicon-check');
    expect(checkIcon).toBeInTheDocument();

    // Thinking indicator must be gone (no longer has old class — inline working progress hides when complete)
    expect(document.querySelector('.inline-working-progress')).not.toBeInTheDocument();
  });
});

// ────────────────────────────────────────────────────────────────────────────
// ChoiceCards rendering & interaction
// ────────────────────────────────────────────────────────────────────────────

describe('ChoiceCards', () => {
  beforeEach(() => {
    vi.clearAllMocks();
    useSettingsStore.getState().reset();
    useSessionHistoryStore.setState({ sessions: [], activeSessionId: null });
  });

  it('renders a numbered list of choices from a markdown choices block', async () => {
    const choicesJson = JSON.stringify([{ label: 'CSV' }, { label: 'Excel' }, { label: 'PDF' }]);
    const content = `Which format?\n\`\`\`choices\n${choicesJson}\n\`\`\``;

    const session = makeFakeSession([
      makeEvent('assistant.message', { messageId: 'msg1', content }),
      IDLE_EVENT,
    ]);
    const client = makeFakeClient(session);
    mockCreate.mockResolvedValue(client as never);

    const { getHook } = renderThreadWithHook();

    await act(async () => {
      await new Promise(r => setTimeout(r, 100));
    });
    await act(async () => {
      void getHook().send('test');
      await new Promise(r => setTimeout(r, 200));
    });

    await waitFor(() => {
      expect(document.querySelector('.aui-choices-wrapper')).toBeInTheDocument();
    });

    // All three choice labels are present
    expect(screen.getByText('CSV')).toBeInTheDocument();
    expect(screen.getByText('Excel')).toBeInTheDocument();
    expect(screen.getByText('PDF')).toBeInTheDocument();

    // Number badges 1, 2, 3 are present (now using VS Code carousel class names)
    const wrapper = document.querySelector('.aui-choices-wrapper')!;
    const badges = wrapper.querySelectorAll('.chat-question-list-number');
    expect(badges[0]?.textContent?.trim()).toBe('1');
    expect(badges[1]?.textContent?.trim()).toBe('2');
    expect(badges[2]?.textContent?.trim()).toBe('3');

    // Freeform textarea is present with the correct placeholder
    const textarea = wrapper.querySelector('textarea');
    expect(textarea).toBeInTheDocument();
    expect(textarea?.placeholder).toBe('Enter custom answer');

    // Freeform row is numbered n+1 (4)
    const freeformNumber = wrapper.querySelector('.chat-question-freeform-number');
    expect(freeformNumber?.textContent?.trim()).toBe('4');
  });

  it('clicking a choice appends the label as a user message', async () => {
    const choicesJson = JSON.stringify([{ label: 'Use CSV' }, { label: 'Use Excel' }]);
    const content = `Pick one:\n\`\`\`choices\n${choicesJson}\n\`\`\``;

    const session = makeFakeSession([
      makeEvent('assistant.message', { messageId: 'msg1', content }),
      IDLE_EVENT,
    ]);
    const client = makeFakeClient(session);
    mockCreate.mockResolvedValue(client as never);

    const { getHook } = renderThreadWithHook();

    await act(async () => {
      await new Promise(r => setTimeout(r, 100));
    });
    await act(async () => {
      void getHook().send('test');
      await new Promise(r => setTimeout(r, 200));
    });

    await waitFor(() => {
      expect(document.querySelector('.aui-choices-wrapper')).toBeInTheDocument();
    });

    // Click the first choice
    const choiceButton = screen.getByText('Use CSV').closest('button')!;
    await act(async () => {
      fireEvent.click(choiceButton);
      await new Promise(r => setTimeout(r, 100));
    });

    // A user message with that label should now appear in the thread
    await waitFor(() => {
      const userMessages = document.querySelectorAll('[data-role="user"]');
      const texts = Array.from(userMessages).map(el => el.textContent);
      expect(texts.some(t => t?.includes('Use CSV'))).toBe(true);
    });
  });

  it('submitting freeform textarea appends the typed text as a user message', async () => {
    const choicesJson = JSON.stringify([{ label: 'Option A' }]);
    const content = `\`\`\`choices\n${choicesJson}\n\`\`\``;

    const session = makeFakeSession([
      makeEvent('assistant.message', { messageId: 'msg1', content }),
      IDLE_EVENT,
    ]);
    const client = makeFakeClient(session);
    mockCreate.mockResolvedValue(client as never);

    const { getHook } = renderThreadWithHook();

    await act(async () => {
      await new Promise(r => setTimeout(r, 100));
    });
    await act(async () => {
      void getHook().send('test');
      await new Promise(r => setTimeout(r, 200));
    });

    await waitFor(() => {
      expect(document.querySelector('textarea[placeholder="Enter custom answer"]')).toBeInTheDocument();
    });

    const textarea = document.querySelector<HTMLTextAreaElement>('textarea[placeholder="Enter custom answer"]')!;

    await act(async () => {
      fireEvent.change(textarea, { target: { value: 'My custom answer' } });
      fireEvent.keyDown(textarea, { key: 'Enter', shiftKey: false });
      await new Promise(r => setTimeout(r, 100));
    });

    await waitFor(() => {
      const userMessages = document.querySelectorAll('[data-role="user"]');
      const texts = Array.from(userMessages).map(el => el.textContent);
      expect(texts.some(t => t?.includes('My custom answer'))).toBe(true);
    });
  });

  it('does not submit freeform text when Shift+Enter is pressed', async () => {
    const choicesJson = JSON.stringify([{ label: 'Option A' }]);
    const content = `\`\`\`choices\n${choicesJson}\n\`\`\``;

    const session = makeFakeSession([
      makeEvent('assistant.message', { messageId: 'msg1', content }),
      IDLE_EVENT,
    ]);
    const client = makeFakeClient(session);
    mockCreate.mockResolvedValue(client as never);

    const { getHook } = renderThreadWithHook();

    await act(async () => {
      await new Promise(r => setTimeout(r, 100));
    });
    await act(async () => {
      void getHook().send('test');
      await new Promise(r => setTimeout(r, 200));
    });

    await waitFor(() => {
      expect(document.querySelector('textarea[placeholder="Enter custom answer"]')).toBeInTheDocument();
    });

    const textarea = document.querySelector<HTMLTextAreaElement>('textarea[placeholder="Enter custom answer"]')!;
    const initialUserMessageCount = document.querySelectorAll('[data-role="user"]').length;

    await act(async () => {
      fireEvent.change(textarea, { target: { value: 'Draft text' } });
      // Shift+Enter should NOT submit
      fireEvent.keyDown(textarea, { key: 'Enter', shiftKey: true });
      await new Promise(r => setTimeout(r, 100));
    });

    // User message count should not have changed
    expect(document.querySelectorAll('[data-role="user"]').length).toBe(initialUserMessageCount);
  });
});

// ────────────────────────────────────────────────────────────────────────────
// SuggestionLinks rendering & interaction
// ────────────────────────────────────────────────────────────────────────────

describe('SuggestionLinks', () => {
  beforeEach(() => {
    vi.clearAllMocks();
    useSettingsStore.getState().reset();
    useSessionHistoryStore.setState({ sessions: [], activeSessionId: null });
  });

  it('renders follow-up suggestion links from a markdown suggestions block', async () => {
    const suggestionsJson = JSON.stringify([
      { label: 'Create a chart' },
      { label: 'Add a formula' },
    ]);
    const content = `Done! Here are some ideas:\n\`\`\`suggestions\n${suggestionsJson}\n\`\`\``;

    const session = makeFakeSession([
      makeEvent('assistant.message', { messageId: 'msg1', content }),
      IDLE_EVENT,
    ]);
    const client = makeFakeClient(session);
    mockCreate.mockResolvedValue(client as never);

    const { getHook } = renderThreadWithHook();

    await act(async () => {
      await new Promise(r => setTimeout(r, 100));
    });
    await act(async () => {
      void getHook().send('test');
      await new Promise(r => setTimeout(r, 200));
    });

    await waitFor(() => {
      expect(document.querySelector('.aui-suggestions-wrapper')).toBeInTheDocument();
    });

    expect(screen.getByText('Create a chart')).toBeInTheDocument();
    expect(screen.getByText('Add a formula')).toBeInTheDocument();

    // Should be plain link-style buttons (no border, bg-transparent)
    const wrapper = document.querySelector('.aui-suggestions-wrapper')!;
    const buttons = wrapper.querySelectorAll('button');
    expect(buttons.length).toBe(2);
  });

  it('does NOT render a freeform textarea (unlike choices)', async () => {
    const suggestionsJson = JSON.stringify([{ label: 'Try something' }]);
    const content = `\`\`\`suggestions\n${suggestionsJson}\n\`\`\``;

    const session = makeFakeSession([
      makeEvent('assistant.message', { messageId: 'msg1', content }),
      IDLE_EVENT,
    ]);
    const client = makeFakeClient(session);
    mockCreate.mockResolvedValue(client as never);

    const { getHook } = renderThreadWithHook();

    await act(async () => {
      await new Promise(r => setTimeout(r, 100));
    });
    await act(async () => {
      void getHook().send('test');
      await new Promise(r => setTimeout(r, 200));
    });

    await waitFor(() => {
      expect(document.querySelector('.aui-suggestions-wrapper')).toBeInTheDocument();
    });

    // Suggestions have no freeform textarea
    expect(document.querySelector('.aui-suggestions-wrapper textarea')).not.toBeInTheDocument();
  });

  it('clicking a suggestion appends the label as a user message', async () => {
    const suggestionsJson = JSON.stringify([{ label: 'Add a pivot table' }]);
    const content = `\`\`\`suggestions\n${suggestionsJson}\n\`\`\``;

    const session = makeFakeSession([
      makeEvent('assistant.message', { messageId: 'msg1', content }),
      IDLE_EVENT,
    ]);
    const client = makeFakeClient(session);
    mockCreate.mockResolvedValue(client as never);

    const { getHook } = renderThreadWithHook();

    await act(async () => {
      await new Promise(r => setTimeout(r, 100));
    });
    await act(async () => {
      void getHook().send('test');
      await new Promise(r => setTimeout(r, 200));
    });

    await waitFor(() => {
      expect(document.querySelector('.aui-suggestions-wrapper')).toBeInTheDocument();
    });

    const suggestionButton = screen.getByText('Add a pivot table').closest('button')!;
    await act(async () => {
      fireEvent.click(suggestionButton);
      await new Promise(r => setTimeout(r, 100));
    });

    await waitFor(() => {
      const userMessages = document.querySelectorAll('[data-role="user"]');
      const texts = Array.from(userMessages).map(el => el.textContent);
      expect(texts.some(t => t?.includes('Add a pivot table'))).toBe(true);
    });
  });
});

// ────────────────────────────────────────────────────────────────────────────
// Thread UI — Reload / BranchPicker / Edit buttons
// ────────────────────────────────────────────────────────────────────────────

describe('Thread UI: Reload, BranchPicker, and Edit buttons', () => {
  beforeEach(() => {
    vi.clearAllMocks();
    useSettingsStore.getState().reset();
    useSessionHistoryStore.setState({ sessions: [], activeSessionId: null });
  });

  it('renders "Regenerate response" button in the assistant action bar', async () => {
    const session = makeFakeSession([
      makeEvent('assistant.message', { messageId: 'msg1', content: 'Hello world.' }),
      IDLE_EVENT,
    ]);
    const client = makeFakeClient(session);
    mockCreate.mockResolvedValue(client as never);

    const { getHook } = renderThreadWithHook();

    await act(async () => {
      await new Promise(r => setTimeout(r, 100));
    });
    await act(async () => {
      void getHook().send('test');
      await new Promise(r => setTimeout(r, 200));
    });

    await waitFor(() => {
      expect(document.querySelector('[data-role="assistant"]')).toBeInTheDocument();
    });

    // The action bar should exist
    expect(document.querySelector('.aui-assistant-action-bar')).toBeInTheDocument();

    // TooltipIconButton renders accessible name via sr-only span — use getByRole
    const reloadBtn = screen.getByRole('button', { name: 'Regenerate response' });
    expect(reloadBtn).toBeInTheDocument();
  });

  it('does NOT render branch navigation buttons when there is only one branch', async () => {
    const session = makeFakeSession([
      makeEvent('assistant.message', { messageId: 'msg1', content: 'Single response.' }),
      IDLE_EVENT,
    ]);
    const client = makeFakeClient(session);
    mockCreate.mockResolvedValue(client as never);

    const { getHook } = renderThreadWithHook();

    await act(async () => {
      await new Promise(r => setTimeout(r, 100));
    });
    await act(async () => {
      void getHook().send('test');
      await new Promise(r => setTimeout(r, 200));
    });

    await waitFor(() => {
      expect(document.querySelector('[data-role="assistant"]')).toBeInTheDocument();
    });

    // BranchPickerPrimitive has hideWhenSingleBranch — prev/next should not be in DOM
    const prevBtn = document.querySelector('button[title="Previous response"]');
    const nextBtn = document.querySelector('button[title="Next response"]');
    expect(prevBtn).not.toBeInTheDocument();
    expect(nextBtn).not.toBeInTheDocument();
  });

  it('renders "Edit message" button in the user message action bar', async () => {
    // Use a full session so the thread properly idles (hideWhenRunning needs the run to complete)
    const session = makeFakeSession([
      makeEvent('assistant.message', { messageId: 'msg1', content: 'Hello!' }),
      IDLE_EVENT,
    ]);
    const client = makeFakeClient(session);
    mockCreate.mockResolvedValue(client as never);

    const { getHook } = renderThreadWithHook();

    await act(async () => {
      await new Promise(r => setTimeout(r, 100));
    });
    await act(async () => {
      void getHook().send('test');
      await new Promise(r => setTimeout(r, 300));
    });

    await waitFor(() => {
      expect(document.querySelector('[data-role="user"]')).toBeInTheDocument();
    });

    // ActionBarPrimitive.Root with autohide="always" requires isHovering=true to render.
    // AUI uses native addEventListener("mouseenter"), so fireEvent.mouseEnter triggers it.
    const userMsg = document.querySelector('[data-role="user"]')!;
    await act(async () => {
      fireEvent.mouseEnter(userMsg);
      await new Promise(r => setTimeout(r, 50));
    });

    // TooltipIconButton renders accessible name via sr-only span — use getByRole
    await waitFor(() => {
      expect(screen.getByRole('button', { name: 'Edit message' })).toBeInTheDocument();
    });
  });

  it('user action bar has aui-user-action-bar class', async () => {
    // Use a full session so the thread properly idles (hideWhenRunning needs the run to complete)
    const session = makeFakeSession([
      makeEvent('assistant.message', { messageId: 'msg1', content: 'Hello!' }),
      IDLE_EVENT,
    ]);
    const client = makeFakeClient(session);
    mockCreate.mockResolvedValue(client as never);

    const { getHook } = renderThreadWithHook();

    await act(async () => {
      await new Promise(r => setTimeout(r, 100));
    });
    await act(async () => {
      void getHook().send('test');
      await new Promise(r => setTimeout(r, 300));
    });

    await waitFor(() => {
      expect(document.querySelector('[data-role="user"]')).toBeInTheDocument();
    });

    // Trigger hover so autohide="always" allows rendering
    const userMsg = document.querySelector('[data-role="user"]')!;
    await act(async () => {
      fireEvent.mouseEnter(userMsg);
      await new Promise(r => setTimeout(r, 50));
    });

    await waitFor(() => {
      expect(document.querySelector('.aui-user-action-bar')).toBeInTheDocument();
    });
  });
});

// ────────────────────────────────────────────────────────────────────────────
// task_complete summary rendering + multi-turn Working box isolation
// ────────────────────────────────────────────────────────────────────────────

/**
 * Multi-turn session: each call to query() yields events from the next turn.
 * Ensures turn 2 tools never bleed into turn 1's message.
 */
function makeMultiTurnSession(turnEvents: SessionEvent[][]) {
  let callCount = 0;
  return {
    sessionId: 'test-session-id',
    async *query() {
      const events = turnEvents[callCount] ?? turnEvents[turnEvents.length - 1];
      callCount++;
      for (const event of events) {
        yield event;
        if (event.type === 'session.idle') return;
      }
    },
    on: vi.fn(),
    onPermissionRequest: vi.fn(() => () => undefined),
    destroy: vi.fn().mockResolvedValue(undefined),
    send: vi.fn().mockResolvedValue('msg-id'),
    registerTools: vi.fn(),
    getToolHandler: vi.fn(),
    respondPermission: vi.fn().mockResolvedValue(undefined),
    _dispatchEvent: vi.fn() as EventEmitter,
  };
}

describe('task_complete: summary rendering and multi-turn Working box isolation', () => {
  beforeEach(() => {
    vi.clearAllMocks();
    useSettingsStore.getState().reset();
    useSessionHistoryStore.setState({ sessions: [], activeSessionId: null });
  });

  it('surfaces task_complete summary as visible text when no text response follows', async () => {
    const session = makeFakeSession([
      makeEvent('tool.execution_start', {
        toolCallId: 'tc1',
        toolName: 'get_range_values',
        arguments: { address: 'A1' },
      }),
      makeEvent('tool.execution_complete', {
        toolCallId: 'tc1',
        success: true,
        result: { content: '[[42]]' },
      }),
      makeEvent('tool.execution_start', {
        toolCallId: 'tc2',
        toolName: 'task_complete',
        arguments: { summary: 'Read 1 cell with value 42.' },
      }),
      makeEvent('tool.execution_complete', {
        toolCallId: 'tc2',
        success: true,
        result: { content: 'Done' },
      }),
      IDLE_EVENT,
    ]);
    const client = makeFakeClient(session);
    mockCreate.mockResolvedValue(client as never);

    const { getHook } = renderThreadWithHook();
    await act(async () => { await new Promise(r => setTimeout(r, 100)); });
    await act(async () => {
      void getHook().send('test');
      await new Promise(r => setTimeout(r, 300));
    });

    // Summary text must be visible — no expanding required
    await waitFor(() => {
      expect(screen.getByText('Read 1 cell with value 42.')).toBeInTheDocument();
    });

    // VS Code: completed box shows the phase label (IChatTask.content).
    // No report_intent in this test → falls back to "Working"
    const doneTitle = document.querySelector('.chat-thinking-title-done');
    expect(doneTitle?.textContent).toBe('Working');
  });

  it('task_complete does NOT create an extra Working box — only real work tools shown', async () => {
    const session = makeFakeSession([
      makeEvent('tool.execution_start', { toolCallId: 'tc1', toolName: 'get_range_values', arguments: {} }),
      makeEvent('tool.execution_complete', { toolCallId: 'tc1', success: true, result: { content: '[[1]]' } }),
      makeEvent('tool.execution_start', { toolCallId: 'tc2', toolName: 'set_range_values', arguments: {} }),
      makeEvent('tool.execution_complete', { toolCallId: 'tc2', success: true, result: { content: 'OK' } }),
      makeEvent('tool.execution_start', { toolCallId: 'tc3', toolName: 'task_complete', arguments: { summary: 'Done.' } }),
      makeEvent('tool.execution_complete', { toolCallId: 'tc3', success: true, result: { content: 'Done' } }),
      IDLE_EVENT,
    ]);
    const client = makeFakeClient(session);
    mockCreate.mockResolvedValue(client as never);

    const { getHook } = renderThreadWithHook();
    await act(async () => { await new Promise(r => setTimeout(r, 100)); });
    await act(async () => {
      void getHook().send('test');
      await new Promise(r => setTimeout(r, 300));
    });

    await waitFor(() => { expect(screen.getByText('Done.')).toBeInTheDocument(); });

    // VS Code: completed box shows phase label — no label here → "Working"
    const doneTitle = document.querySelector('.chat-thinking-title-done');
    expect(doneTitle?.textContent).toBe('Working');
  });

  it('REGRESSION: turn 2 Working box appears in a new message, not inside turn 1 message', async () => {
    const session = makeMultiTurnSession([
      // Turn 1: one work step + task_complete
      [
        makeEvent('tool.execution_start', { toolCallId: 'tc1', toolName: 'get_range_values', arguments: {} }),
        makeEvent('tool.execution_complete', { toolCallId: 'tc1', success: true, result: { content: '[[1]]' } }),
        makeEvent('tool.execution_start', { toolCallId: 'tc2', toolName: 'task_complete', arguments: { summary: 'Turn 1 done.' } }),
        makeEvent('tool.execution_complete', { toolCallId: 'tc2', success: true, result: { content: 'Done' } }),
        IDLE_EVENT,
      ],
      // Turn 2: starts a new working session
      [
        makeEvent('tool.execution_start', { toolCallId: 'tc3', toolName: 'set_range_values', arguments: {} }),
        makeEvent('tool.execution_complete', { toolCallId: 'tc3', success: true, result: { content: 'OK' } }),
        IDLE_EVENT,
      ],
    ]);
    const client = makeFakeClient(session);
    mockCreate.mockResolvedValue(client as never);

    const { getHook } = renderThreadWithHook();
    await act(async () => { await new Promise(r => setTimeout(r, 100)); });

    // Turn 1
    await act(async () => {
      void getHook().send('do task 1');
      await new Promise(r => setTimeout(r, 300));
    });
    await waitFor(() => { expect(screen.getByText('Turn 1 done.')).toBeInTheDocument(); });

    // Exactly 1 assistant message after turn 1
    expect(document.querySelectorAll('[data-role="assistant"]')).toHaveLength(1);

    // Turn 2
    await act(async () => {
      void getHook().send('do task 2');
      await new Promise(r => setTimeout(r, 300));
    });
    await waitFor(() => {
      expect(document.querySelectorAll('[data-role="assistant"]')).toHaveLength(2);
    });

    const assistantMsgs = document.querySelectorAll('[data-role="assistant"]');

    // Each turn must have exactly ONE Working box — never shared
    expect(assistantMsgs[0].querySelectorAll('.chat-thinking-box')).toHaveLength(1);
    expect(assistantMsgs[1].querySelectorAll('.chat-thinking-box')).toHaveLength(1);

    // Turn 1's box must be collapsed (auto-collapsed on completion)
    expect(assistantMsgs[0].querySelector('.chat-thinking-collapsed')).toBeInTheDocument();

    // Turn 1's summary text must still be visible in its own message
    expect(assistantMsgs[0].textContent).toContain('Turn 1 done.');

    // Turn 2's Working box must be in turn 2's message (not turn 1's)
    expect(assistantMsgs[1].querySelector('.chat-thinking-box')).toBeInTheDocument();
    expect(assistantMsgs[0].textContent).not.toContain('Set range values');
  });
});
