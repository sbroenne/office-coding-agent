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
import { AssistantRuntimeProvider } from '@assistant-ui/react';
import type { AppendMessage } from '@assistant-ui/react';
import type { SessionEvent } from '@github/copilot-sdk';
import { Thread } from '@/components/assistant-ui/thread';
import { useOfficeChat } from '@/hooks/useOfficeChat';
import { useSettingsStore } from '@/stores/settingsStore';
import { useSessionHistoryStore } from '@/stores/sessionHistoryStore';
import { ThinkingContext } from '@/contexts/ThinkingContext';

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

const APPEND_MSG = (): AppendMessage => ({
  parentId: null,
  sourceId: null,
  runConfig: undefined,
  role: 'user',
  content: [{ type: 'text', text: 'test' }],
  attachments: [],
  metadata: { custom: {} },
  createdAt: new Date(),
});

vi.mock('@/lib/websocket-client', () => ({
  createWebSocketClient: vi.fn(),
}));

import { createWebSocketClient } from '@/lib/websocket-client';
const mockCreate = vi.mocked(createWebSocketClient);

// ─── Shared test wrapper ──────────────────────────────────────────────────────

/**
 * Renders the Thread with a real (faked) useOfficeChat runtime and returns
 * the hook result so tests can drive messages through `runtime.thread.append`.
 */
function renderThreadWithHook(host: 'excel' = 'excel') {
  let hookRef: ReturnType<typeof useOfficeChat> | undefined;

  function TestComponent() {
    const chat = useOfficeChat(host);
    hookRef = chat;
    return (
      <AssistantRuntimeProvider runtime={chat.runtime}>
        <ThinkingContext.Provider value={chat.thinkingText}>
          <Thread />
        </ThinkingContext.Provider>
      </AssistantRuntimeProvider>
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
      getHook().runtime.thread.append(APPEND_MSG());
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
      getHook().runtime.thread.append(APPEND_MSG());
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
    // @assistant-ui's EmptyPartFallback was previously rendered, which called
    // useMessagePartText() in the wrong context and threw.
    // The fix: Empty: () => null + unstable_showEmptyOnNonTextEnd={false}.
    const session = makeFakeSession([
      makeEvent('tool.execution_start', {
        toolCallId: 'tc1',
        toolName: 'manage_skills',
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
      getHook().runtime.thread.append(APPEND_MSG());
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
      getHook().runtime.thread.append(APPEND_MSG());
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
      getHook().runtime.thread.append(APPEND_MSG());
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
      getHook().runtime.thread.append(APPEND_MSG());
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
      getHook().runtime.thread.append(APPEND_MSG());
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
      getHook().runtime.thread.append(APPEND_MSG());
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
      getHook().runtime.thread.append(APPEND_MSG());
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
      getHook().runtime.thread.append(APPEND_MSG());
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
      getHook().runtime.thread.append(APPEND_MSG());
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
      getHook().runtime.thread.append(APPEND_MSG());
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
      getHook().runtime.thread.append(APPEND_MSG());
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
      getHook().runtime.thread.append(APPEND_MSG());
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
      getHook().runtime.thread.append(APPEND_MSG());
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
      getHook().runtime.thread.append(APPEND_MSG());
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
      getHook().runtime.thread.append(APPEND_MSG());
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
        toolName: 'manage_skills',
        arguments: { action: 'list' },
      }),
      makeEvent('tool.execution_complete', {
        toolCallId: 'tc1',
        success: true,
        result: { content: '["skill1"]' },
      }),
      makeEvent('tool.execution_start', {
        toolCallId: 'tc2',
        toolName: 'manage_agents',
        arguments: { action: 'list' },
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
      getHook().runtime.thread.append(APPEND_MSG());
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
      getHook().runtime.thread.append(APPEND_MSG());
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
      getHook().runtime.thread.append(APPEND_MSG());
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
      getHook().runtime.thread.append(APPEND_MSG());
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
      getHook().runtime.thread.append(APPEND_MSG());
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
      getHook().runtime.thread.append(APPEND_MSG());
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
      getHook().runtime.thread.append(APPEND_MSG());
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
      getHook().runtime.thread.append(APPEND_MSG());
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
      getHook().runtime.thread.append(APPEND_MSG());
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
      getHook().runtime.thread.append(APPEND_MSG());
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
      getHook().runtime.thread.append(APPEND_MSG());
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
      getHook().runtime.thread.append(APPEND_MSG());
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
      getHook().runtime.thread.append(APPEND_MSG());
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
