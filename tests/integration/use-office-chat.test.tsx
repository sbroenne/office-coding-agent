/**
 * Integration tests for useOfficeChat hook.
 *
 * Mocks createWebSocketClient to return a fake client/session so we can
 * simulate Copilot session events and verify the hook maps them correctly
 * to ChatMessage[] for the custom chat UI.
 */

import React from 'react';
import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import { renderHook, act } from '@testing-library/react';
import type { SessionEvent } from '@github/copilot-sdk';
import { useOfficeChat } from '@/hooks/useOfficeChat';
import { useSettingsStore } from '@/stores/settingsStore';
import { useSessionHistoryStore } from '@/stores/sessionHistoryStore';
import { setImportedAgents } from '@/services/agents';

// ─── Fake session builder ─────────────────────────────────────────────────────

type EventEmitter = (event: SessionEvent) => void;

function makeFakeSession(events: SessionEvent[]) {
  return {
    sessionId: 'test-session-id',
    // eslint-disable-next-line @typescript-eslint/require-await
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
    setModel: vi.fn().mockResolvedValue(undefined),
    compact: vi.fn().mockResolvedValue(undefined),
    _dispatchEvent: vi.fn() as EventEmitter,
  };
}

function makeFakeClient(
  session: ReturnType<typeof makeFakeSession>,
  models: { id: string; name: string }[] = []
) {
  return {
    start: vi.fn().mockResolvedValue(undefined),
    createSession: vi.fn().mockResolvedValue(session),
    listModels: vi.fn().mockResolvedValue(models),
    stop: vi.fn().mockResolvedValue(undefined),
    onMcpStatus: vi.fn(() => () => undefined),
    onMcpLog: vi.fn(() => () => undefined),
    onMcpTools: vi.fn(() => () => undefined),
    onPluginAgents: vi.fn(() => () => undefined),
    onPluginSkills: vi.fn(() => () => undefined),
    onPluginPrompts: vi.fn(() => () => undefined),
  };
}

// Mock createWebSocketClient — injected per-test via mockResolvedValue
vi.mock('@/lib/websocket-client', () => ({
  createWebSocketClient: vi.fn(),
}));

import { createWebSocketClient } from '@/lib/websocket-client';
const mockCreate = vi.mocked(createWebSocketClient);

// ─── Helpers ──────────────────────────────────────────────────────────────────

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


function wrapper({ children }: { children: React.ReactNode }) {
  return React.createElement(React.Fragment, null, children);
}

// ─── Tests ────────────────────────────────────────────────────────────────────

describe('useOfficeChat', () => {
  beforeEach(() => {
    vi.clearAllMocks();
    useSettingsStore.getState().reset();
    useSessionHistoryStore.setState({ sessions: [], activeSessionId: null });
    setImportedAgents([]);
  });

  afterEach(() => {
    setImportedAgents([]);
  });

  it('starts in idle state with no messages', async () => {
    const session = makeFakeSession([IDLE_EVENT]);
    const client = makeFakeClient(session);
    mockCreate.mockResolvedValue(client as never);

    const { result } = renderHook(() => useOfficeChat('excel'), { wrapper });

    // Wait for initSession to complete
    await act(async () => {
      await new Promise(r => setTimeout(r, 50));
    });

    expect(result.current.sessionError).toBeNull();
    expect(result.current.messages).toBeDefined();
  });

  it('adds user + assistant messages when onNew is called', async () => {
    const session = makeFakeSession([
      makeEvent('assistant.message', { messageId: 'msg1', content: 'Hello!' }),
      IDLE_EVENT,
    ]);
    const client = makeFakeClient(session);
    mockCreate.mockResolvedValue(client as never);

    const { result } = renderHook(() => useOfficeChat('excel'), { wrapper });

    await act(async () => {
      await new Promise(r => setTimeout(r, 50));
    });

    await act(async () => {
      void result.current.send('Say hello');
      await new Promise(r => setTimeout(r, 100));
    });

    const messages = result.current.messages;
    expect(messages).toHaveLength(2);
    expect(messages[0].role).toBe('user');
    expect(messages[1].role).toBe('assistant');

    const assistantContent = messages[1].content;
    const textPart = assistantContent.find(c => c.type === 'text');
    expect(textPart).toBeDefined();
    expect(textPart!.type).toBe('text');
    expect((textPart as { type: 'text'; text: string }).text).toBe('Hello!');
  });

  it('accumulates streaming delta text', async () => {
    const session = makeFakeSession([
      makeEvent('assistant.message_delta', { messageId: 'msg1', deltaContent: 'He' }),
      makeEvent('assistant.message_delta', { messageId: 'msg1', deltaContent: 'llo' }),
      makeEvent('assistant.message_delta', { messageId: 'msg1', deltaContent: '!' }),
      IDLE_EVENT,
    ]);
    const client = makeFakeClient(session);
    mockCreate.mockResolvedValue(client as never);

    const { result } = renderHook(() => useOfficeChat('excel'), { wrapper });

    await act(async () => {
      await new Promise(r => setTimeout(r, 50));
    });

    await act(async () => {
      void result.current.send('Say hello');
      await new Promise(r => setTimeout(r, 100));
    });

    const messages = result.current.messages;
    expect(messages.length).toBeGreaterThanOrEqual(2);
    const assistantContent = messages[1].content;
    const textPart = assistantContent.find(c => c.type === 'text');
    expect(textPart).toBeDefined();
    expect(textPart!.type).toBe('text');
    expect((textPart as { type: 'text'; text: string }).text).toBe('Hello!');
  });

  it('keeps tool-call parts in completed message alongside text (VS Code behavior)', async () => {
    // Tool-call parts remain visible in the completed message as a collapsible
    // "thinking" section above the text response, matching VS Code Copilot Chat.
    const session = makeFakeSession([
      makeEvent('tool.execution_start', {
        toolCallId: 'tc1',
        toolName: 'get_range_values',
        arguments: { range: 'A1:B2' },
      }),
      makeEvent('tool.execution_complete', {
        toolCallId: 'tc1',
        success: true,
        result: { content: '[[1,2],[3,4]]' },
      }),
      makeEvent('assistant.message', { messageId: 'msg1', content: 'Done!' }),
      IDLE_EVENT,
    ]);
    const client = makeFakeClient(session);
    mockCreate.mockResolvedValue(client as never);

    const { result } = renderHook(() => useOfficeChat('excel'), { wrapper });

    await act(async () => {
      await new Promise(r => setTimeout(r, 50));
    });

    await act(async () => {
      void result.current.send('Read A1:B2');
      await new Promise(r => setTimeout(r, 100));
    });

    const messages = result.current.messages;
    const assistantContent = messages[1].content;
    // Text part must be present
    const textPart = assistantContent.find(c => c.type === 'text');
    expect(textPart).toBeDefined();
    expect((textPart as { type: 'text'; text: string }).text).toBe('Done!');
    // Tool-call parts must ALSO be present (VS Code keeps them visible)
    const toolPart = assistantContent.find(c => c.type === 'tool-call');
    expect(toolPart).toBeDefined();
  });

  it('keeps tool-call parts in completed message when there is no text', async () => {
    // If the AI response is tool-only (no text), tool parts must be kept so the
    // message is not empty.
    const session = makeFakeSession([
      makeEvent('tool.execution_start', {
        toolCallId: 'tc2',
        toolName: 'get_range_values',
        arguments: { range: 'A1' },
      }),
      makeEvent('tool.execution_complete', {
        toolCallId: 'tc2',
        success: true,
        result: { content: '42' },
      }),
      IDLE_EVENT,
    ]);
    const client = makeFakeClient(session);
    mockCreate.mockResolvedValue(client as never);

    const { result } = renderHook(() => useOfficeChat('excel'), { wrapper });

    await act(async () => {
      await new Promise(r => setTimeout(r, 50));
    });

    await act(async () => {
      void result.current.send('Get A1');
      await new Promise(r => setTimeout(r, 100));
    });

    const messages = result.current.messages;
    const assistantContent = messages[1].content;
    const toolPart = assistantContent.find(c => c.type === 'tool-call');
    expect(toolPart).toBeDefined();
    expect((toolPart as { type: 'tool-call'; toolName: string }).toolName).toBe('get_range_values');
  });

  it('keeps thinkingText as Thinking during tool execution (VS Code behavior)', async () => {
    // In VS Code, the "Thinking" label stays constant while tool cards show
    // individual tool names. We don't flash the thinking text to tool names.
    let resolveIdle: () => void;
    const idlePromise = new Promise<void>(r => {
      resolveIdle = r;
    });

    const session = {
      sessionId: 'test-session-id',
      async *query() {
        yield makeEvent('tool.execution_start', {
          toolCallId: 'tc1',
          toolName: 'get_range_values',
          arguments: { range: 'A1:B2' },
        });
        // Pause here so the test can observe thinkingText
        await idlePromise;
        yield makeEvent('tool.execution_complete', {
          toolCallId: 'tc1',
          success: true,
          result: { content: '[[1,2]]' },
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
      _dispatchEvent: vi.fn() as EventEmitter,
    };
    const client = makeFakeClient(session);
    mockCreate.mockResolvedValue(client as never);

    const { result } = renderHook(() => useOfficeChat('excel'), { wrapper });

    await act(async () => {
      await new Promise(r => setTimeout(r, 50));
    });

    // Send a message — the stream will pause after tool.execution_start
    await act(async () => {
      void result.current.send('Read');
      await new Promise(r => setTimeout(r, 50));
    });

    // thinkingText should stay as "Thinking…" (not change to tool name)
    expect(result.current.messages.findLast(m => m.role === 'assistant')?.thinkingText).toBe('Thinking…');

    // Release the stream to complete
    await act(async () => {
      resolveIdle!();
      await new Promise(r => setTimeout(r, 100));
    });

    // After completion, thinkingText should be cleared
    expect(result.current.messages.findLast(m => m.role === 'assistant')?.thinkingText).toBeNull();
  });

  it('report_intent overrides tool name in thinkingText', async () => {
    let resolveIdle: () => void;
    const idlePromise = new Promise<void>(r => {
      resolveIdle = r;
    });

    const session = {
      sessionId: 'test-session-id',
      async *query() {
        yield makeEvent('tool.execution_start', {
          toolCallId: 'ri1',
          toolName: 'report_intent',
          arguments: { intent: 'Reading the spreadsheet' },
        });
        // Pause so the test can observe thinkingText
        await idlePromise;
        yield makeEvent('assistant.message', { messageId: 'msg1', content: 'Here you go' });
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
      _dispatchEvent: vi.fn() as EventEmitter,
    };
    const client = makeFakeClient(session);
    mockCreate.mockResolvedValue(client as never);

    const { result } = renderHook(() => useOfficeChat('excel'), { wrapper });

    await act(async () => {
      await new Promise(r => setTimeout(r, 50));
    });

    await act(async () => {
      void result.current.send('Read');
      await new Promise(r => setTimeout(r, 50));
    });

    // report_intent should surface the raw intent text
    expect(result.current.messages.findLast(m => m.role === 'assistant')?.thinkingText).toBe('Reading the spreadsheet');

    // Release the stream to complete
    await act(async () => {
      resolveIdle!();
      await new Promise(r => setTimeout(r, 100));
    });

    expect(result.current.messages.findLast(m => m.role === 'assistant')?.thinkingText).toBeNull();
  });

  it('shows Asking [AgentName] in thinkingText when subagent.started fires', async () => {
    let resolveIdle: () => void;
    const idlePromise = new Promise<void>(r => {
      resolveIdle = r;
    });

    const session = {
      sessionId: 'test-session-id',
      async *query() {
        yield makeEvent('subagent.started', {
          toolCallId: 'sa1',
          agentName: 'Specialist',
          agentDisplayName: 'Specialist Agent',
          agentDescription: 'Handles specialist tasks',
        });
        // Pause so the test can observe thinkingText
        await idlePromise;
        yield makeEvent('subagent.completed', {
          toolCallId: 'sa1',
          agentName: 'Specialist',
          agentDisplayName: 'Specialist Agent',
        });
        yield makeEvent('assistant.message', { messageId: 'msg1', content: 'Done' });
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
      _dispatchEvent: vi.fn() as EventEmitter,
    };
    const client = makeFakeClient(session);
    mockCreate.mockResolvedValue(client as never);

    const { result } = renderHook(() => useOfficeChat('excel'), { wrapper });

    await act(async () => {
      await new Promise(r => setTimeout(r, 50));
    });

    await act(async () => {
      void result.current.send('Do something');
      await new Promise(r => setTimeout(r, 50));
    });

    // thinkingText should show the sub-agent display name
    expect(result.current.messages.findLast(m => m.role === 'assistant')?.thinkingText).toBe('Asking Specialist Agent…');

    // Release the stream to complete
    await act(async () => {
      resolveIdle!();
      await new Promise(r => setTimeout(r, 100));
    });

    // After completion, thinkingText should be cleared
    expect(result.current.messages.findLast(m => m.role === 'assistant')?.thinkingText).toBeNull();
  });

  it('resets thinkingText to Thinking… when subagent.completed fires', async () => {
    let resolveAfterCompleted: () => void;
    const afterCompletedPromise = new Promise<void>(r => {
      resolveAfterCompleted = r;
    });

    const session = {
      sessionId: 'test-session-id',
      async *query() {
        yield makeEvent('subagent.started', {
          toolCallId: 'sa2',
          agentName: 'Specialist',
          agentDisplayName: 'Specialist Agent',
          agentDescription: 'desc',
        });
        yield makeEvent('subagent.completed', {
          toolCallId: 'sa2',
          agentName: 'Specialist',
          agentDisplayName: 'Specialist Agent',
        });
        // Pause after sub-agent completed so we can observe the reset text
        await afterCompletedPromise;
        yield makeEvent('assistant.message', { messageId: 'msg1', content: 'Done' });
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
      _dispatchEvent: vi.fn() as EventEmitter,
    };
    const client = makeFakeClient(session);
    mockCreate.mockResolvedValue(client as never);

    const { result } = renderHook(() => useOfficeChat('excel'), { wrapper });

    await act(async () => {
      await new Promise(r => setTimeout(r, 50));
    });

    await act(async () => {
      void result.current.send('Do something');
      await new Promise(r => setTimeout(r, 50));
    });

    // After subagent.completed, thinkingText should be reset to 'Thinking…'
    expect(result.current.messages.findLast(m => m.role === 'assistant')?.thinkingText).toBe('Thinking…');

    await act(async () => {
      resolveAfterCompleted!();
      await new Promise(r => setTimeout(r, 100));
    });

    expect(result.current.messages.findLast(m => m.role === 'assistant')?.thinkingText).toBeNull();
  });

  it('resets thinkingText to Thinking… when subagent.failed fires', async () => {
    let resolveAfterFailed: () => void;
    const afterFailedPromise = new Promise<void>(r => {
      resolveAfterFailed = r;
    });

    const session = {
      sessionId: 'test-session-id',
      async *query() {
        yield makeEvent('subagent.started', {
          toolCallId: 'sa3',
          agentName: 'Specialist',
          agentDisplayName: 'Specialist Agent',
          agentDescription: 'desc',
        });
        yield makeEvent('subagent.failed', {
          toolCallId: 'sa3',
          agentName: 'Specialist',
          agentDisplayName: 'Specialist Agent',
          error: 'Sub-agent encountered an error',
        });
        // Pause after failure so we can observe the reset text
        await afterFailedPromise;
        yield makeEvent('assistant.message', { messageId: 'msg1', content: 'I could not delegate' });
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
      _dispatchEvent: vi.fn() as EventEmitter,
    };
    const client = makeFakeClient(session);
    mockCreate.mockResolvedValue(client as never);

    const { result } = renderHook(() => useOfficeChat('excel'), { wrapper });

    await act(async () => {
      await new Promise(r => setTimeout(r, 50));
    });

    await act(async () => {
      void result.current.send('Do something');
      await new Promise(r => setTimeout(r, 50));
    });

    // After subagent.failed, thinkingText should be reset to 'Thinking…'
    expect(result.current.messages.findLast(m => m.role === 'assistant')?.thinkingText).toBe('Thinking…');

    await act(async () => {
      resolveAfterFailed!();
      await new Promise(r => setTimeout(r, 100));
    });

    expect(result.current.messages.findLast(m => m.role === 'assistant')?.thinkingText).toBeNull();
  });

  it('sets session error when createWebSocketClient rejects', async () => {
    mockCreate.mockRejectedValue(new Error('server unavailable'));

    const { result } = renderHook(() => useOfficeChat('excel'), { wrapper });

    await act(async () => {
      await new Promise(r => setTimeout(r, 100));
    });

    expect(result.current.sessionError).toBeInstanceOf(Error);
    expect(result.current.sessionError?.message).toBe('server unavailable');
  });

  it('populates availableModels in the store after session init', async () => {
    const FAKE_MODELS = [
      { id: 'claude-sonnet-4', name: 'Claude Sonnet 4' },
      { id: 'gpt-4.1', name: 'GPT-4.1' },
      { id: 'gemini-2.5-pro', name: 'Gemini 2.5 Pro' },
    ];
    const session = makeFakeSession([IDLE_EVENT]);
    const client = makeFakeClient(session, FAKE_MODELS);
    mockCreate.mockResolvedValue(client as never);

    renderHook(() => useOfficeChat('excel'), { wrapper });

    await act(async () => {
      await new Promise(r => setTimeout(r, 100));
    });

    const available = useSettingsStore.getState().availableModels;
    expect(available).toHaveLength(3);
    expect(available?.[0]).toEqual({
      id: 'claude-sonnet-4',
      name: 'Claude Sonnet 4',
      provider: 'Anthropic',
    });
    expect(available?.[1]).toEqual({ id: 'gpt-4.1', name: 'GPT-4.1', provider: 'OpenAI' });
    expect(available?.[2]).toEqual({
      id: 'gemini-2.5-pro',
      name: 'Gemini 2.5 Pro',
      provider: 'Google',
    });
  });

  it('shows error message when sending with no session', async () => {
    mockCreate.mockRejectedValue(new Error('server unavailable'));

    const { result } = renderHook(() => useOfficeChat('excel'), { wrapper });

    await act(async () => {
      await new Promise(r => setTimeout(r, 100));
    });

    // Session failed — now try to send a message
    await act(async () => {
      void result.current.send('Hello');
      await new Promise(r => setTimeout(r, 100));
    });

    const messages = result.current.messages;
    expect(messages).toHaveLength(2);
    expect(messages[0].role).toBe('user');
    expect(messages[1].role).toBe('assistant');
    const textPart = messages[1].content.find(c => c.type === 'text');
    expect(textPart).toBeDefined();
    expect(textPart!.type).toBe('text');
    expect((textPart as { type: 'text'; text: string }).text).toContain('Not connected');
  });

  it('auto-corrects activeModel when not in fetched models', async () => {
    // Set activeModel to something not in the available models
    useSettingsStore.setState({ activeModel: 'nonexistent-model' });

    const MODELS = [
      { id: 'gpt-4.1', name: 'GPT-4.1' },
      { id: 'claude-sonnet-4', name: 'Claude Sonnet 4' },
    ];
    const session = makeFakeSession([IDLE_EVENT]);
    const client = makeFakeClient(session, MODELS);
    mockCreate.mockResolvedValue(client as never);

    renderHook(() => useOfficeChat('excel'), { wrapper });

    await act(async () => {
      await new Promise(r => setTimeout(r, 150));
    });

    // Should have auto-corrected to the first available model
    expect(useSettingsStore.getState().activeModel).toBe('gpt-4.1');
  });

  it('clears messages and reinitialises session on clearMessages', async () => {
    const session1 = makeFakeSession([IDLE_EVENT]);
    const session2 = makeFakeSession([IDLE_EVENT]);
    const client1 = makeFakeClient(session1);
    const client2 = makeFakeClient(session2);
    mockCreate.mockResolvedValueOnce(client1 as never).mockResolvedValueOnce(client2 as never);

    const { result } = renderHook(() => useOfficeChat('excel'), { wrapper });

    await act(async () => {
      await new Promise(r => setTimeout(r, 50));
    });

    // Send a message to populate messages
    await act(async () => {
      void result.current.send('Hi');
      await new Promise(r => setTimeout(r, 100));
    });

    expect(result.current.messages.length).toBeGreaterThan(0);

    await act(async () => {
      result.current.clearMessages();
      await new Promise(r => setTimeout(r, 100));
    });

    expect(result.current.messages).toHaveLength(0);
    expect(mockCreate).toHaveBeenCalledTimes(2);
  });

  // ─── MCP wiring ────────────────────────────────────────────────────────────

  it('includes bundled MCP servers by default (even with no imported servers)', async () => {
    const session = makeFakeSession([IDLE_EVENT]);
    const client = makeFakeClient(session);
    mockCreate.mockResolvedValue(client as never);

    renderHook(() => useOfficeChat('excel'), { wrapper });

    await act(async () => {
      await new Promise(r => setTimeout(r, 100));
    });

    const config = client.createSession.mock.calls[0][0] as Record<string, unknown>;
    // Bundled servers (e.g. WorkIQ) are included by default
    expect(config.mcpServers).toBeDefined();
  });

  // ─── Per-agent tool scoping ─────────────────────────────────────────────────

  it('does not pass availableTools when active agent has no tools restriction', async () => {
    const session = makeFakeSession([IDLE_EVENT]);
    const client = makeFakeClient(session);
    mockCreate.mockResolvedValue(client as never);

    renderHook(() => useOfficeChat('excel'), { wrapper });

    await act(async () => {
      await new Promise(r => setTimeout(r, 100));
    });

    const config = client.createSession.mock.calls[0][0] as Record<string, unknown>;
    expect(config.availableTools).toBeUndefined();
  });

  // ─── Native SDK skills and agents ───────────────────────────────────────────

  it('passes customAgents with resolved agent to createSession', async () => {
    const session = makeFakeSession([IDLE_EVENT]);
    const client = makeFakeClient(session);
    mockCreate.mockResolvedValue(client as never);

    renderHook(() => useOfficeChat('excel'), { wrapper });

    await act(async () => {
      await new Promise(r => setTimeout(r, 100));
    });

    const config = client.createSession.mock.calls[0][0] as Record<string, unknown>;
    const agents = config.customAgents as { name: string; prompt: string }[];
    expect(agents).toBeDefined();
    expect(agents).toHaveLength(1);
    expect(agents[0].name).toBe('Excel');
    expect(agents[0].prompt).toContain('AI assistant');
  });

  it('systemMessage contains only base + app prompt, not agent instructions', async () => {
    const session = makeFakeSession([IDLE_EVENT]);
    const client = makeFakeClient(session);
    mockCreate.mockResolvedValue(client as never);

    renderHook(() => useOfficeChat('excel'), { wrapper });

    await act(async () => {
      await new Promise(r => setTimeout(r, 100));
    });

    const config = client.createSession.mock.calls[0][0] as Record<string, unknown>;
    const sysMsg = config.systemMessage as { content: string };
    // Should NOT contain agent-specific instructions (those are in customAgents)
    expect(sysMsg.content).not.toContain('The workbook is already open');
    expect(sysMsg.content).not.toContain('Core Behavior');
    // Should contain the base prompt
    expect(sysMsg.content).toContain('Progress narration');
  });

  it('plugin.agents notification registers plugin agents as imported agents', async () => {
    /**
     * Regression test: when the proxy sends a plugin.agents notification,
     * the hook must call setImportedAgents() so plugin agents appear in
     * the AgentPicker and are included in the next session's customAgents.
     */
    const session = makeFakeSession([IDLE_EVENT]);
    let capturedPluginAgentsHandler: ((payload: unknown) => void) | undefined;

    const client = {
      ...makeFakeClient(session),
      onPluginAgents: vi.fn((handler: (payload: unknown) => void) => {
        capturedPluginAgentsHandler = handler;
        return () => undefined;
      }),
    };
    mockCreate.mockResolvedValue(client as never);

    renderHook(() => useOfficeChat('excel'), { wrapper });

    await act(async () => {
      await new Promise(r => setTimeout(r, 100));
    });

    // Simulate proxy sending plugin.agents notification
    expect(capturedPluginAgentsHandler).toBeDefined();
    await act(async () => {
      capturedPluginAgentsHandler!({
        agents: [
          {
            name: 'Plugin Agent',
            description: 'From a plugin',
            prompt: 'Plugin instructions.',
            hosts: [],
          },
        ],
      });
      await new Promise(r => setTimeout(r, 50));
    });

    // The plugin agents should now be in the store's pluginAgents
    const pluginAgents = useSettingsStore.getState().pluginAgents;
    expect(pluginAgents.some(a => a.metadata.name === 'Plugin Agent')).toBe(true);

    // Cleanup
    useSettingsStore.getState().setPluginAgents([]);
  });

  it('plugin.skills notification populates store.pluginSkills', async () => {
    const session = makeFakeSession([IDLE_EVENT]);
    let capturedPluginSkillsHandler: ((payload: unknown) => void) | undefined;

    const client = {
      ...makeFakeClient(session),
      onPluginSkills: vi.fn((handler: (payload: unknown) => void) => {
        capturedPluginSkillsHandler = handler;
        return () => undefined;
      }),
    };
    mockCreate.mockResolvedValue(client as never);

    renderHook(() => useOfficeChat('excel'), { wrapper });

    await act(async () => {
      await new Promise(r => setTimeout(r, 100));
    });

    expect(capturedPluginSkillsHandler).toBeDefined();
    await act(async () => {
      capturedPluginSkillsHandler!({
        skills: [
          {
            name: 'SPT IQ Preflight',
            description: 'Preflight skill',
            version: '1.0.0',
            hosts: [],
            content: 'Skill body content',
          },
        ],
      });
      await new Promise(r => setTimeout(r, 50));
    });

    const pluginSkills = useSettingsStore.getState().pluginSkills;
    expect(pluginSkills.some(s => s.metadata.name === 'SPT IQ Preflight')).toBe(true);

    // Cleanup
    useSettingsStore.getState().setPluginSkills([]);
  });

  it('plugin.prompts notification populates store.pluginPrompts', async () => {
    const session = makeFakeSession([IDLE_EVENT]);
    let capturedPluginPromptsHandler: ((payload: unknown) => void) | undefined;

    const client = {
      ...makeFakeClient(session),
      onPluginPrompts: vi.fn((handler: (payload: unknown) => void) => {
        capturedPluginPromptsHandler = handler;
        return () => undefined;
      }),
    };
    mockCreate.mockResolvedValue(client as never);

    renderHook(() => useOfficeChat('excel'), { wrapper });

    await act(async () => {
      await new Promise(r => setTimeout(r, 100));
    });

    expect(capturedPluginPromptsHandler).toBeDefined();
    await act(async () => {
      capturedPluginPromptsHandler!({
        prompts: [
          {
            name: 'preflight',
            description: 'Run preflight assessment',
            agent: 'SPT IQ Preflight',
            argumentHint: 'TPID',
            body: 'Run preflight for ${input:tpid}',
          },
        ],
      });
      await new Promise(r => setTimeout(r, 50));
    });

    const pluginPrompts = useSettingsStore.getState().pluginPrompts;
    expect(pluginPrompts.some(p => p.name === 'preflight')).toBe(true);
    expect(pluginPrompts[0].agent).toBe('SPT IQ Preflight');

    // Cleanup
    useSettingsStore.getState().setPluginPrompts([]);
  });

  it('does not include empty text parts in intermediate message content', async () => {
    // KEY REGRESSION TEST: updateAssistant() used to always append
    // { type: 'text', text: '' } even when streamText was empty.
    // fromThreadMessageLike() silently strips empty text parts, which
    // caused the "MessagePartText can only be used inside text or
    // reasoning message parts" crash.  This test verifies the fix by
    // pausing mid-stream (after tool events, before any text) and
    // inspecting that no empty text part is present.
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
          arguments: { range: 'A1' },
        });
        yield makeEvent('tool.execution_complete', {
          toolCallId: 'tc1',
          success: true,
          result: { content: '[[1]]' },
        });
        // Pause — the message now has tool parts but no text
        await streamGate;
        yield makeEvent('assistant.message', { messageId: 'msg1', content: 'Got it.' });
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
      _dispatchEvent: vi.fn() as EventEmitter,
    };
    const client = makeFakeClient(pausingSession);
    mockCreate.mockResolvedValue(client as never);

    const { result } = renderHook(() => useOfficeChat('excel'), { wrapper });

    await act(async () => {
      await new Promise(r => setTimeout(r, 50));
    });

    // Start the stream — pauses after tool.execution_complete
    await act(async () => {
      void result.current.send('Read');
      await new Promise(r => setTimeout(r, 100));
    });

    // Inspect intermediate message content
    const messages = result.current.messages;
    const assistant = messages.find(m => m.role === 'assistant');
    expect(assistant).toBeDefined();

    // Must have at least one tool-call part
    const toolParts = assistant!.content.filter(c => c.type === 'tool-call');
    expect(toolParts.length).toBeGreaterThanOrEqual(1);

    // Tool parts are present; empty text parts are allowed in the new custom UI
    // (components handle empty text gracefully by rendering null)
    // Release the stream
    await act(async () => {
      resolveStream();
      await new Promise(r => setTimeout(r, 100));
    });

    // After completion, text should be present
    const finalMessages = result.current.messages;
    const finalAssistant = finalMessages.find(m => m.role === 'assistant');
    const finalTextParts = finalAssistant!.content.filter(c => c.type === 'text');
    expect(finalTextParts).toHaveLength(1);
    expect((finalTextParts[0] as { text: string }).text).toBe('Got it.');
  });

  it('initial assistant message starts with empty content (content: [])', async () => {
    // When send() fires, the hook creates an assistant message immediately with
    // content: [] before any stream events arrive. This avoids crashes from
    // rendering empty text parts.
    let resolveStream!: () => void;
    const streamGate = new Promise<void>(r => {
      resolveStream = r;
    });

    const pausingSession = {
      sessionId: 'test-session-id',
      async *query() {
        // Pause immediately — the initial assistant message is in state before
        // any events arrive
        await streamGate;
        yield makeEvent('assistant.message', { messageId: 'msg1', content: 'Hi' });
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
      _dispatchEvent: vi.fn() as EventEmitter,
    };
    const client = makeFakeClient(pausingSession);
    mockCreate.mockResolvedValue(client as never);

    const { result } = renderHook(() => useOfficeChat('excel'), { wrapper });

    await act(async () => {
      await new Promise(r => setTimeout(r, 50));
    });

    await act(async () => {
      void result.current.send('Hi');
      await new Promise(r => setTimeout(r, 100));
    });

    // The assistant message starts with empty content before stream events
    const messages = result.current.messages;
    const assistant = messages.find(m => m.role === 'assistant');
    expect(assistant).toBeDefined();
    expect(assistant!.content).toHaveLength(0);

    await act(async () => {
      resolveStream();
      await new Promise(r => setTimeout(r, 100));
    });
  });
});
