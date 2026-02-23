/**
 * Integration tests for useOfficeChat hook.
 *
 * Mocks createWebSocketClient to return a fake client/session so we can
 * simulate Copilot session events and verify the hook maps them correctly
 * to ThreadMessageLike[] for assistant-ui.
 */

import React from 'react';
import { describe, it, expect, vi, beforeEach } from 'vitest';
import { renderHook, act } from '@testing-library/react';
import type { AppendMessage } from '@assistant-ui/react';
import type { SessionEvent } from '@github/copilot-sdk';
import { useOfficeChat } from '@/hooks/useOfficeChat';
import { useSettingsStore } from '@/stores/settingsStore';

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
    destroy: vi.fn().mockResolvedValue(undefined),
    send: vi.fn().mockResolvedValue('msg-id'),
    registerTools: vi.fn(),
    getToolHandler: vi.fn(),
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

const APPEND_MSG = (text: string): AppendMessage => ({
  parentId: null,
  sourceId: null,
  runConfig: undefined,
  role: 'user',
  content: [{ type: 'text', text }],
  attachments: [],
  metadata: { custom: {} },
  createdAt: new Date(),
});

function wrapper({ children }: { children: React.ReactNode }) {
  return React.createElement(React.Fragment, null, children);
}

// ─── Tests ────────────────────────────────────────────────────────────────────

describe('useOfficeChat', () => {
  beforeEach(() => {
    vi.clearAllMocks();
    useSettingsStore.getState().reset();
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
    expect(result.current.runtime).toBeTruthy();
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
      result.current.runtime.thread.append(APPEND_MSG('Say hello'));
      await new Promise(r => setTimeout(r, 100));
    });

    const messages = result.current.runtime.thread.getState().messages;
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
      result.current.runtime.thread.append(APPEND_MSG('Say hello'));
      await new Promise(r => setTimeout(r, 100));
    });

    const messages = result.current.runtime.thread.getState().messages;
    expect(messages.length).toBeGreaterThanOrEqual(2);
    const assistantContent = messages[1].content;
    const textPart = assistantContent.find(c => c.type === 'text');
    expect(textPart).toBeDefined();
    expect(textPart!.type).toBe('text');
    expect((textPart as { type: 'text'; text: string }).text).toBe('Hello!');
  });

  it('includes tool-call parts when tool events fire', async () => {
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
      result.current.runtime.thread.append(APPEND_MSG('Read A1:B2'));
      await new Promise(r => setTimeout(r, 100));
    });

    const messages = result.current.runtime.thread.getState().messages;
    const assistantContent = messages[1].content;
    const toolPart = assistantContent.find(c => c.type === 'tool-call');
    expect(toolPart).toBeDefined();
    expect(toolPart!.type).toBe('tool-call');
    expect((toolPart as { type: 'tool-call'; toolName: string }).toolName).toBe('get_range_values');
  });

  it('sets thinkingText to humanized tool name on tool.execution_start', async () => {
    // Use a slow session so we can observe thinkingText mid-stream
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
        yield makeEvent('assistant.message', { messageId: 'msg1', content: 'Done' });
        yield IDLE_EVENT;
      },
      on: vi.fn(),
      destroy: vi.fn().mockResolvedValue(undefined),
      send: vi.fn().mockResolvedValue('msg-id'),
      registerTools: vi.fn(),
      getToolHandler: vi.fn(),
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
      result.current.runtime.thread.append(APPEND_MSG('Read'));
      await new Promise(r => setTimeout(r, 50));
    });

    // thinkingText should show humanized tool name
    expect(result.current.thinkingText).toBe('Get range values…');

    // Release the stream to complete
    await act(async () => {
      resolveIdle!();
      await new Promise(r => setTimeout(r, 100));
    });

    // After completion, thinkingText should be cleared
    expect(result.current.thinkingText).toBeNull();
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
      destroy: vi.fn().mockResolvedValue(undefined),
      send: vi.fn().mockResolvedValue('msg-id'),
      registerTools: vi.fn(),
      getToolHandler: vi.fn(),
      _dispatchEvent: vi.fn() as EventEmitter,
    };
    const client = makeFakeClient(session);
    mockCreate.mockResolvedValue(client as never);

    const { result } = renderHook(() => useOfficeChat('excel'), { wrapper });

    await act(async () => {
      await new Promise(r => setTimeout(r, 50));
    });

    await act(async () => {
      result.current.runtime.thread.append(APPEND_MSG('Read'));
      await new Promise(r => setTimeout(r, 50));
    });

    // report_intent should surface the raw intent text
    expect(result.current.thinkingText).toBe('Reading the spreadsheet');

    // Release the stream to complete
    await act(async () => {
      resolveIdle!();
      await new Promise(r => setTimeout(r, 100));
    });

    expect(result.current.thinkingText).toBeNull();
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
      result.current.runtime.thread.append(APPEND_MSG('Hello'));
      await new Promise(r => setTimeout(r, 100));
    });

    const messages = result.current.runtime.thread.getState().messages;
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
      result.current.runtime.thread.append(APPEND_MSG('Hi'));
      await new Promise(r => setTimeout(r, 100));
    });

    expect(result.current.runtime.thread.getState().messages.length).toBeGreaterThan(0);

    await act(async () => {
      result.current.clearMessages();
      await new Promise(r => setTimeout(r, 100));
    });

    expect(result.current.runtime.thread.getState().messages).toHaveLength(0);
    expect(mockCreate).toHaveBeenCalledTimes(2);
  });

  // ─── MCP wiring ────────────────────────────────────────────────────────────

  it('passes mcpServers to createSession when servers are active in the store', async () => {
    useSettingsStore.getState().importMcpServers([
      { name: 'my-server', url: 'https://example.com/mcp', transport: 'http' },
    ]);
    // activeMcpServerNames null = all enabled

    const session = makeFakeSession([IDLE_EVENT]);
    const client = makeFakeClient(session);
    mockCreate.mockResolvedValue(client as never);

    renderHook(() => useOfficeChat('excel'), { wrapper });

    await act(async () => {
      await new Promise(r => setTimeout(r, 100));
    });

    const config = client.createSession.mock.calls[0][0] as Record<string, unknown>;
    expect(config.mcpServers).toBeDefined();
    expect(config.mcpServers).toHaveProperty('my-server');
    expect((config.mcpServers as Record<string, unknown>)['my-server']).toMatchObject({
      url: 'https://example.com/mcp',
      tools: ['*'],
    });
  });

  it('does not pass mcpServers when no MCP servers are imported', async () => {
    const session = makeFakeSession([IDLE_EVENT]);
    const client = makeFakeClient(session);
    mockCreate.mockResolvedValue(client as never);

    renderHook(() => useOfficeChat('excel'), { wrapper });

    await act(async () => {
      await new Promise(r => setTimeout(r, 100));
    });

    const config = client.createSession.mock.calls[0][0] as Record<string, unknown>;
    expect(config.mcpServers).toBeUndefined();
  });

  it('does not pass mcpServers when all servers are toggled off', async () => {
    useSettingsStore.getState().importMcpServers([
      { name: 'srv', url: 'https://example.com/mcp', transport: 'http' },
    ]);
    useSettingsStore.setState({ activeMcpServerNames: [] });

    const session = makeFakeSession([IDLE_EVENT]);
    const client = makeFakeClient(session);
    mockCreate.mockResolvedValue(client as never);

    renderHook(() => useOfficeChat('excel'), { wrapper });

    await act(async () => {
      await new Promise(r => setTimeout(r, 100));
    });

    const config = client.createSession.mock.calls[0][0] as Record<string, unknown>;
    expect(config.mcpServers).toBeUndefined();
  });

  it('SSE server is mapped with type:sse in mcpServers config', async () => {
    useSettingsStore.getState().importMcpServers([
      { name: 'sse-srv', url: 'https://sse.example.com', transport: 'sse' },
    ]);

    const session = makeFakeSession([IDLE_EVENT]);
    const client = makeFakeClient(session);
    mockCreate.mockResolvedValue(client as never);

    renderHook(() => useOfficeChat('excel'), { wrapper });

    await act(async () => {
      await new Promise(r => setTimeout(r, 100));
    });

    const config = client.createSession.mock.calls[0][0] as Record<string, unknown>;
    const servers = config.mcpServers as Record<string, { type: string }>;
    expect(servers['sse-srv'].type).toBe('sse');
  });

  // ─── Per-agent tool scoping ─────────────────────────────────────────────────

  it('passes availableTools to createSession when active agent specifies tools', async () => {
    useSettingsStore.getState().importAgents([
      {
        metadata: {
          name: 'Scoped',
          description: 'desc',
          version: '1.0.0',
          hosts: ['excel'],
          defaultForHosts: [],
          tools: ['create_chart', 'format_range'],
        },
        instructions: 'Use only these tools.',
      },
    ]);
    useSettingsStore.getState().setActiveAgent('Scoped');

    const session = makeFakeSession([IDLE_EVENT]);
    const client = makeFakeClient(session);
    mockCreate.mockResolvedValue(client as never);

    renderHook(() => useOfficeChat('excel'), { wrapper });

    await act(async () => {
      await new Promise(r => setTimeout(r, 100));
    });

    const config = client.createSession.mock.calls[0][0] as Record<string, unknown>;
    expect(config.availableTools).toEqual(['create_chart', 'format_range']);
  });

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

  it("agent's mcpServers allowlist filters active servers to only permitted ones", async () => {
    useSettingsStore.getState().importMcpServers([
      { name: 'allowed', url: 'https://allowed.com/mcp', transport: 'http' },
      { name: 'blocked', url: 'https://blocked.com/mcp', transport: 'http' },
    ]);
    useSettingsStore.getState().importAgents([
      {
        metadata: {
          name: 'Filtered',
          description: 'desc',
          version: '1.0.0',
          hosts: ['excel'],
          defaultForHosts: [],
          mcpServers: ['allowed'],
        },
        instructions: '',
      },
    ]);
    useSettingsStore.getState().setActiveAgent('Filtered');

    const session = makeFakeSession([IDLE_EVENT]);
    const client = makeFakeClient(session);
    mockCreate.mockResolvedValue(client as never);

    renderHook(() => useOfficeChat('excel'), { wrapper });

    await act(async () => {
      await new Promise(r => setTimeout(r, 100));
    });

    const config = client.createSession.mock.calls[0][0] as Record<string, unknown>;
    const serverKeys = Object.keys(config.mcpServers as object);
    expect(serverKeys).toContain('allowed');
    expect(serverKeys).not.toContain('blocked');
  });

  it('agent with empty mcpServers allowlist blocks all MCP servers', async () => {
    useSettingsStore.getState().importMcpServers([
      { name: 'srv', url: 'https://srv.com/mcp', transport: 'http' },
    ]);
    useSettingsStore.getState().importAgents([
      {
        metadata: {
          name: 'NoMcp',
          description: 'desc',
          version: '1.0.0',
          hosts: ['excel'],
          defaultForHosts: [],
          mcpServers: [],
        },
        instructions: '',
      },
    ]);
    useSettingsStore.getState().setActiveAgent('NoMcp');

    const session = makeFakeSession([IDLE_EVENT]);
    const client = makeFakeClient(session);
    mockCreate.mockResolvedValue(client as never);

    renderHook(() => useOfficeChat('excel'), { wrapper });

    await act(async () => {
      await new Promise(r => setTimeout(r, 100));
    });

    const config = client.createSession.mock.calls[0][0] as Record<string, unknown>;
    // empty allowlist → no servers should be forwarded
    expect(config.mcpServers).toBeUndefined();
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
    const agents = config.customAgents as Array<{ name: string; prompt: string }>;
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

  it('passes imported skills in the skills array', async () => {
    useSettingsStore.getState().importSkills([
      {
        metadata: {
          name: 'TestSkill',
          description: 'A test skill',
          version: '1.0.0',
          hosts: ['excel'],
          tags: [],
        },
        content: 'Test skill instructions.',
      },
    ]);

    const session = makeFakeSession([IDLE_EVENT]);
    const client = makeFakeClient(session);
    mockCreate.mockResolvedValue(client as never);

    renderHook(() => useOfficeChat('excel'), { wrapper });

    await act(async () => {
      await new Promise(r => setTimeout(r, 100));
    });

    const config = client.createSession.mock.calls[0][0] as Record<string, unknown>;
    const skills = config.skills as Array<{ name: string; content: string }>;
    expect(skills).toBeDefined();
    expect(skills.some(s => s.name === 'TestSkill')).toBe(true);
  });

  it('computes disabledSkills from activeSkillNames', async () => {
    // Toggle off a skill by setting activeSkillNames to exclude it
    useSettingsStore.getState().toggleSkill('excel');

    const session = makeFakeSession([IDLE_EVENT]);
    const client = makeFakeClient(session);
    mockCreate.mockResolvedValue(client as never);

    renderHook(() => useOfficeChat('excel'), { wrapper });

    await act(async () => {
      await new Promise(r => setTimeout(r, 100));
    });

    const config = client.createSession.mock.calls[0][0] as Record<string, unknown>;
    const disabled = config.disabledSkills as string[];
    expect(disabled).toContain('excel');
  });
});
