/**
 * Unit tests for useOfficeChat hook.
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
  models: Array<{ id: string; name: string }> = []
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
    mockCreate.mockResolvedValue(client as ReturnType<typeof makeFakeClient> as never);

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
      await result.current.runtime.thread.append(APPEND_MSG('Say hello'));
      await new Promise(r => setTimeout(r, 100));
    });

    const messages = result.current.runtime.thread.getState().messages;
    expect(messages).toHaveLength(2);
    expect(messages[0].role).toBe('user');
    expect(messages[1].role).toBe('assistant');

    const assistantContent = messages[1].content;
    const textPart = assistantContent.find(c => c.type === 'text');
    expect(textPart).toBeTruthy();
    if (textPart?.type === 'text') {
      expect(textPart.text).toBe('Hello!');
    }
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
      await result.current.runtime.thread.append(APPEND_MSG('Say hello'));
      await new Promise(r => setTimeout(r, 100));
    });

    const messages = result.current.runtime.thread.getState().messages;
    const assistantContent = messages[1].content;
    const textPart = assistantContent.find(c => c.type === 'text');
    if (textPart?.type === 'text') {
      expect(textPart.text).toBe('Hello!');
    }
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
      await result.current.runtime.thread.append(APPEND_MSG('Read A1:B2'));
      await new Promise(r => setTimeout(r, 100));
    });

    const messages = result.current.runtime.thread.getState().messages;
    const assistantContent = messages[1].content;
    const toolPart = assistantContent.find(c => c.type === 'tool-call');
    expect(toolPart).toBeTruthy();
    if (toolPart?.type === 'tool-call') {
      expect(toolPart.toolName).toBe('get_range_values');
    }
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
    expect(available![0]).toEqual({
      id: 'claude-sonnet-4',
      name: 'Claude Sonnet 4',
      provider: 'Anthropic',
    });
    expect(available![1]).toEqual({ id: 'gpt-4.1', name: 'GPT-4.1', provider: 'OpenAI' });
    expect(available![2]).toEqual({
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
      await result.current.runtime.thread.append(APPEND_MSG('Hello'));
      await new Promise(r => setTimeout(r, 100));
    });

    const messages = result.current.runtime.thread.getState().messages;
    expect(messages).toHaveLength(2);
    expect(messages[0].role).toBe('user');
    expect(messages[1].role).toBe('assistant');
    const textPart = messages[1].content.find(c => c.type === 'text');
    if (textPart?.type === 'text') {
      expect(textPart.text).toContain('Not connected');
    }
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
      await result.current.runtime.thread.append(APPEND_MSG('Hi'));
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
});
