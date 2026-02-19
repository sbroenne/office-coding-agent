import { useState, useRef, useCallback, useEffect } from 'react';
import { useExternalStoreRuntime } from '@assistant-ui/react';
import type { ThreadMessageLike, AppendMessage } from '@assistant-ui/react';
import type { WebSocketCopilotClient, BrowserCopilotSession } from '@/lib/websocket-client';
import { createWebSocketClient } from '@/lib/websocket-client';
import { getToolsForHost } from '@/tools';
import { buildSkillContext } from '@/services/skills';
import { resolveActiveAgent } from '@/services/agents';
import { useSettingsStore } from '@/stores';
import { buildSystemPrompt } from '@/services/ai/systemPrompt';
import type { OfficeHostApp } from '@/services/office/host';
import { generateId } from '@/utils/id';

function getWsUrl(): string {
  if (typeof window === 'undefined') return 'wss://localhost:3000/api/copilot';
  const proto = window.location.protocol === 'https:' ? 'wss:' : 'ws:';
  return `${proto}//${window.location.host}/api/copilot`;
}

export function useOfficeChat(host: OfficeHostApp) {
  const activeModel = useSettingsStore(s => s.activeModel);
  const activeSkillNames = useSettingsStore(s => s.activeSkillNames);
  const activeAgentId = useSettingsStore(s => s.activeAgentId);

  const clientRef = useRef<WebSocketCopilotClient | null>(null);
  const sessionRef = useRef<BrowserCopilotSession | null>(null);
  const cancelRef = useRef(false);

  const [messages, setMessages] = useState<ThreadMessageLike[]>([]);
  const [isRunning, setIsRunning] = useState(false);
  const [sessionError, setSessionError] = useState<Error | null>(null);

  const initSession = useCallback(async () => {
    if (clientRef.current) {
      try {
        await clientRef.current.stop();
      } catch {
        /* ignore */
      }
      clientRef.current = null;
      sessionRef.current = null;
    }

    try {
      const client = await createWebSocketClient(getWsUrl());
      clientRef.current = client;

      const resolvedAgent = resolveActiveAgent(activeAgentId, host);
      const agentInstructions = resolvedAgent?.instructions ?? '';
      const skillContext = buildSkillContext(activeSkillNames ?? undefined);
      const systemContent = `${buildSystemPrompt(host)}\n\n${agentInstructions}${skillContext}`;

      const session = await client.createSession({
        model: activeModel,
        systemMessage: { mode: 'replace', content: systemContent },
        tools: getToolsForHost(host),
      });
      sessionRef.current = session;
      setSessionError(null);
    } catch (err) {
      setSessionError(err instanceof Error ? err : new Error(String(err)));
    }
  }, [activeModel, host, activeSkillNames, activeAgentId]);

  useEffect(() => {
    void initSession();
    return () => {
      const client = clientRef.current;
      if (client) {
        void client.stop().catch(_err => undefined);
        clientRef.current = null;
        sessionRef.current = null;
      }
    };
  }, [initSession]);

  const onNew = useCallback(async (message: AppendMessage) => {
    const userText = (message.content as readonly { type: string; text?: string }[])
      .filter(c => c.type === 'text' && typeof c.text === 'string')
      .map(c => c.text!)
      .join('\n');

    if (!userText.trim() || !sessionRef.current) return;

    const assistantId = generateId();
    cancelRef.current = false;

    const userMsg: ThreadMessageLike = {
      id: generateId(),
      role: 'user',
      content: [{ type: 'text', text: userText }],
      createdAt: new Date(),
    };

    const assistantMsg: ThreadMessageLike = {
      id: assistantId,
      role: 'assistant',
      content: [{ type: 'text', text: '' }],
      status: { type: 'running' },
      createdAt: new Date(),
    };

    setMessages(prev => [...prev, userMsg, assistantMsg]);
    setIsRunning(true);

    const toolParts = new Map<
      string,
      {
        type: 'tool-call';
        toolCallId: string;
        toolName: string;
        argsText: string;
        result?: unknown;
      }
    >();
    let streamText = '';

    const updateAssistant = (extra?: Partial<Pick<ThreadMessageLike, 'status'>>) => {
      const content: ThreadMessageLike['content'] = [
        ...Array.from(toolParts.values()),
        { type: 'text', text: streamText },
      ];
      setMessages(prev => prev.map(m => (m.id === assistantId ? { ...m, content, ...extra } : m)));
    };

    try {
      const session = sessionRef.current;
      for await (const event of session.query({ prompt: userText })) {
        if (cancelRef.current) break;

        if (event.type === 'assistant.message_delta') {
          streamText += event.data.deltaContent;
          updateAssistant();
        } else if (event.type === 'tool.execution_start') {
          const { toolCallId, toolName, arguments: args } = event.data;
          toolParts.set(toolCallId, {
            type: 'tool-call',
            toolCallId,
            toolName,
            argsText: JSON.stringify(args ?? {}),
          });
          updateAssistant();
        } else if (event.type === 'tool.execution_complete') {
          const { toolCallId, result } = event.data;
          const existing = toolParts.get(toolCallId);
          if (existing) {
            const resultText = result
              ? typeof result.content === 'string'
                ? result.content
                : JSON.stringify(result)
              : '';
            toolParts.set(toolCallId, { ...existing, result: resultText });
            updateAssistant();
          }
        } else if (event.type === 'assistant.message') {
          streamText = event.data.content;
          updateAssistant({ status: { type: 'complete', reason: 'stop' } });
        } else if (event.type === 'session.error') {
          updateAssistant({
            status: { type: 'incomplete', reason: 'error', error: event.data.message },
          });
          break;
        }
      }
    } catch (err) {
      const errMsg = err instanceof Error ? err.message : String(err);
      setMessages(prev =>
        prev.map(m =>
          m.id === assistantId
            ? { ...m, status: { type: 'incomplete', reason: 'error', error: errMsg } }
            : m
        )
      );
    } finally {
      setIsRunning(false);
    }
  }, []);

  const clearMessages = useCallback(() => {
    setMessages([]);
    void initSession();
  }, [initSession]);

  const runtime = useExternalStoreRuntime<ThreadMessageLike>({
    isRunning,
    messages,
    onNew,
    onCancel: () => {
      cancelRef.current = true;
      setIsRunning(false);
      return Promise.resolve();
    },
    convertMessage: (msg: ThreadMessageLike) => msg,
  });

  return { runtime, sessionError, clearMessages };
}
