import { useState, useRef, useCallback, useEffect } from 'react';
import { flushSync } from 'react-dom';
import { useExternalStoreRuntime } from '@assistant-ui/react';
import type { ThreadMessageLike, AppendMessage } from '@assistant-ui/react';
import type { WebSocketCopilotClient, BrowserCopilotSession } from '@/lib/websocket-client';
import { createWebSocketClient } from '@/lib/websocket-client';
import { getToolsForHost } from '@/tools';
import { getSkills, getImportedSkills } from '@/services/skills';
import { resolveActiveAgent } from '@/services/agents';
import { resolveActiveMcpServers, toSdkMcpServers } from '@/services/mcp';
import { useSettingsStore } from '@/stores';
import { buildSystemPrompt } from '@/services/ai/systemPrompt';
import { humanizeToolName } from '@/utils/humanizeToolName';
import { inferProvider } from '@/types';
import { skillToMarkdown } from '@/services/extensions/zipExportService';
import type { AgentHost } from '@/types/agent';
import type { OfficeHostApp } from '@/services/office/host';
import { generateId } from '@/utils/id';

const MODEL_FETCH_TIMEOUT_MS = 10_000;

/** Race a promise against a timeout. */
function withTimeout<T>(promise: Promise<T>, ms: number, label: string): Promise<T> {
  return new Promise<T>((resolve, reject) => {
    const timer = setTimeout(() => reject(new Error(`${label} timed out after ${ms}ms`)), ms);
    promise.then(
      v => {
        clearTimeout(timer);
        resolve(v);
      },
      e => {
        clearTimeout(timer);
        reject(e instanceof Error ? e : new Error(String(e)));
      }
    );
  });
}

/** Fetch available models from the Copilot SDK and update the store. */
async function loadAvailableModels(client: WebSocketCopilotClient): Promise<void> {
  try {
    const modelInfos = await withTimeout(client.listModels(), MODEL_FETCH_TIMEOUT_MS, 'listModels');
    const models = modelInfos.map(m => ({
      id: m.id,
      name: m.name,
      provider: inferProvider(m.id),
    }));
    useSettingsStore.getState().setAvailableModels(models);

    // Auto-correct activeModel if it's not in the fetched list
    const { activeModel } = useSettingsStore.getState();
    if (models.length > 0 && !models.some(m => m.id === activeModel)) {
      console.warn(
        `[useOfficeChat] activeModel '${activeModel}' not in available models, switching to '${models[0].id}'`
      );
      useSettingsStore.getState().setActiveModel(models[0].id);
    }
  } catch (err) {
    console.warn('[useOfficeChat] Failed to load available models:', err);
  }
}

function getWsUrl(): string {
  if (typeof window === 'undefined') return 'wss://localhost:3000/api/copilot';
  const { hostname, protocol, host } = window.location;
  // When served from GitHub Pages (staging) or any non-localhost origin,
  // the WebSocket proxy is always on localhost:3000.
  if (hostname !== 'localhost' && hostname !== '127.0.0.1') {
    return 'wss://localhost:3000/api/copilot';
  }
  const proto = protocol === 'https:' ? 'wss:' : 'ws:';
  return `${proto}//${host}/api/copilot`;
}

export function useOfficeChat(host: OfficeHostApp) {
  const activeModel = useSettingsStore(s => s.activeModel);
  const activeSkillNames = useSettingsStore(s => s.activeSkillNames);
  const activeAgentId = useSettingsStore(s => s.activeAgentId);
  const importedMcpServers = useSettingsStore(s => s.importedMcpServers);
  const activeMcpServerNames = useSettingsStore(s => s.activeMcpServerNames);

  const clientRef = useRef<WebSocketCopilotClient | null>(null);
  const sessionRef = useRef<BrowserCopilotSession | null>(null);
  const cancelRef = useRef(false);
  // Guard against concurrent/stale initSession calls (React StrictMode double-mount)
  const initCounterRef = useRef(0);

  const [messages, setMessages] = useState<ThreadMessageLike[]>([]);
  const [isRunning, setIsRunning] = useState(false);
  const [sessionError, setSessionError] = useState<Error | null>(null);
  const [isConnecting, setIsConnecting] = useState(true);
  const [thinkingText, setThinkingText] = useState<string | null>(null);

  const initSession = useCallback(async () => {
    // Increment counter — any in-flight init with a stale counter will be discarded
    const thisInit = ++initCounterRef.current;

    if (clientRef.current) {
      try {
        await clientRef.current.stop();
      } catch {
        /* ignore */
      }
      clientRef.current = null;
      sessionRef.current = null;
    }

    const wsUrl = getWsUrl();
    console.log('[chat] initSession: connecting to', wsUrl);
    setIsConnecting(true);
    setSessionError(null);

    try {
      const client = await withTimeout(createWebSocketClient(wsUrl), 15_000, 'WebSocket connect');

      // If a newer initSession started while we were connecting, discard this one
      if (initCounterRef.current !== thisInit) {
        void client.stop().catch(() => {
          /* discard */
        });
        return;
      }

      clientRef.current = client;
      console.log('[chat] WebSocket connected');

      const resolvedAgent = resolveActiveAgent(activeAgentId, host);

      // System prompt: only base + app prompt (no agent/skill concatenation)
      const systemContent = buildSystemPrompt(host);

      // Build imported skill payloads for the proxy to write to disk
      const importedHostSkills = getImportedSkills().filter(
        s => s.metadata.hosts.length === 0 || s.metadata.hosts.includes(host as AgentHost)
      );
      const skills = importedHostSkills.map(s => ({
        name: s.metadata.name,
        content: skillToMarkdown(s),
      }));

      // Compute disabled skill names from activeSkillNames
      const allHostSkillNames = getSkills()
        .filter(s => s.metadata.hosts.length === 0 || s.metadata.hosts.includes(host as AgentHost))
        .map(s => s.metadata.name);
      const disabledSkills =
        activeSkillNames !== null
          ? allHostSkillNames.filter(name => !activeSkillNames.includes(name))
          : [];

      // Build custom agent config for the SDK
      const customAgents = resolvedAgent
        ? [
            {
              name: resolvedAgent.metadata.name,
              description: resolvedAgent.metadata.description,
              prompt: resolvedAgent.instructions,
            },
          ]
        : undefined;

      // Resolve active MCP servers, intersect with agent allowlist if specified
      let activeServers = resolveActiveMcpServers(importedMcpServers, activeMcpServerNames);
      if (resolvedAgent?.metadata.mcpServers !== undefined) {
        const agentMcpAllowlist = new Set(resolvedAgent.metadata.mcpServers);
        activeServers = activeServers.filter(s => agentMcpAllowlist.has(s.name));
      }
      const mcpServers = activeServers.length > 0 ? toSdkMcpServers(activeServers) : undefined;

      // Per-agent tool restriction (omit = all tools available)
      const availableTools = resolvedAgent?.metadata.tools;

      const session = await withTimeout(
        client.createSession({
          model: activeModel,
          systemMessage: { mode: 'replace', content: systemContent },
          tools: getToolsForHost(host),
          mcpServers,
          availableTools,
          skills,
          disabledSkills,
          customAgents,
        }),
        60_000,
        'session.create'
      );

      // If a newer initSession started while we were creating the session, discard
      if (initCounterRef.current !== thisInit) {
        void client.stop().catch(() => {
          /* discard */
        });
        return;
      }

      sessionRef.current = session;
      setSessionError(null);
      console.log('[chat] Session created:', session.sessionId);

      // Fetch available models (non-blocking, with timeout)
      void loadAvailableModels(client);
    } catch (err) {
      // If superseded by a newer init, silently bail
      if (initCounterRef.current !== thisInit) return;
      console.error('[chat] initSession failed:', err);
      setSessionError(err instanceof Error ? err : new Error(String(err)));
    } finally {
      if (initCounterRef.current === thisInit) {
        setIsConnecting(false);
      }
    }
  }, [
    activeModel,
    host,
    activeSkillNames,
    activeAgentId,
    importedMcpServers,
    activeMcpServerNames,
  ]);

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
      .filter(
        (c): c is { type: string; text: string } => c.type === 'text' && typeof c.text === 'string'
      )
      .map(c => c.text)
      .join('\n');

    if (!userText.trim()) return;

    if (!sessionRef.current) {
      const errorMsg: ThreadMessageLike = {
        id: generateId(),
        role: 'assistant',
        content: [
          {
            type: 'text',
            text: 'Not connected to Copilot. Check that the server is running and try clicking **Retry** above, or start a new conversation.',
          },
        ],
        status: { type: 'incomplete', reason: 'error' },
        createdAt: new Date(),
      };
      setMessages(prev => [
        ...prev,
        {
          id: generateId(),
          role: 'user',
          content: [{ type: 'text', text: userText }],
          createdAt: new Date(),
        },
        errorMsg,
      ]);
      return;
    }

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
    // Set explicit default text so the standalone ThinkingIndicator renders
    // immediately via React context — no dependency on the runtime's deferred
    // useEffect adapter sync.
    setThinkingText('Thinking…');

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

      // Stale-response watchdog: if no event arrives within 30s, warn the user.
      // Reset on every event; cleared when the stream ends.
      const STALE_TIMEOUT = 30_000;
      let staleTimer: ReturnType<typeof setTimeout> | null = null;
      const resetStaleTimer = () => {
        if (staleTimer) clearTimeout(staleTimer);
        staleTimer = setTimeout(() => {
          setThinkingText('Still waiting for a response…');
        }, STALE_TIMEOUT);
      };
      resetStaleTimer();

      for await (const event of session.query({ prompt: userText })) {
        resetStaleTimer();
        if (cancelRef.current) break;

        if (event.type === 'assistant.message_delta') {
          streamText += event.data.deltaContent;
          updateAssistant();
        } else if (event.type === 'tool.execution_start') {
          const { toolCallId, toolName, arguments: args } = event.data;
          // report_intent is an internal SDK tool — surface intent as thinking text
          if (toolName === 'report_intent') {
            const intent = (args as Record<string, unknown> | undefined)?.intent;
            if (typeof intent === 'string' && intent) {
              // flushSync forces React to commit this state update to the DOM
              // immediately, before the for-await loop processes the next
              // buffered event.  Without it, React 18 automatic batching can
              // merge this update with a later setThinkingText(null), so the
              // intermediate text never appears on screen.
              flushSync(() => setThinkingText(intent));
            }
            continue;
          }
          flushSync(() => setThinkingText(`${humanizeToolName(toolName)}…`));
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
          setThinkingText(null);
          updateAssistant({ status: { type: 'complete', reason: 'stop' } });
        } else if (event.type === 'session.idle') {
          // Stream ended — finalize message if it wasn't already completed by
          // an assistant.message event (e.g. streaming-only responses).
          setThinkingText(null);
          updateAssistant({ status: { type: 'complete', reason: 'stop' } });
        } else if (event.type === 'session.error') {
          setThinkingText(null);
          updateAssistant({
            status: { type: 'incomplete', reason: 'error', error: event.data.message },
          });
          break;
        }
      }
      if (staleTimer) clearTimeout(staleTimer);
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
      setThinkingText(null);
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

  return { runtime, sessionError, isConnecting, clearMessages, thinkingText };
}
