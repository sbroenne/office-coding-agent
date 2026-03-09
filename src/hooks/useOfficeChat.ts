import { useState, useRef, useCallback, useEffect } from 'react';
import { flushSync } from 'react-dom';
import type { WebSocketCopilotClient, BrowserCopilotSession } from '@/lib/websocket-client';
import type { PermissionRequestPayload } from '@/lib/websocket-client';
import { createWebSocketClient } from '@/lib/websocket-client';
import { getToolsForHost } from '@/tools';
import { resolveActiveAgent, getAgents, SUPPORTED_AGENT_HOSTS } from '@/services/agents';
import { toSdkMcpServers } from '@/services/mcp';
import { useSettingsStore } from '@/stores';
import { useSessionHistoryStore } from '@/stores';
import { usePermissionStore } from '@/stores';
import { useMcpStatusStore } from '@/stores';
import { buildSystemPrompt } from '@/services/ai/systemPrompt';
import { inferProvider, BUNDLED_MCP_SERVERS } from '@/types';
import type { ChatMessage, ToolCallPart } from '@/types';
import type { AgentHost } from '@/types/agent';
import type { AgentSkill } from '@/types/skill';
import type { OfficeHostApp } from '@/services/office/host';
import { generateId } from '@/utils/id';

const MODEL_FETCH_TIMEOUT_MS = 10_000;
const DEFAULT_THINKING_TEXT = 'Thinking…';

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
  const activeAgentId = useSettingsStore(s => s.activeAgentId);
  const disabledSkillNames = useSettingsStore(s => s.disabledSkillNames);
  const disabledMcpServerNames = useSettingsStore(s => s.disabledMcpServerNames);
  const sessions = useSessionHistoryStore(s => s.sessions);
  const activeSessionId = useSessionHistoryStore(s => s.activeSessionId);
  const createSession = useSessionHistoryStore(s => s.createSession);
  const setActiveSession = useSessionHistoryStore(s => s.setActiveSession);
  const upsertActiveSession = useSessionHistoryStore(s => s.upsertActiveSession);
  const deleteSessionHistoryItem = useSessionHistoryStore(s => s.deleteSession);
  const evaluatePermission = usePermissionStore(s => s.evaluate);
  const addPermissionRule = usePermissionStore(s => s.addRule);
  const allowAllPermissions = usePermissionStore(s => s.allowAll);

  const clientRef = useRef<WebSocketCopilotClient | null>(null);
  const sessionRef = useRef<BrowserCopilotSession | null>(null);
  const cancelRef = useRef(false);
  // Guard against concurrent/stale initSession calls (React StrictMode double-mount)
  const initCounterRef = useRef(0);
  const restoredInitialSessionRef = useRef(false);

  // Stable refs for settings that should NOT trigger session re-init when they change.
  // initSession reads from these refs so the useCallback only re-creates when host
  // changes — not on every skill/agent/MCP/model toggle mid-conversation.
  // Without this, any store update (e.g. WorkIQ connecting, switching model) would
  // tear down and restart the Copilot session, losing all conversation context.
  // Model changes take effect on the next new conversation (same as VS Code Copilot).
  const activeModelRef = useRef(activeModel);
  const activeAgentIdRef = useRef(activeAgentId);
  const disabledSkillNamesRef = useRef(disabledSkillNames);
  const disabledMcpServerNamesRef = useRef(disabledMcpServerNames);
  const evaluatePermissionRef = useRef(evaluatePermission);
  // Keep refs in sync on every render (runs synchronously, before any effects)
  activeModelRef.current = activeModel;
  activeAgentIdRef.current = activeAgentId;
  disabledSkillNamesRef.current = disabledSkillNames;
  disabledMcpServerNamesRef.current = disabledMcpServerNames;
  evaluatePermissionRef.current = evaluatePermission;

  // Switch model mid-session when the user picks a different model
  useEffect(() => {
    const session = sessionRef.current;
    if (session && activeModel) {
      session.setModel(activeModel).catch(err => {
        console.warn('[chat] model.switch failed:', err);
      });
    }
  }, [activeModel]);

  const [messages, setMessages] = useState<ChatMessage[]>([]);
  const [isRunning, setIsRunning] = useState(false);
  const [sessionError, setSessionError] = useState<Error | null>(null);
  const [isConnecting, setIsConnecting] = useState(true);
  const [pendingPermission, setPendingPermission] = useState<PermissionRequestPayload | null>(null);
  const activePermissionRequestRef = useRef<string | null>(null);

  // Prompt queue: prompts enqueued via Ctrl+Q while a response is in flight.
  // Ref is the source of truth (avoids stale closures in send); state drives UI.
  const queueRef = useRef<string[]>([]);
  const [queuedPrompts, setQueuedPrompts] = useState<string[]>([]);
  // Ref to call send() from inside its own finally block (auto-dequeue).
  // eslint-disable-next-line @typescript-eslint/no-empty-function
  const sendRef = useRef<(text: string) => Promise<void>>(async () => {});

  const deserializeMessages = useCallback((rawMessages: unknown[]): ChatMessage[] => {
    return rawMessages
      .filter((msg): msg is Record<string, unknown> => typeof msg === 'object' && msg !== null)
      .map(msg => {
        const createdAtRaw = msg.createdAt;
        const createdAt =
          typeof createdAtRaw === 'string' || typeof createdAtRaw === 'number'
            ? new Date(createdAtRaw)
            : new Date();
        return {
          ...(msg as unknown as ChatMessage),
          createdAt,
          thinkingText: null,
        };
      });
  }, []);

  const deriveSessionTitle = useCallback((nextMessages: ChatMessage[]): string => {
    const firstUser = nextMessages.find(m => m.role === 'user');
    const contentParts: unknown[] = Array.isArray(firstUser?.content) ? firstUser.content : [];
    const textPart = contentParts.find(
      (part): part is { type: 'text'; text?: string } =>
        typeof part === 'object' &&
        part !== null &&
        'type' in part &&
        (part as { type?: unknown }).type === 'text'
    );
    const text = textPart ? String(textPart.text ?? '') : '';
    const trimmed = text.trim();
    if (!trimmed) return 'New conversation';
    return trimmed.length > 60 ? `${trimmed.slice(0, 60)}…` : trimmed;
  }, []);

  // Stable ref so onNew can call the latest initSession without adding it to deps
  const initSessionRef = useRef<() => Promise<void>>(() => Promise.resolve());

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

      // Register MCP event handlers to forward to the status store
      const mcpStore = useMcpStatusStore.getState();
      mcpStore.clearAll();
      client.onMcpStatus(payload => {
        useMcpStatusStore.getState().setStatus(payload.server, payload.status, payload.error);
      });
      client.onMcpLog(payload => {
        useMcpStatusStore.getState().addLog(payload.server, {
          timestamp: payload.timestamp,
          level: payload.level,
          message: payload.message,
        });
      });
      client.onMcpTools(payload => {
        useMcpStatusStore.getState().setTools(payload.server, payload.tools);
      });

      // Register plugin.agents notification handler so discovered plugin agents are
      // reflected in the browser-side agentService (AgentPicker). Agents without hosts
      // keep hosts:[] which getAgents() treats as "all hosts" (universal agent).
      client.onPluginAgents(({ agents }) => {
        const agentConfigs = agents.map(agent => ({
          metadata: {
            name: agent.name,
            description: agent.description,
            version: '0.0.0',
            hosts: agent.hosts.filter((h): h is AgentHost =>
              SUPPORTED_AGENT_HOSTS.includes(h as AgentHost)
            ),
            defaultForHosts: [] as AgentHost[],
          },
          instructions: agent.prompt,
        }));
        // Update the store — this syncs agentService AND triggers AgentPicker re-render
        useSettingsStore.getState().setPluginAgents(agentConfigs);
      });

      // Register plugin.skills notification — adds plugin skills to SkillPicker and
      // the skill context injected into the system prompt.
      client.onPluginSkills(({ skills }) => {
        const skillObjects: AgentSkill[] = skills.map(s => ({
          metadata: {
            name: s.name,
            description: s.description,
            version: s.version,
            tags: [],
            hosts: s.hosts.filter((h): h is AgentHost =>
              SUPPORTED_AGENT_HOSTS.includes(h as AgentHost)
            ),
          },
          content: s.content,
        }));
        useSettingsStore.getState().setPluginSkills(skillObjects);
      });

      // Register plugin.prompts notification — populates the slash command menu
      // in ChatComposer.
      client.onPluginPrompts(({ prompts }) => {
        useSettingsStore.getState().setPluginPrompts(prompts);
      });

      const resolvedAgent = resolveActiveAgent(activeAgentIdRef.current, host);

      // System prompt: base + app prompt + user memories
      let systemContent = buildSystemPrompt(host);

      // Inject persistent user memories if any exist
      try {
        const { useMemoryStore } = await import('@/stores/memoryStore');
        const memoryContext = useMemoryStore.getState().buildMemoryContext();
        if (memoryContext) {
          systemContent += `\n\n${memoryContext}`;
        }
      } catch {
        // Memory store not available — continue without memories
      }

      // Build custom agent configs for ALL agents in this host — this enables sub-agent
      // delegation where the active agent can invoke other agents as sub-agents.
      // Each agent carries its own tool allowlist so per-agent restrictions are enforced
      // by the SDK rather than at the session level.
      const allHostAgents = getAgents(host);
      const customAgents =
        allHostAgents.length > 0
          ? allHostAgents.map(agent => ({
              name: agent.metadata.name,
              description: agent.metadata.description,
              prompt: agent.instructions,
              // null = all tools; undefined means the same but we use null for explicitness
              tools: agent.metadata.tools ?? null,
            }))
          : undefined;

      // Resolve active MCP servers: bundled list → agent allowlist filter → user disable filter.
      let activeServers = [...BUNDLED_MCP_SERVERS];
      if (resolvedAgent?.metadata.mcpServers !== undefined) {
        const agentMcpAllowlist = new Set(resolvedAgent.metadata.mcpServers);
        activeServers = activeServers.filter(s => agentMcpAllowlist.has(s.name));
      }
      activeServers = activeServers.filter(
        s => !disabledMcpServerNamesRef.current.includes(s.name)
      );
      const mcpServers = activeServers.length > 0 ? toSdkMcpServers(activeServers) : undefined;

      const disabledSkills =
        disabledSkillNamesRef.current.length > 0 ? disabledSkillNamesRef.current : undefined;

      const session = await withTimeout(
        client.createSession({
          model: activeModelRef.current,
          systemMessage: { mode: 'replace', content: systemContent },
          tools: getToolsForHost(host),
          mcpServers,
          host,
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
      setPendingPermission(null);
      activePermissionRequestRef.current = null;
      console.log('[chat] Session created:', session.sessionId);

      session.onPermissionRequest(payload => {
        const autoDecision = evaluatePermissionRef.current(payload.request);
        if (autoDecision === 'approved') {
          void session.respondPermission(payload.requestId, 'approved');
          return;
        }
        activePermissionRequestRef.current = payload.requestId;
        setPendingPermission(payload);
      });

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
    // Only re-init the session when host changes — that requires a genuinely new
    // connection. All other settings (model, skills, agents, MCP servers) are read
    // via refs above so mid-conversation store updates never tear down an active
    // session. The new values take effect on the next fresh conversation.
    host,
  ]);

  useEffect(() => {
    initSessionRef.current = initSession;
  }, [initSession]);

  useEffect(() => {
    if (restoredInitialSessionRef.current) return;

    if (!activeSessionId) {
      createSession(host);
      restoredInitialSessionRef.current = true;
      return;
    }

    const active = sessions.find(s => s.id === activeSessionId);
    if (active && active.messages.length > 0) {
      setMessages(deserializeMessages(active.messages));
    }

    restoredInitialSessionRef.current = true;
  }, [activeSessionId, createSession, deserializeMessages, host, sessions]);

  useEffect(() => {
    if (!restoredInitialSessionRef.current) return;
    const title = deriveSessionTitle(messages);
    upsertActiveSession({
      host,
      title,
      messages,
    });
  }, [deriveSessionTitle, host, messages, upsertActiveSession]);

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

  const send = useCallback(async (userText: string) => {
    if (!userText.trim()) return;

    const client = clientRef.current;
    if (!sessionRef.current || !client) {
      const errorMsg: ChatMessage = {
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

    // Detect multi-slide PowerPoint requests → use orchestrator
    const isMultiSlideRequest =
      host === 'powerpoint' &&
      /\b(\d+)\s*(slides?|folien?|seiten?)\b/i.test(userText) &&
      !userText.toLowerCase().includes('this slide');

    // Detect deep-mode Word document requests → use document orchestrator
    // Triggers on: deep keywords OR multi-section requests (like "write a report with 5 sections")
    const isDeepWordRequest =
      host === 'word' &&
      (/\b(deep|gründlich|ausführlich|thoroughly|think|go\s*deep|detail(liert)?|qualit)/i.test(
        userText
      ) ||
        /\b(\d+)\s*(sections?|abschnitt(e|en)?|kapitel|teil(e|en)?|chapters?)\b/i.test(userText) ||
        /\b(erstell|schreib|create|write|build|generate|verfass)\w*\b.{0,30}\b(report|bericht|dokument|document|paper|aufsatz|memo|proposal|angebot|zusammenfassung)\b/i.test(
          userText
        ));

    if (isDeepWordRequest) {
      const assistantId = generateId();
      cancelRef.current = false;

      setMessages(prev => [
        ...prev,
        {
          id: generateId(),
          role: 'user',
          content: [{ type: 'text', text: userText }],
          createdAt: new Date(),
        },
        {
          id: assistantId,
          role: 'assistant',
          content: [{ type: 'text', text: '' }],
          status: { type: 'running' },
          createdAt: new Date(),
        },
      ]);
      setIsRunning(true);

      let streamText = '';
      const updateText = (extra?: Partial<Pick<ChatMessage, 'status'>>) => {
        setMessages(prev =>
          prev.map(m =>
            m.id === assistantId
              ? { ...m, content: [{ type: 'text', text: streamText }], ...extra }
              : m
          )
        );
      };

      const abortController = new AbortController();
      const origCancel = cancelRef.current;
      const cancelCheck = setInterval(() => {
        if (cancelRef.current && !origCancel) abortController.abort();
      }, 500);

      try {
        const { orchestrateDocument } = await import('@/hooks/useDocumentOrchestrator');
        const docMode =
          /\b(deep|gründlich|ausführlich|thoroughly|think|go\s*deep|detail(liert)?|qualit)/i.test(
            userText
          )
            ? ('deep' as const)
            : ('fast' as const);
        const currentModelForDoc = useSettingsStore.getState().activeModel;
        await orchestrateDocument(
          client,
          currentModelForDoc,
          userText,
          {
            onPlan: () => {
              /* plan received */
            },
            onSectionProgress: () => {
              /* section status changed */
            },
            onText: (text: string) => {
              streamText += text;
              updateText();
            },
            onWorkerEvent: () => {
              /* worker tool events */
            },
            onComplete: () => {
              updateText({ status: { type: 'complete', reason: 'stop' } });
            },
            onError: (error: string) => {
              streamText += `\n\n❌ Error: ${error}`;
              updateText({ status: { type: 'incomplete', reason: 'error', error } });
            },
          },
          abortController.signal,
          docMode
        );
      } catch (err) {
        const errMsg = err instanceof Error ? err.message : String(err);
        streamText += `\n\n❌ ${errMsg}`;
        updateText({ status: { type: 'incomplete', reason: 'error', error: errMsg } });
      } finally {
        clearInterval(cancelCheck);
        setIsRunning(false);
      }
      return;
    }

    if (isMultiSlideRequest) {
      const assistantId = generateId();
      cancelRef.current = false;

      setMessages(prev => [
        ...prev,
        {
          id: generateId(),
          role: 'user',
          content: [{ type: 'text', text: userText }],
          createdAt: new Date(),
        },
        {
          id: assistantId,
          role: 'assistant',
          content: [{ type: 'text', text: '' }],
          status: { type: 'running' },
          createdAt: new Date(),
        },
      ]);
      setIsRunning(true);

      let streamText = '';
      const updateText = (extra?: Partial<Pick<ChatMessage, 'status'>>) => {
        setMessages(prev =>
          prev.map(m =>
            m.id === assistantId
              ? { ...m, content: [{ type: 'text', text: streamText }], ...extra }
              : m
          )
        );
      };

      const abortController = new AbortController();
      const origCancel = cancelRef.current;
      // Check cancel periodically
      const cancelCheck = setInterval(() => {
        if (cancelRef.current && !origCancel) abortController.abort();
      }, 500);

      try {
        const { orchestrateDeck } = await import('@/hooks/useDeckOrchestrator');
        const deckMode = /\b(deep|detail|qualit)/i.test(userText)
          ? ('deep' as const)
          : ('fast' as const);
        const currentModelForDeck = useSettingsStore.getState().activeModel;
        await orchestrateDeck(
          client,
          currentModelForDeck,
          userText,
          {
            onPlan: () => {
              /* plan received */
            },
            onSlideProgress: () => {
              /* slide status changed */
            },
            onText: (text: string) => {
              streamText += text;
              updateText();
            },
            onWorkerEvent: () => {
              /* worker tool events */
            },
            onComplete: () => {
              updateText({ status: { type: 'complete', reason: 'stop' } });
            },
            onError: (error: string) => {
              streamText += `\n\n❌ Error: ${error}`;
              updateText({ status: { type: 'incomplete', reason: 'error', error } });
            },
          },
          abortController.signal,
          deckMode
        );
      } catch (err) {
        const errMsg = err instanceof Error ? err.message : String(err);
        streamText += `\n\n❌ ${errMsg}`;
        updateText({ status: { type: 'incomplete', reason: 'error', error: errMsg } });
      } finally {
        clearInterval(cancelCheck);
        setIsRunning(false);
      }
      return;
    }

    const assistantId = generateId();
    cancelRef.current = false;

    const userMsg: ChatMessage = {
      id: generateId(),
      role: 'user',
      content: [{ type: 'text', text: userText }],
      createdAt: new Date(),
    };

    const assistantMsg: ChatMessage = {
      id: assistantId,
      role: 'assistant',
      content: [],
      status: { type: 'running' },
      thinkingText: DEFAULT_THINKING_TEXT,
      createdAt: new Date(),
    };

    setMessages(prev => [...prev, userMsg, assistantMsg]);
    setIsRunning(true);

    /** Update thinkingText on the specific assistant message */
    const setThinkingForAssistant = (text: string | null) => {
      setMessages(prev => prev.map(m => (m.id === assistantId ? { ...m, thinkingText: text } : m)));
    };

    const toolParts = new Map<string, ToolCallPart>();
    let streamText = '';
    // Tracks the current phase index. Increments each time report_intent fires
    // AFTER at least one tool has been added, creating a new Working box segment.
    let currentPhase = 0;
    // Tracks the current phase label (from report_intent). This becomes the
    // Working box header text — matching VS Code's IChatTask.content behavior.
    let currentPhaseLabel: string | undefined = undefined;

    const updateAssistant = (extra?: Partial<Pick<ChatMessage, 'status' | 'thinkingText'>>) => {
      // Text part is ALWAYS at index 0 — even when empty — to prevent tearing.
      // Tool cards appear visually above text via CSS order (ToolGroup has order: -1).
      const content: ChatMessage['content'] = [
        { type: 'text' as const, text: streamText },
        ...Array.from(toolParts.values()),
      ];
      setMessages(prev => prev.map(m => (m.id === assistantId ? { ...m, content, ...extra } : m)));
    };

    // Stale-response watchdog: if no event arrives within 30s, warn the user.
    // Reset on every event; cleared when the stream ends.
    // Declared outside `try` so the `finally` block can clear it.
    const STALE_TIMEOUT = 30_000;
    let staleTimer: ReturnType<typeof setTimeout> | null = null;

    try {
      const session = sessionRef.current;
      const resetStaleTimer = () => {
        if (staleTimer) clearTimeout(staleTimer);
        staleTimer = setTimeout(() => {
          setThinkingForAssistant('Still waiting for a response…');
        }, STALE_TIMEOUT);
      };
      resetStaleTimer();

      for await (const event of session.query({ prompt: userText })) {
        resetStaleTimer();
        if (cancelRef.current) break;

        if (event.type === 'assistant.message_delta') {
          // First streaming delta clears the thinking indicator
          streamText += event.data.deltaContent;
          updateAssistant({ thinkingText: null });
        } else if (event.type === 'tool.execution_start') {
          const { toolCallId, toolName, arguments: args } = event.data;
          // report_intent is an internal SDK tool — surface intent as thinking text
          if (toolName === 'report_intent') {
            const intent = (args as Record<string, unknown> | undefined)?.intent;
            if (typeof intent === 'string' && intent) {
              // If tools have already been added, this intent starts a NEW phase
              if (toolParts.size > 0) {
                currentPhase++;
              }
              // The intent text labels the Working box (VS Code: IChatTask.content)
              currentPhaseLabel = intent;
              flushSync(() => setThinkingForAssistant(intent));
            }
            continue;
          }
          toolParts.set(toolCallId, {
            type: 'tool-call',
            toolCallId,
            toolName,
            argsText: JSON.stringify(args ?? {}),
            status: { type: 'running' },
            phaseIndex: currentPhase,
            phaseLabel: currentPhaseLabel,
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
            toolParts.set(toolCallId, {
              ...existing,
              result: resultText,
              status: { type: 'complete' },
            });
            updateAssistant();
          }
          // Reset thinking text to "Thinking…" so the shimmer reappears
          // in the gap between this tool completing and the next action.
          setThinkingForAssistant(DEFAULT_THINKING_TEXT);
        } else if (event.type === 'assistant.message') {
          // Update text content but DON'T clear thinking or mark complete here.
          // The SDK sends assistant.message BEFORE tool calls (model's initial
          // response) AND after (final response). Only session.idle reliably
          // indicates the response is truly finished.
          streamText = event.data.content;
          updateAssistant();
        } else if (event.type === 'session.idle') {
          // Stream truly ended — clear thinking and finalize
          updateAssistant({ status: { type: 'complete', reason: 'stop' }, thinkingText: null });
        } else if (event.type === 'session.error') {
          updateAssistant({
            status: { type: 'incomplete', reason: 'error', error: event.data.message },
            thinkingText: null,
          });
          break;
        } else if (event.type === 'subagent.started') {
          // A sub-agent has been invoked — show which agent is being asked
          flushSync(() =>
            setThinkingForAssistant(
              `Asking ${event.data.agentDisplayName || event.data.agentName}…`
            )
          );
        } else if (event.type === 'subagent.completed' || event.type === 'subagent.failed') {
          // Sub-agent finished — return to generic thinking indicator
          setThinkingForAssistant(DEFAULT_THINKING_TEXT);
        }
      }
    } catch (err) {
      const errMsg = err instanceof Error ? err.message : String(err);
      // Auto-reconnect if the Copilot session was lost (e.g. proxy restart, laptop sleep)
      const isSessionLost =
        errMsg.includes('Session') ||
        errMsg.includes('not connected') ||
        errMsg.includes('disconnected') ||
        errMsg.includes('WebSocket closed');
      if (isSessionLost) {
        setMessages(prev =>
          prev.map(m =>
            m.id === assistantId
              ? {
                  ...m,
                  content: [{ type: 'text', text: '🔄 Session lost — reconnecting…' }],
                  status: { type: 'complete', reason: 'stop' },
                  thinkingText: null,
                }
              : m
          )
        );
        void initSessionRef.current();
      } else {
        setMessages(prev =>
          prev.map(m =>
            m.id === assistantId
              ? {
                  ...m,
                  status: { type: 'incomplete', reason: 'error', error: errMsg },
                  thinkingText: null,
                }
              : m
          )
        );
      }
    } finally {
      if (staleTimer) clearTimeout(staleTimer);
      // Ensure thinkingText is cleared and isRunning is reset
      setMessages(prev => prev.map(m => (m.id === assistantId ? { ...m, thinkingText: null } : m)));
      setIsRunning(false);

      // Auto-dequeue: if prompts were enqueued via Ctrl+Q, send the next one.
      const next = queueRef.current.shift();
      if (next !== undefined) {
        setQueuedPrompts([...queueRef.current]);
        // Defer to avoid re-entering send() synchronously from its own finally
        setTimeout(() => void sendRef.current(next), 0);
      }
    }
  }, []);

  // Keep sendRef in sync so auto-dequeue can call the latest send()
  sendRef.current = send;

  /** Enqueue a prompt to run after the current response finishes. */
  const enqueue = useCallback((text: string) => {
    const trimmed = text.trim();
    if (!trimmed) return;
    queueRef.current = [...queueRef.current, trimmed];
    setQueuedPrompts([...queueRef.current]);
  }, []);

  /** Remove a single queued prompt by index. */
  const dequeue = useCallback((index: number) => {
    queueRef.current = queueRef.current.filter((_, i) => i !== index);
    setQueuedPrompts([...queueRef.current]);
  }, []);

  /** Clear all queued prompts. */
  const clearQueue = useCallback(() => {
    queueRef.current = [];
    setQueuedPrompts([]);
  }, []);

  const cancel = useCallback(() => {
    cancelRef.current = true;

    // Clear any queued prompts — user cancelled, don't auto-send more
    queueRef.current = [];
    setQueuedPrompts([]);

    // Immediately update UI: mark the running message as incomplete/cancelled
    // and clear thinking text so the user gets instant feedback.
    flushSync(() => {
      setIsRunning(false);
      setMessages(prev => {
        const last = prev[prev.length - 1];
        if (last?.role === 'assistant' && last.status?.type !== 'complete') {
          return [
            ...prev.slice(0, -1),
            { ...last, status: { type: 'incomplete', reason: 'cancelled' }, thinkingText: null },
          ];
        }
        return prev;
      });
    });

    // Also destroy and recreate the session to actually stop the SDK query.
    // Without this, the proxy/SDK keeps processing events until the response
    // finishes naturally.
    const session = sessionRef.current;
    if (session) {
      void session.destroy().catch(() => {
        /* ignore */
      });
      sessionRef.current = null;
      void initSessionRef.current();
    }
  }, []);

  const clearMessages = useCallback(() => {
    setMessages([]);
    setPendingPermission(null);
    activePermissionRequestRef.current = null;
    // Clear queued prompts on new conversation
    queueRef.current = [];
    setQueuedPrompts([]);
    createSession(host);
    void initSession();
  }, [createSession, host, initSession]);

  const restoreSession = useCallback(
    (sessionId: string) => {
      const session = sessions.find(s => s.id === sessionId);
      if (!session) return;
      setActiveSession(sessionId);
      setMessages(deserializeMessages(session.messages));
      setPendingPermission(null);
      activePermissionRequestRef.current = null;
      void initSession();
    },
    [deserializeMessages, initSession, sessions, setActiveSession]
  );

  const deleteSession = useCallback(
    (sessionId: string) => {
      deleteSessionHistoryItem(sessionId);
      if (activeSessionId === sessionId) {
        setMessages([]);
        createSession(host);
        void initSession();
      }
    },
    [activeSessionId, createSession, deleteSessionHistoryItem, host, initSession]
  );

  const respondPermission = useCallback(async (decision: 'approved' | 'denied') => {
    const session = sessionRef.current;
    const requestId = activePermissionRequestRef.current;
    if (!session || !requestId) return;
    try {
      await session.respondPermission(requestId, decision);
    } finally {
      if (activePermissionRequestRef.current === requestId) {
        activePermissionRequestRef.current = null;
        setPendingPermission(null);
      }
    }
  }, []);

  const approvePermission = useCallback(() => {
    void respondPermission('approved');
  }, [respondPermission]);

  const denyPermission = useCallback(() => {
    void respondPermission('denied');
  }, [respondPermission]);

  const allowPermissionAlways = useCallback(() => {
    const request = pendingPermission?.request;
    if (!request) return;
    const pathPrefix =
      (typeof request.path === 'string' && request.path) ||
      (typeof request.fileName === 'string' && request.fileName) ||
      (typeof request.fullCommandText === 'string' && request.fullCommandText) ||
      null;

    if (pathPrefix) {
      addPermissionRule({
        kind: request.kind,
        pathPrefix,
      });
    }
    void respondPermission('approved');
  }, [addPermissionRule, pendingPermission, respondPermission]);

  const compactSession = useCallback(async () => {
    const session = sessionRef.current;
    if (session) {
      try {
        await session.compact();
        console.log('[chat] session compacted');
      } catch (err) {
        console.warn('[chat] session.compact failed:', err);
      }
    }
  }, []);

  /**
   * Switch to a different model mid-session without starting a new conversation.
   * Requires an active session with the Copilot SDK v0.1.30+.
   *
   * @param modelId - The model ID to switch to
   * @returns Promise that resolves when the model has been switched
   * @throws Error if no active session exists or the switch fails
   */
  const switchModel = useCallback(
    async (modelId: string) => {
      const session = sessionRef.current;
      if (!session) {
        throw new Error('Cannot switch model: no active session');
      }

      try {
        console.log(`[chat] Switching model from ${activeModel} to ${modelId}`);
        await session.setModel(modelId);
        // Update the store so the UI reflects the new model
        useSettingsStore.getState().setActiveModel(modelId);
        console.log(`[chat] Model switched successfully to ${modelId}`);
      } catch (err) {
        console.error('[chat] Failed to switch model:', err);
        throw err instanceof Error ? err : new Error(String(err));
      }
    },
    [activeModel]
  );

  return {
    messages,
    isRunning,
    send,
    cancel,
    sessionError,
    isConnecting,
    clearMessages,
    restoreSession,
    deleteSession,
    sessions,
    activeSessionId,
    pendingPermission,
    allowAllPermissions,
    approvePermission,
    denyPermission,
    allowPermissionAlways,
    compactSession,
    switchModel,
    enqueue,
    queuedPrompts,
    dequeue,
    clearQueue,
  };
}
