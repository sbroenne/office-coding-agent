/*---------------------------------------------------------------------------------------------
 *  WebSocket-based CopilotClient for browser environments
 *  Connects to the Copilot CLI via WebSocket proxy (src/server.mjs)
 *  Source: https://github.com/patniko/github-copilot-office
 *--------------------------------------------------------------------------------------------*/

import { createMessageConnection, type MessageConnection } from 'vscode-jsonrpc';
import { WebSocketMessageReader, WebSocketMessageWriter } from './websocket-transport';
import type {
  SessionConfig,
  SessionEvent,
  SessionEventHandler,
  MessageOptions,
  Tool,
  ToolHandler,
  ToolInvocation,
  ToolResultObject,
  PermissionRequestResult,
} from '@github/copilot-sdk';

interface ToolCallRequestPayload {
  sessionId: string;
  toolCallId: string;
  toolName: string;
  arguments: unknown;
}

interface ToolCallResponsePayload {
  result: ToolResultObject;
}

export interface PermissionRequestPayload {
  sessionId: string;
  requestId: string;
  request: {
    kind: string;
    intention?: string;
    fullCommandText?: string;
    commands?: readonly { identifier: string }[];
    fileName?: string;
    diff?: string;
    path?: string;
    serverName?: string;
    toolName?: string;
    args?: unknown;
    [key: string]: unknown;
  };
  promptRequest?: Record<string, unknown>;
  locationKey?: string;
}

export type SessionMode = 'interactive' | 'plan' | 'autopilot';

export interface PlanState {
  exists: boolean;
  content: string | null;
  path: string | null;
}

export interface ExitPlanModeRequestPayload {
  sessionId: string;
  requestId: string;
  actions: string[];
  planContent: string;
  recommendedAction: string;
  summary: string;
}

/** Extended session config for browser → proxy communication. */
export interface BrowserSessionConfig extends Omit<SessionConfig, 'tools' | 'onPermissionRequest'> {
  tools?: Tool[];
  /** Office host identifier (e.g. 'excel', 'powerpoint'). Used by proxy for per-host skill loading. */
  host?: string;
}

export interface AgentInfo {
  name: string;
  displayName: string;
  description: string;
}

/**
 * Browser-compatible Copilot session over WebSocket.
 */
export class BrowserCopilotSession {
  private eventHandlers = new Set<SessionEventHandler>();
  private toolHandlers = new Map<string, ToolHandler>();
  private permissionHandlers = new Set<(payload: PermissionRequestPayload) => void>();

  constructor(
    public readonly sessionId: string,
    private connection: MessageConnection,
    public readonly workspacePath?: string
  ) {}

  async send(options: MessageOptions): Promise<string> {
    const response = await this.connection.sendRequest<{ messageId: string }>('session.send', {
      sessionId: this.sessionId,
      prompt: options.prompt,
      attachments: options.attachments,
      mode: options.mode,
    });
    return response.messageId;
  }

  /** Send a prompt and iterate over response events. */
  async *query(options: MessageOptions): AsyncGenerator<SessionEvent, void, undefined> {
    const queue: SessionEvent[] = [];
    let resolve: (() => void) | null = null;
    let done = false;
    let sendError: Error | undefined;

    const unsubscribe = this.on(event => {
      queue.push(event);
      resolve?.();
      if (event.type === 'session.idle') {
        done = true;
      }
    });

    this.send(options).catch(err => {
      sendError = err instanceof Error ? err : new Error(String(err));
      done = true;
      resolve?.();
    });

    try {
      while (!done || queue.length > 0) {
        if (queue.length > 0) {
          const item = queue.shift();
          if (item !== undefined) yield item;
        } else {
          await new Promise<void>(r => {
            resolve = r;
          });
          resolve = null;
        }
      }
      if (sendError !== undefined) throw sendError;
    } finally {
      unsubscribe();
    }
  }

  on(handler: SessionEventHandler): () => void {
    this.eventHandlers.add(handler);
    return () => {
      this.eventHandlers.delete(handler);
    };
  }

  _dispatchEvent(event: SessionEvent): void {
    for (const handler of this.eventHandlers) {
      try {
        handler(event);
      } catch {
        // ignore
      }
    }
  }

  registerTools(tools?: Tool[]): void {
    this.toolHandlers.clear();
    if (tools) {
      for (const tool of tools) {
        if (tool.handler) {
          this.toolHandlers.set(tool.name, tool.handler);
        }
      }
    }
  }

  getToolHandler(name: string): ToolHandler | undefined {
    return this.toolHandlers.get(name);
  }

  onPermissionRequest(handler: (payload: PermissionRequestPayload) => void): () => void {
    this.permissionHandlers.add(handler);
    return () => {
      this.permissionHandlers.delete(handler);
    };
  }

  _dispatchPermissionRequest(payload: PermissionRequestPayload): void {
    for (const handler of this.permissionHandlers) {
      try {
        handler(payload);
      } catch {
        // ignore
      }
    }
  }

  async respondPermission(requestId: string, decision: PermissionRequestResult): Promise<void> {
    await this.connection.sendRequest('permission.respond', {
      sessionId: this.sessionId,
      requestId,
      decision,
    });
  }

  async setModel(modelId: string): Promise<void> {
    await this.connection.sendRequest('model.switch', {
      sessionId: this.sessionId,
      model: modelId,
    });
  }

  async listAgents(): Promise<AgentInfo[]> {
    const result = await this.connection.sendRequest<{ agents: AgentInfo[] }>('agent.list', {
      sessionId: this.sessionId,
    });
    return result.agents;
  }

  async selectAgent(agentName: string): Promise<AgentInfo> {
    const result = await this.connection.sendRequest<{ agent: AgentInfo }>('agent.select', {
      sessionId: this.sessionId,
      name: agentName,
    });
    return result.agent;
  }

  async deselectAgent(): Promise<void> {
    await this.connection.sendRequest('agent.deselect', {
      sessionId: this.sessionId,
    });
  }

  async compact(): Promise<void> {
    await this.connection.sendRequest('session.compact', {
      sessionId: this.sessionId,
    });
  }

  async abort(): Promise<void> {
    await this.connection.sendRequest('session.abort', {
      sessionId: this.sessionId,
    });
  }

  async disconnect(): Promise<void> {
    await this.connection.sendRequest('session.disconnect', {
      sessionId: this.sessionId,
    });
    this.eventHandlers.clear();
    this.toolHandlers.clear();
    this.permissionHandlers.clear();
  }

  async getMode(): Promise<SessionMode> {
    const result = await this.connection.sendRequest<{ mode: SessionMode }>('session.mode.get', {
      sessionId: this.sessionId,
    });
    return result.mode;
  }

  async setMode(mode: SessionMode): Promise<void> {
    await this.connection.sendRequest('session.mode.set', {
      sessionId: this.sessionId,
      mode,
    });
  }

  async readPlan(): Promise<PlanState> {
    return this.connection.sendRequest<PlanState>('session.plan.read', {
      sessionId: this.sessionId,
    });
  }

  async updatePlan(content: string): Promise<void> {
    await this.connection.sendRequest('session.plan.update', {
      sessionId: this.sessionId,
      content,
    });
  }

  async deletePlan(): Promise<void> {
    await this.connection.sendRequest('session.plan.delete', {
      sessionId: this.sessionId,
    });
  }

  async setApproveAll(enabled: boolean): Promise<void> {
    await this.connection.sendRequest('permissions.setApproveAll', {
      sessionId: this.sessionId,
      enabled,
    });
  }

  async resetSessionApprovals(): Promise<void> {
    await this.connection.sendRequest('permissions.resetSessionApprovals', {
      sessionId: this.sessionId,
    });
  }

  async initiateMcpOAuth(serverName: string): Promise<McpOAuthResult> {
    return this.connection.sendRequest<McpOAuthResult>('mcp.initiateOAuth', {
      sessionId: this.sessionId,
      serverName,
    });
  }

  async initiateMcpOAuthWithHint(serverName: string, loginHint?: string): Promise<McpOAuthResult> {
    return this.connection.sendRequest<McpOAuthResult>('mcp.initiateOAuth', {
      sessionId: this.sessionId,
      serverName,
      loginHint,
    });
  }

  async destroy(): Promise<void> {
    await this.disconnect();
  }
}

/** MCP status notification payload */
export interface McpStatusPayload {
  server: string;
  status:
    | 'stopped'
    | 'starting'
    | 'connected'
    | 'needs-auth'
    | 'pending'
    | 'disabled'
    | 'not_configured'
    | 'failed'
    | 'error';
  error?: string;
}

export type McpOAuthResult =
  | { status: 'success'; authorizationUrl?: string; oauthAlias?: string }
  | { status: 'public' }
  | { status: 'error'; message: string };

export interface McpOAuthRequiredPayload {
  requestId: string;
  serverName: string;
  serverUrl?: string;
}

export interface McpOAuthCompletedPayload {
  requestId: string;
}

/** MCP log notification payload */
export interface McpLogPayload {
  server: string;
  timestamp: string;
  level: 'info' | 'warn' | 'error';
  message: string;
}

/** MCP tools notification payload */
export interface McpToolsPayload {
  server: string;
  tools: { name: string; description: string }[];
}

/**
 * Browser-compatible Copilot client connected via WebSocket proxy.
 */
export class WebSocketCopilotClient {
  private connection: MessageConnection | null = null;
  private wsSocket: WebSocket | null = null;
  private sessions = new Map<string, BrowserCopilotSession>();
  private mcpStatusHandlers = new Set<(payload: McpStatusPayload) => void>();
  private mcpLogHandlers = new Set<(payload: McpLogPayload) => void>();
  private mcpToolsHandlers = new Set<(payload: McpToolsPayload) => void>();
  private mcpOAuthRequiredHandlers = new Set<(payload: McpOAuthRequiredPayload) => void>();
  private mcpOAuthCompletedHandlers = new Set<(payload: McpOAuthCompletedPayload) => void>();

  constructor(private url: string) {}

  async start(): Promise<void> {
    if (this.connection) return;

    await new Promise<void>((resolve, reject) => {
      this.wsSocket = new WebSocket(this.url);

      this.wsSocket.addEventListener('open', () => {
        console.log('[ws] Connected to', this.url);
        const socket = this.wsSocket;
        if (!socket) return;
        const reader = new WebSocketMessageReader(socket);
        const writer = new WebSocketMessageWriter(socket);
        this.connection = createMessageConnection(reader, writer);
        this.attachConnectionHandlers();
        this.connection.listen();
        resolve();
      });

      this.wsSocket.addEventListener('error', event => {
        console.error('[ws] Connection error to', this.url, event);
        reject(new Error(`Failed to connect to ${this.url}`));
      });
    });
  }

  async createSession(config: BrowserSessionConfig = {}): Promise<BrowserCopilotSession> {
    if (!this.connection) {
      throw new Error('Client not connected. Call start() first.');
    }

    const response = await this.connection.sendRequest<{
      sessionId: string;
      workspacePath?: string;
    }>('session.create', {
      model: config.model,
      sessionId: config.sessionId,
      systemMessage: config.systemMessage,
      tools: config.tools?.map(tool => ({
        name: tool.name,
        description: tool.description,
        parameters: tool.parameters,
        skipPermission: tool.skipPermission,
      })),
      mcpServers: config.mcpServers,
      availableTools: config.availableTools,
      host: config.host,
      agent: config.agent,
    });

    const sessionId = response.sessionId;
    const session = new BrowserCopilotSession(sessionId, this.connection, response.workspacePath);
    session.registerTools(config.tools);
    this.sessions.set(sessionId, session);
    return session;
  }

  async listModels(): Promise<ListModelsResult[]> {
    if (!this.connection) {
      throw new Error('Client not connected. Call start() first.');
    }
    const result = await this.connection.sendRequest<{ models: ListModelsResult[] }>(
      'models.list',
      {}
    );
    return result.models;
  }

  onMcpStatus(handler: (payload: McpStatusPayload) => void): () => void {
    this.mcpStatusHandlers.add(handler);
    return () => {
      this.mcpStatusHandlers.delete(handler);
    };
  }

  onMcpLog(handler: (payload: McpLogPayload) => void): () => void {
    this.mcpLogHandlers.add(handler);
    return () => {
      this.mcpLogHandlers.delete(handler);
    };
  }

  onMcpTools(handler: (payload: McpToolsPayload) => void): () => void {
    this.mcpToolsHandlers.add(handler);
    return () => {
      this.mcpToolsHandlers.delete(handler);
    };
  }

  onMcpOAuthRequired(handler: (payload: McpOAuthRequiredPayload) => void): () => void {
    this.mcpOAuthRequiredHandlers.add(handler);
    return () => {
      this.mcpOAuthRequiredHandlers.delete(handler);
    };
  }

  onMcpOAuthCompleted(handler: (payload: McpOAuthCompletedPayload) => void): () => void {
    this.mcpOAuthCompletedHandlers.add(handler);
    return () => {
      this.mcpOAuthCompletedHandlers.delete(handler);
    };
  }

  async stop(): Promise<void> {
    for (const session of this.sessions.values()) {
      try {
        await session.disconnect();
      } catch {
        // ignore
      }
    }
    this.sessions.clear();

    if (this.connection) {
      this.connection.dispose();
      this.connection = null;
    }

    if (this.wsSocket) {
      this.wsSocket.close();
      this.wsSocket = null;
    }
  }

  private attachConnectionHandlers(): void {
    if (!this.connection) return;

    this.connection.onNotification('session.event', (notification: unknown) => {
      const n = notification as { sessionId?: string; event?: SessionEvent };
      if (n.sessionId && n.event) {
        this.sessions.get(n.sessionId)?._dispatchEvent(n.event);
      }
    });

    this.connection.onNotification('permission.request', (notification: unknown) => {
      const payload = notification as PermissionRequestPayload;
      if (!payload?.sessionId || !payload?.requestId) return;
      this.sessions.get(payload.sessionId)?._dispatchPermissionRequest(payload);
    });

    this.connection.onNotification('mcp.status', (notification: unknown) => {
      const payload = notification as McpStatusPayload;
      for (const handler of this.mcpStatusHandlers) {
        try {
          handler(payload);
        } catch {
          /* ignore */
        }
      }
    });

    this.connection.onNotification('mcp.log', (notification: unknown) => {
      const payload = notification as McpLogPayload;
      for (const handler of this.mcpLogHandlers) {
        try {
          handler(payload);
        } catch {
          /* ignore */
        }
      }
    });

    this.connection.onNotification('mcp.tools', (notification: unknown) => {
      const payload = notification as McpToolsPayload;
      for (const handler of this.mcpToolsHandlers) {
        try {
          handler(payload);
        } catch {
          /* ignore */
        }
      }
    });

    this.connection.onNotification('mcp.oauth-required', (notification: unknown) => {
      const payload = notification as McpOAuthRequiredPayload;
      for (const handler of this.mcpOAuthRequiredHandlers) {
        try {
          handler(payload);
        } catch {
          /* ignore */
        }
      }
    });

    this.connection.onNotification('mcp.oauth-completed', (notification: unknown) => {
      const payload = notification as McpOAuthCompletedPayload;
      for (const handler of this.mcpOAuthCompletedHandlers) {
        try {
          handler(payload);
        } catch {
          /* ignore */
        }
      }
    });

    this.connection.onRequest(
      'tool.call',
      async (params: ToolCallRequestPayload): Promise<ToolCallResponsePayload> => {
        const session = this.sessions.get(params.sessionId);
        const handler = session?.getToolHandler(params.toolName);
        if (!handler) {
          return {
            result: {
              textResultForLlm: `Tool '${params.toolName}' not supported`,
              resultType: 'failure',
              error: `tool '${params.toolName}' not supported`,
              toolTelemetry: {},
            },
          };
        }
        try {
          const invocation: ToolInvocation = {
            sessionId: params.sessionId,
            toolCallId: params.toolCallId,
            toolName: params.toolName,
            arguments: params.arguments,
          };
          const result = await handler(params.arguments, invocation);
          return { result: result as ToolResultObject };
        } catch (error) {
          const message = error instanceof Error ? error.message : String(error);
          return {
            result: {
              textResultForLlm: message,
              resultType: 'failure' as const,
              error: message,
              toolTelemetry: {},
            },
          };
        }
      }
    );
  }
}

/** Model info returned by the Copilot CLI */
export interface ListModelsResult {
  id: string;
  name: string;
}

/** Creates and connects a WebSocketCopilotClient. */
export async function createWebSocketClient(url: string): Promise<WebSocketCopilotClient> {
  const client = new WebSocketCopilotClient(url);
  await client.start();
  return client;
}
