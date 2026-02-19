/*---------------------------------------------------------------------------------------------
 *  WebSocket-based CopilotClient for browser environments
 *  Connects to the Copilot CLI via WebSocket proxy (src/server.js)
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

/**
 * Browser-compatible Copilot session over WebSocket.
 */
export class BrowserCopilotSession {
  private eventHandlers = new Set<SessionEventHandler>();
  private toolHandlers = new Map<string, ToolHandler>();

  constructor(
    public readonly sessionId: string,
    private connection: MessageConnection
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

    const unsubscribe = this.on(event => {
      queue.push(event);
      resolve?.();
      if (event.type === 'session.idle') {
        done = true;
      }
    });

    this.send(options).catch(() => {
      done = true;
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
        this.toolHandlers.set(tool.name, tool.handler);
      }
    }
  }

  getToolHandler(name: string): ToolHandler | undefined {
    return this.toolHandlers.get(name);
  }

  async destroy(): Promise<void> {
    await this.connection.sendRequest('session.destroy', {
      sessionId: this.sessionId,
    });
    this.eventHandlers.clear();
    this.toolHandlers.clear();
  }
}

/**
 * Browser-compatible Copilot client connected via WebSocket proxy.
 */
export class WebSocketCopilotClient {
  private connection: MessageConnection | null = null;
  private wsSocket: WebSocket | null = null;
  private sessions = new Map<string, BrowserCopilotSession>();

  constructor(private url: string) {}

  async start(): Promise<void> {
    if (this.connection) return;

    await new Promise<void>((resolve, reject) => {
      this.wsSocket = new WebSocket(this.url);

      this.wsSocket.addEventListener('open', () => {
        const socket = this.wsSocket;
        if (!socket) return;
        const reader = new WebSocketMessageReader(socket);
        const writer = new WebSocketMessageWriter(socket);
        this.connection = createMessageConnection(reader, writer);
        this.attachConnectionHandlers();
        this.connection.listen();
        resolve();
      });

      this.wsSocket.addEventListener('error', () => {
        reject(new Error(`Failed to connect to ${this.url}`));
      });
    });
  }

  async createSession(config: SessionConfig = {}): Promise<BrowserCopilotSession> {
    if (!this.connection) {
      throw new Error('Client not connected. Call start() first.');
    }

    const response = await this.connection.sendRequest<{ sessionId: string }>('session.create', {
      model: config.model,
      sessionId: config.sessionId,
      systemMessage: config.systemMessage,
      tools: config.tools?.map(tool => ({
        name: tool.name,
        description: tool.description,
        parameters: tool.parameters,
      })),
    });

    const sessionId = response.sessionId;
    const session = new BrowserCopilotSession(sessionId, this.connection);
    session.registerTools(config.tools);
    this.sessions.set(sessionId, session);
    return session;
  }

  async stop(): Promise<void> {
    for (const session of this.sessions.values()) {
      try {
        await session.destroy();
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

/** Creates and connects a WebSocketCopilotClient. */
export async function createWebSocketClient(url: string): Promise<WebSocketCopilotClient> {
  const client = new WebSocketCopilotClient(url);
  await client.start();
  return client;
}
