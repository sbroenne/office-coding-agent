/**
 * copilotProxy.mjs — bridge browser WebSocket to @github/copilot-sdk.
 *
 * One shared CopilotClient (singleton) is created when the server starts.
 * Each WebSocket connection gets its own set of sessions backed by that
 * shared client so the CLI process is only spawned once.
 *
 * Source: https://github.com/patniko/github-copilot-office
 */

import { WebSocketServer } from 'ws';
import { CopilotClient } from '@github/copilot-sdk';

// ── LSP framing helpers ─────────────────────────────────────────────────────

/** Wrap a JSON payload in an LSP Content-Length frame. */
function lspFrame(obj) {
  const body = JSON.stringify(obj);
  const len = Buffer.byteLength(body, 'utf8');
  return `Content-Length: ${len}\r\n\r\n${body}`;
}

// ── Singleton CopilotClient ───────────────────────────────────────────────
// One shared client for the lifetime of the server process — avoids spawning
// a new CLI subprocess (and re-authenticating) on every WebSocket connection.
//
// IMPORTANT: The SDK's listModels() does NOT auto-start the client (only
// createSession/resumeSession do). We therefore manage a single start promise
// so every caller awaits the same connect attempt instead of racing.

/** @type {import('@github/copilot-sdk').CopilotClient | null} */
let _sharedClient = null;
/** @type {Promise<void> | null} */
let _startPromise = null;

function getSharedClient() {
  if (!_sharedClient) {
    console.log('[proxy] Creating CopilotClient singleton...');
    _sharedClient = new CopilotClient({ autoStart: false });
  }
  return _sharedClient;
}

/**
 * Ensure the shared client is started. Idempotent — concurrent callers share
 * the same promise so the CLI is only spawned once.
 * @returns {Promise<void>}
 */
function ensureStarted() {
  if (_startPromise) return _startPromise;
  const client = getSharedClient();
  // Already connected — wrap in resolved promise
  if (client.state === 'connected') {
    _startPromise = Promise.resolve();
    return _startPromise;
  }
  console.log('[proxy] Starting Copilot CLI...');
  _startPromise = client
    .start()
    .then(() => {
      console.log('[proxy] Copilot CLI connected.');
    })
    .catch(err => {
      console.warn('[proxy] CLI start failed:', err.message);
      _startPromise = null; // allow a retry on the next request
      throw err;
    });
  return _startPromise;
}

// ── Per-connection handler ──────────────────────────────────────────────────

async function handleConnection(ws) {
  console.log('[proxy] WebSocket client connected (active:', _activeConnections + 1, ')');
  const client = getSharedClient();

  // IMPORTANT: All state and event handlers must be set up synchronously,
  // BEFORE any await. If the message handler is registered after an await,
  // the browser's first message (session.create, sent immediately after the
  // WS open event) can arrive and fire while we're suspended — Node's
  // EventEmitter drops events that have no listener, losing the message
  // permanently and causing a 60 s timeout every time.

  /** @type {Map<string, import('@github/copilot-sdk').CopilotSession>} */
  const sessions = new Map();

  /** @type {Map<string, () => void>} */
  const eventUnsubs = new Map();

  /** Send a JSON-RPC response back to the browser. */
  function sendResponse(id, result) {
    if (ws.readyState === ws.OPEN) {
      ws.send(lspFrame({ jsonrpc: '2.0', id, result }));
    }
  }

  /** Send a JSON-RPC error back to the browser. */
  function sendError(id, code, message) {
    if (ws.readyState === ws.OPEN) {
      ws.send(lspFrame({ jsonrpc: '2.0', id, error: { code, message } }));
    }
  }

  /** Send a JSON-RPC notification (no id) to the browser. */
  function sendNotification(method, params) {
    if (ws.readyState === ws.OPEN) {
      ws.send(lspFrame({ jsonrpc: '2.0', method, params }));
    }
  }

  /**
   * Send a JSON-RPC request to the browser and wait for a response.
   * Used for tool.call (browser executes tools, returns result).
   */
  let nextRequestId = 1;
  /** @type {Map<number, { resolve: Function, reject: Function }>} */
  const pendingRequests = new Map();

  function sendRequest(method, params) {
    return new Promise((resolve, reject) => {
      const id = nextRequestId++;
      pendingRequests.set(id, { resolve, reject });
      if (ws.readyState === ws.OPEN) {
        ws.send(lspFrame({ jsonrpc: '2.0', id, method, params }));
      } else {
        pendingRequests.delete(id);
        reject(new Error('WebSocket closed'));
      }
    });
  }

  // ── Message router ──────────────────────────────────────────────────────

  // Buffer for incomplete LSP messages from the browser
  let buffer = '';

  ws.on('message', rawData => {
    buffer += typeof rawData === 'string' ? rawData : rawData.toString('utf8');

    // Process all complete LSP frames in the buffer
    while (true) {
      const headerEnd = buffer.indexOf('\r\n\r\n');
      if (headerEnd === -1) break;

      const header = buffer.slice(0, headerEnd);
      const match = header.match(/Content-Length:\s*(\d+)/i);
      if (!match) {
        buffer = buffer.slice(headerEnd + 4);
        continue;
      }

      const contentLength = parseInt(match[1], 10);
      const contentStart = headerEnd + 4;
      const messageEnd = contentStart + contentLength;

      if (buffer.length < messageEnd) break; // incomplete — wait for more data

      const body = buffer.slice(contentStart, messageEnd);
      buffer = buffer.slice(messageEnd);

      let msg;
      try {
        msg = JSON.parse(body);
      } catch {
        continue;
      }

      // JSON-RPC response (from browser answering our tool.call request)
      if ('result' in msg || 'error' in msg) {
        const pending = pendingRequests.get(msg.id);
        if (pending) {
          pendingRequests.delete(msg.id);
          if (msg.error) {
            pending.reject(new Error(msg.error.message || 'RPC error'));
          } else {
            pending.resolve(msg.result);
          }
        }
        continue;
      }

      // JSON-RPC request (from browser calling proxy methods)
      void handleMethod(msg).catch(err => {
        if (msg.id != null) {
          sendError(msg.id, -32603, err.message || 'Internal error');
        }
      });
    }
  });

  async function handleMethod(msg) {
    const { id, method, params } = msg;

    switch (method) {
      case 'session.create': {
        const { model, sessionId, systemMessage, tools: toolDefs } = params || {};
        console.log(
          `[proxy] session.create requested (model=${model}, sessionId=${sessionId}, tools=${(toolDefs || []).length})`
        );
        // Build SDK Tool[] with handlers that forward tool calls to the browser
        const tools = (toolDefs || []).map(t => ({
          name: t.name,
          description: t.description,
          parameters: t.parameters,
          handler: async (args, invocation) => {
            const response = await sendRequest('tool.call', {
              sessionId: invocation.sessionId,
              toolCallId: invocation.toolCallId,
              toolName: invocation.toolName,
              arguments: args,
            });
            return response.result;
          },
        }));

        let session;
        try {
          await ensureStarted();
          session = await client.createSession({
            model,
            sessionId,
            systemMessage,
            tools,
          });
        } catch (err) {
          console.error('[proxy] session.create failed:', err);
          sendError(id, -32603, err.message || 'Failed to create session');
          break;
        }

        sessions.set(session.sessionId, session);
        markHealthy();
        console.log(`[proxy] session.create succeeded (sessionId=${session.sessionId})`); 

        // Subscribe to all session events and forward them to the browser
        const unsub = session.on(event => {
          sendNotification('session.event', {
            sessionId: session.sessionId,
            event,
          });
        });
        eventUnsubs.set(session.sessionId, unsub);

        sendResponse(id, { sessionId: session.sessionId });
        break;
      }

      case 'session.send': {
        const { sessionId, prompt, attachments, mode } = params || {};
        const session = sessions.get(sessionId);
        if (!session) {
          sendError(id, -32602, `Session '${sessionId}' not found`);
          return;
        }
        const messageId = await session.send({ prompt, attachments, mode });
        sendResponse(id, { messageId });
        break;
      }

      case 'session.destroy': {
        const { sessionId } = params || {};
        const session = sessions.get(sessionId);
        if (session) {
          const unsub = eventUnsubs.get(sessionId);
          unsub?.();
          eventUnsubs.delete(sessionId);
          await session.destroy();
          sessions.delete(sessionId);
        }
        sendResponse(id, {});
        break;
      }

      case 'models.list': {
        let models;
        try {
          await ensureStarted();
          models = await client.listModels();
          console.log(
            `[proxy] models.list returned ${models.length} model(s):`,
            models.map(m => m.id)
          );
        } catch (err) {
          console.error('[proxy] models.list failed:', err);
          sendError(id, -32603, err.message || 'Failed to list models');
          break;
        }
        sendResponse(id, { models });
        break;
      }

      default:
        sendError(id, -32601, `Method '${method}' not supported`);
    }
  }

  // ── Cleanup ─────────────────────────────────────────────────────────────

  async function cleanup() {
    _activeConnections--;
    for (const unsub of eventUnsubs.values()) {
      unsub();
    }
    eventUnsubs.clear();

    // Reject any pending tool.call promises — without this they hang forever
    // because the browser that was supposed to reply has disconnected.
    for (const pending of pendingRequests.values()) {
      pending.reject(new Error('WebSocket disconnected'));
    }
    pendingRequests.clear();

    // Destroy server-side sessions so the shared CopilotClient doesn't
    // accumulate open sessions across reconnects.
    for (const session of sessions.values()) {
      try {
        await session.destroy();
      } catch {
        // Session may already be gone — ignore.
      }
    }
    sessions.clear();
  }

  ws.on('close', () => {
    console.log('[proxy] WebSocket client disconnected');
    void cleanup();
  });
  ws.on('error', err => {
    console.error('[proxy] WebSocket error:', err);
    void cleanup();
  });

  // All event handlers are registered synchronously above — do async init last.
  // Non-fatal: individual method handlers also call ensureStarted() on demand.
  _activeConnections++;
  await ensureStarted().catch(() => {});
}

// ── Health tracking ───────────────────────────────────────────────────────

let _lastHealthy = 0;
let _activeConnections = 0;

function markHealthy() {
  _lastHealthy = Date.now();
}

/**
 * Lightweight health check: returns whether any WebSocket client
 * has successfully created a Copilot session recently (within 5 min).
 * Does NOT spawn new CLI subprocesses.
 */
export function checkCopilotHealth() {
  const staleMs = 5 * 60 * 1000;
  const ok = _activeConnections > 0 && Date.now() - _lastHealthy < staleMs;
  return { ok, activeConnections: _activeConnections };
}

// ── Setup ─────────────────────────────────────────────────────────────────

export function setupCopilotProxy(httpsServer) {
  // Kick off the CLI start immediately so it's ready before the first browser
  // connection.  The promise is cached — all subsequent calls share it.
  ensureStarted()
    .then(() => getSharedClient().listModels())
    .then(models => {
      console.log(`[proxy] CLI ready — ${models.length} model(s) available`);
    })
    .catch(err => {
      console.warn('[proxy] CLI warm-up failed (will retry on first connection):', err.message);
    });

  const wss = new WebSocketServer({ noServer: true });

  const upgradeHandler = (request, socket, head) => {
    const url = new URL(request.url, `https://${request.headers.host}`);

    if (url.pathname === '/api/copilot') {
      wss.handleUpgrade(request, socket, head, ws => {
        wss.emit('connection', ws, request);
      });
    }
    // Let other WebSocket connections (e.g., Vite HMR) pass through
  };

  httpsServer.on('upgrade', upgradeHandler);

  httpsServer.closeWebSockets = () => {
    wss.clients.forEach(client => client.terminate());
    wss.close();
  };

  wss.on('connection', ws => void handleConnection(ws));
}
