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
import { mkdir, writeFile, rm, readdir, readFile } from 'node:fs/promises';
import { existsSync } from 'node:fs';
import { randomUUID } from 'node:crypto';
import { tmpdir, homedir } from 'node:os';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { execFile } from 'node:child_process';
import { promisify } from 'node:util';
import {
  slugify,
  discoverPluginSkillDirs,
  discoverPluginAgents,
} from './pluginDiscovery.mjs';

const execFileAsync = promisify(execFile);
/** On Windows, npm is npm.cmd — use the .cmd variant when on win32. */
const NPM_CMD = process.platform === 'win32' ? 'npm.cmd' : 'npm';
/** execFileAsync options for npm: shell:true is required on Windows for .cmd files. */
const NPM_EXEC_OPTS = process.platform === 'win32' ? { shell: true } : {};

// ── LSP framing helpers ─────────────────────────────────────────────────────

/**
 * Scan a node_modules directory for skillpm-compatible skill directories.
 * The skillpm spec places skills at: node_modules/<pkg>/skills/<name>/SKILL.md
 * Returns an array of skill subdirectory paths (the directory containing SKILL.md).
 *
 * @param {string} nodeModulesDir - path to node_modules
 * @returns {Promise<string[]>}
 */
async function findSkillDirs(nodeModulesDir) {
  const skillDirs = [];
  let pkgEntries;
  try {
    pkgEntries = await readdir(nodeModulesDir, { withFileTypes: true });
  } catch {
    return skillDirs; // node_modules doesn't exist
  }

  for (const entry of pkgEntries) {
    // Handle scoped packages (@org/...)
    if (entry.isDirectory() && entry.name.startsWith('@')) {
      const scopeDir = join(nodeModulesDir, entry.name);
      let scopedEntries;
      try {
        scopedEntries = await readdir(scopeDir, { withFileTypes: true });
      } catch {
        continue;
      }
      for (const scopedEntry of scopedEntries) {
        if (scopedEntry.isDirectory()) {
          const pkgDir = join(scopeDir, scopedEntry.name);
          const dirs = await findSkillDirsInPackage(pkgDir);
          skillDirs.push(...dirs);
        }
      }
    } else if (entry.isDirectory()) {
      const pkgDir = join(nodeModulesDir, entry.name);
      const dirs = await findSkillDirsInPackage(pkgDir);
      skillDirs.push(...dirs);
    }
  }
  return skillDirs;
}

/**
 * Within a single package directory, find all skills/<name>/ subdirs that contain SKILL.md.
 *
 * @param {string} pkgDir
 * @returns {Promise<string[]>}
 */
async function findSkillDirsInPackage(pkgDir) {
  const skillsRoot = join(pkgDir, 'skills');
  const result = [];
  let skillSubdirs;
  try {
    skillSubdirs = await readdir(skillsRoot, { withFileTypes: true });
  } catch {
    return result; // no skills/ directory in this package
  }
  for (const sub of skillSubdirs) {
    if (sub.isDirectory()) {
      const skillDir = join(skillsRoot, sub.name);
      if (existsSync(join(skillDir, 'SKILL.md'))) {
        result.push(skillDir);
      }
    }
  }
  return result;
}

/** Root directory for bundled skills; each host has its own subdirectory. */
const BUNDLED_SKILLS_ROOT = resolve(dirname(fileURLToPath(import.meta.url)), 'skills');

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

  /** @type {Map<string, string>} Temp skill directories keyed by sessionId for cleanup. */
  const sessionTempDirs = new Map();

  /** @type {Map<string, string[]>} MCP server names keyed by sessionId for stop notifications. */
  const sessionMcpServerNames = new Map();

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

  /** @type {Map<string, { sessionId: string, resolve: (decision: 'approved'|'denied') => void, timer: NodeJS.Timeout }>} */
  const pendingPermissionResponses = new Map();

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

  /** Request explicit permission decision from the browser UI. */
  function requestPermissionDecision(sessionId, request) {
    const requestId = randomUUID();
    sendNotification('permission.request', {
      sessionId,
      requestId,
      request,
    });

    return new Promise(resolve => {
      const timer = setTimeout(() => {
        pendingPermissionResponses.delete(requestId);
        console.warn(`[proxy] permission.request timed out (${requestId}) — default deny`);
        resolve('denied');
      }, 60_000);

      pendingPermissionResponses.set(requestId, {
        sessionId,
        resolve,
        timer,
      });
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
        const {
          host,
          model,
          sessionId,
          systemMessage,
          tools: toolDefs,
          mcpServers,
          availableTools,
          skills,
          disabledSkills,
          customAgents,
        } = params || {};
        console.log(
          `[proxy] session.create requested (host=${host}, model=${model}, sessionId=${sessionId}, tools=${(toolDefs || []).length}, mcpServers=${Object.keys(mcpServers || {}).length}, skills=${(skills || []).length}, customAgents=${(customAgents || []).length})`
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

        // Discover all bundled skill directories for the current host.
        // Skill dirs use the naming convention: <host>/ and <host>-<specialization>/
        // e.g. for host "powerpoint": powerpoint/, powerpoint-deck-builder/, powerpoint-formatting/
        const skillDirectories = [];
        if (host && existsSync(BUNDLED_SKILLS_ROOT)) {
          const hostSlug = slugify(host);
          try {
            const entries = await readdir(BUNDLED_SKILLS_ROOT, { withFileTypes: true });
            for (const entry of entries) {
              if (entry.isDirectory() && (entry.name === hostSlug || entry.name.startsWith(hostSlug + '-'))) {
                skillDirectories.push(join(BUNDLED_SKILLS_ROOT, entry.name));
              }
            }
          } catch {
            // Fall back to single-directory approach if readdir fails
            const hostSkillDir = join(BUNDLED_SKILLS_ROOT, hostSlug);
            if (existsSync(hostSkillDir)) {
              skillDirectories.push(hostSkillDir);
            }
          }
        }

        // Write imported skills to a temp directory so the SDK can load them
        let tempSkillDir = null;
        if (skills && skills.length > 0) {
          tempSkillDir = join(tmpdir(), `oca-skills-${randomUUID()}`);
          await mkdir(tempSkillDir, { recursive: true });
          for (const skill of skills) {
            const skillDir = join(tempSkillDir, slugify(skill.name));
            await mkdir(skillDir, { recursive: true });
            await writeFile(join(skillDir, 'SKILL.md'), skill.content, 'utf8');
          }
          skillDirectories.push(tempSkillDir);
        }

        // Discover skills from installed Copilot CLI plugins (~/.copilot/config.json)
        try {
          const pluginSkillDirs = await discoverPluginSkillDirs(host);
          if (pluginSkillDirs.length > 0) {
            skillDirectories.push(...pluginSkillDirs);
            console.log(`[proxy] Added ${pluginSkillDirs.length} skill dir(s) from installed plugins`);
          }
        } catch (err) {
          console.warn('[proxy] Plugin skill discovery failed:', err.message);
        }

        // Discover agents from installed Copilot CLI plugins
        const allCustomAgents = [...(customAgents || [])];
        try {
          const pluginAgents = await discoverPluginAgents(host);
          if (pluginAgents.length > 0) {
            allCustomAgents.push(...pluginAgents);
            console.log(`[proxy] Added ${pluginAgents.length} agent(s) from installed plugins`);
          }
        } catch (err) {
          console.warn('[proxy] Plugin agent discovery failed:', err.message);
        }

        // Emit 'starting' status for each configured MCP server
        const mcpServerNames = Object.keys(mcpServers || {});
        for (const name of mcpServerNames) {
          sendNotification('mcp.status', { server: name, status: 'starting' });
          sendNotification('mcp.log', {
            server: name,
            timestamp: new Date().toISOString(),
            level: 'info',
            message: `Connecting to MCP server '${name}'...`,
          });
        }

        let session;
        try {
          await ensureStarted();
          session = await client.createSession({
            clientName: 'office-coding-agent',
            model,
            sessionId,
            systemMessage,
            tools,
            mcpServers,
            availableTools,
            skillDirectories,
            disabledSkills: disabledSkills?.length > 0 ? disabledSkills : undefined,
            customAgents: allCustomAgents.length > 0 ? allCustomAgents : undefined,
            onPermissionRequest: async request => {
              console.log(`[proxy] permission.request received: ${request.kind}`);
              // Auto-approve custom-tool permissions — these are tools explicitly
              // registered by the session creator, not built-in filesystem tools.
              // The SDK v0.1.28+ denies all permissions by default; our own tools
              // should always be allowed to execute.
              if (request.kind === 'custom-tool') {
                console.log(`[proxy] permission.request auto-approved: ${request.kind}`);
                return { kind: 'approved' };
              }
              const decision = await requestPermissionDecision(session.sessionId, request);
              console.log(`[proxy] permission.request resolved: ${request.kind} => ${decision}`);
              return { kind: decision };
            },
          });
        } catch (err) {
          // Clean up temp directories on failure
          if (tempSkillDir) {
            void rm(tempSkillDir, { recursive: true, force: true }).catch(() => {});
          }
          // Emit error status for all MCP servers
          for (const name of mcpServerNames) {
            sendNotification('mcp.status', {
              server: name,
              status: 'error',
              error: err.message || 'Session creation failed',
            });
            sendNotification('mcp.log', {
              server: name,
              timestamp: new Date().toISOString(),
              level: 'error',
              message: `Failed to connect: ${err.message || 'unknown error'}`,
            });
          }
          console.error('[proxy] session.create failed:', err);
          sendError(id, -32603, err.message || 'Failed to create session');
          break;
        }

        // Emit 'connected' status for each MCP server on success
        for (const name of mcpServerNames) {
          sendNotification('mcp.status', { server: name, status: 'connected' });
          sendNotification('mcp.log', {
            server: name,
            timestamp: new Date().toISOString(),
            level: 'info',
            message: `Server '${name}' connected successfully`,
          });
        }

        sessions.set(session.sessionId, session);
        if (mcpServerNames.length > 0) {
          sessionMcpServerNames.set(session.sessionId, mcpServerNames);
        }
        if (tempSkillDir) {
          sessionTempDirs.set(session.sessionId, tempSkillDir);
        }
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

      case 'model.switch': {
        const { sessionId, model } = params || {};
        const session = sessions.get(sessionId);
        if (!session) {
          sendError(id, -32602, `Session '${sessionId}' not found`);
          return;
        }
        try {
          await session.setModel(model);
          console.log(`[proxy] model.switch: session ${sessionId} switched to ${model}`);
          sendResponse(id, {});
        } catch (err) {
          console.error(`[proxy] model.switch failed:`, err);
          sendError(id, -32603, err.message || 'Failed to switch model');
        }
        break;
      }

      case 'session.compact': {
        const { sessionId } = params || {};
        const session = sessions.get(sessionId);
        if (!session) {
          sendError(id, -32602, `Session '${sessionId}' not found`);
          return;
        }
        try {
          await session.compact();
          console.log(`[proxy] session.compact: session ${sessionId} compacted`);
          sendResponse(id, {});
        } catch (err) {
          console.error(`[proxy] session.compact failed:`, err);
          sendError(id, -32603, err.message || 'Failed to compact session');
        }
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
          // Emit 'stopped' status for MCP servers in this session
          const mcpNames = sessionMcpServerNames.get(sessionId);
          if (mcpNames) {
            for (const name of mcpNames) {
              sendNotification('mcp.status', { server: name, status: 'stopped' });
              sendNotification('mcp.log', {
                server: name,
                timestamp: new Date().toISOString(),
                level: 'info',
                message: `Server '${name}' stopped`,
              });
            }
            sessionMcpServerNames.delete(sessionId);
          }
          // Clean up temp skill directory
          const tempDir = sessionTempDirs.get(sessionId);
          if (tempDir) {
            sessionTempDirs.delete(sessionId);
            void rm(tempDir, { recursive: true, force: true }).catch(() => {});
          }
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

      case 'permission.respond': {
        const { sessionId, requestId, decision } = params || {};
        const pending = pendingPermissionResponses.get(requestId);
        if (!pending) {
          sendError(id, -32602, `Permission request '${requestId}' not found`);
          return;
        }
        if (pending.sessionId !== sessionId) {
          sendError(
            id,
            -32602,
            `Permission request '${requestId}' does not belong to session '${sessionId}'`
          );
          return;
        }
        const normalizedDecision = decision === 'approved' ? 'approved' : 'denied';
        clearTimeout(pending.timer);
        pendingPermissionResponses.delete(requestId);
        pending.resolve(normalizedDecision);
        sendResponse(id, {});
        break;
      }

      default:
        sendError(id, -32601, `Method '${method}' not supported`);
    }
  }

  // ── Cleanup ─────────────────────────────────────────────────────────────

  let _cleaned = false;
  async function cleanup() {
    if (_cleaned) return;
    _cleaned = true;
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

    for (const pending of pendingPermissionResponses.values()) {
      clearTimeout(pending.timer);
      pending.resolve('denied');
    }
    pendingPermissionResponses.clear();

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

    // Clean up all temp skill directories for this connection
    for (const tempDir of sessionTempDirs.values()) {
      void rm(tempDir, { recursive: true, force: true }).catch(() => {});
    }
    sessionTempDirs.clear();

    // Clear MCP server tracking
    sessionMcpServerNames.clear();
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
