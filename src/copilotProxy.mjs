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
import { mkdir, writeFile, readdir, readFile } from 'node:fs/promises';
import { existsSync } from 'node:fs';
import { randomUUID } from 'node:crypto';
import { homedir } from 'node:os';
import { join } from 'node:path';
import { execFile } from 'node:child_process';
import { promisify } from 'node:util';
import {
  slugify,
  discoverPluginSkillDirs,
  discoverPluginSkillObjects,
  discoverPluginAgents,
  discoverPluginPrompts,
  readMcpServersForPlugin,
  listAllPluginConfigs,
  readPluginManifest,
  isPluginForHost,
} from './pluginDiscovery.mjs';
import { getPluginManager } from './plugins/pluginManager.mjs';
import { isTrustedRequestOrigin } from './serverSecurity.mjs';

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

/** Wrap a JSON payload in an LSP Content-Length frame. */
function lspFrame(obj) {
  const body = JSON.stringify(obj);
  const len = Buffer.byteLength(body, 'utf8');
  return `Content-Length: ${len}\r\n\r\n${body}`;
}

/**
 * Return the set of tool names actually active for this session.
 *
 * If availableTools is provided, it narrows the session-level tool list.
 *
 * @param {Array<{name?: string}>} [toolDefs]
 * @param {string[]} [availableTools]
 * @returns {string[]}
 */
export function getRegisteredToolNames(toolDefs = [], availableTools) {
  const definedToolNames = toolDefs
    .map(tool => (typeof tool?.name === 'string' ? tool.name : ''))
    .filter(Boolean);
  const uniqueToolNames = Array.from(new Set(definedToolNames));

  if (!Array.isArray(availableTools) || availableTools.length === 0) {
    return uniqueToolNames;
  }

  const allowedToolNames = new Set(
    availableTools.filter(toolName => typeof toolName === 'string' && toolName.length > 0)
  );
  return uniqueToolNames.filter(toolName => allowedToolNames.has(toolName));
}

/**
 * Ensure every custom agent gets an explicit tool allowlist before it reaches the SDK.
 *
 * Plugin- and browser-provided agents often omit `tools`, intending "all active session
 * tools". The SDK expects an explicit allowlist for custom agents, so we expand omitted,
 * null, empty, or "*" values to the current session tool names here.
 *
 * @param {Array<{tools?: string[] | null}>} [customAgents]
 * @param {string[]} [sessionToolNames]
 * @returns {Array<{tools?: string[]}>}
 */
export function applySessionToolAccessToCustomAgents(customAgents = [], sessionToolNames = []) {
  const fallbackToolNames = Array.from(
    new Set(sessionToolNames.filter(toolName => typeof toolName === 'string' && toolName.length > 0))
  );

  return customAgents.map(agent => {
    const requestedTools = Array.isArray(agent?.tools)
      ? Array.from(
          new Set(
            agent.tools.filter(toolName => typeof toolName === 'string' && toolName.length > 0)
          )
        )
      : null;

    if (!requestedTools || requestedTools.length === 0 || requestedTools.includes('*')) {
      return {
        ...agent,
        tools: fallbackToolNames.length > 0 ? fallbackToolNames : undefined,
      };
    }

    if (fallbackToolNames.length === 0) {
      return { ...agent, tools: requestedTools };
    }

    return {
      ...agent,
      tools: requestedTools.filter(toolName => fallbackToolNames.includes(toolName)),
    };
  });
}

export function mergePluginMcpServers(baseServers = {}, pluginServers = {}) {
  const merged = { ...(baseServers || {}) };
  for (const [name, config] of Object.entries(pluginServers || {})) {
    if (merged[name]) continue;
    merged[name] = config;
  }
  return merged;
}

async function discoverPluginMcpServers(host, configPath, disabledMcpServerNames = []) {
  const plugins = await listAllPluginConfigs(configPath);
  const disabled = new Set(disabledMcpServerNames);
  const servers = {};
  for (const plugin of plugins) {
    if (!plugin.enabled || !plugin.cache_path) continue;
    if (host && !isPluginForHost(plugin.name, host)) continue;
    const manifest = await readPluginManifest(plugin.cache_path);
    const pluginServers = await readMcpServersForPlugin(plugin, manifest);
    for (const [name, config] of Object.entries(pluginServers)) {
      if (disabled.has(name) || servers[name]) continue;
      servers[name] = config;
    }
  }
  return servers;
}

/**
 * Merge discovered plugin skill content into the system message as a defensive
 * fallback when the SDK does not surface `skillDirectories` strongly enough for
 * prompt-only validation scenarios.
 *
 * @param {{ mode?: string, content?: string } | undefined} systemMessage
 * @param {Array<{name: string, description?: string, content?: string}>} pluginSkills
 * @param {string[] | undefined} disabledSkills
 * @returns {{ mode: string, content: string } | undefined}
 */
export function mergePluginSkillsIntoSystemMessage(
  systemMessage,
  pluginSkills = [],
  disabledSkills = []
) {
  const disabledSkillNames = new Set(
    Array.isArray(disabledSkills)
      ? disabledSkills.filter(skillName => typeof skillName === 'string' && skillName.length > 0)
      : []
  );
  const enabledPluginSkills = pluginSkills.filter(
    skill =>
      typeof skill?.content === 'string' &&
      skill.content.trim().length > 0 &&
      !disabledSkillNames.has(skill.name)
  );

  if (enabledPluginSkills.length === 0) {
    return systemMessage;
  }

  const skillContext = enabledPluginSkills
    .map(skill => {
      const header = skill.description
        ? `### ${skill.name}\n${skill.description}`
        : `### ${skill.name}`;
      return `${header}\n${skill.content.trim()}`;
    })
    .join('\n\n');

  const existingContent = systemMessage?.content ?? '';
  const mergedContent = existingContent
    ? `${existingContent}\n\n## Plugin Skill Context\n\n${skillContext}`
    : `## Plugin Skill Context\n\n${skillContext}`;

  return {
    mode: systemMessage?.mode ?? 'replace',
    content: mergedContent,
  };
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
          systemMessage: systemMessageParam,
          tools: toolDefs,
          mcpServers,
          availableTools,
          disabledSkills,
          disabledMcpServerNames,
          customAgents,
          pluginConfigPath,
        } = params || {};
        let systemMessage = systemMessageParam;
        console.log(
          `[proxy] session.create requested (host=${host}, model=${model}, sessionId=${sessionId}, tools=${(toolDefs || []).length}, mcpServers=${Object.keys(mcpServers || {}).length}, customAgents=${(customAgents || []).length})`
        );
        const sessionToolNames = getRegisteredToolNames(toolDefs || [], availableTools);
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

        const skillDirectories = [];
        let pluginSkills = [];
        let pluginAgents = [];
        let pluginPrompts = [];
        let pluginMcpServers = {};

        // Discover plugin inputs from either a test override config or the sandboxed plugin manager.
        try {
          if (pluginConfigPath) {
            const [pluginSkillDirs, discoveredSkills, discoveredAgents, discoveredPrompts, discoveredMcp] =
              await Promise.all([
                discoverPluginSkillDirs(host, pluginConfigPath),
                discoverPluginSkillObjects(host, pluginConfigPath),
                discoverPluginAgents(host, pluginConfigPath),
                discoverPluginPrompts(host, pluginConfigPath),
                discoverPluginMcpServers(host, pluginConfigPath, disabledMcpServerNames),
              ]);
            skillDirectories.push(...pluginSkillDirs);
            pluginSkills = discoveredSkills;
            pluginAgents = discoveredAgents;
            pluginPrompts = discoveredPrompts;
            pluginMcpServers = discoveredMcp;
          } else {
            const manager = await getPluginManager();
            const inputs = manager.getSessionInputs(host, disabledMcpServerNames);
            skillDirectories.push(...inputs.skillDirectories);
            pluginSkills = inputs.skills;
            pluginAgents = inputs.customAgents;
            pluginPrompts = inputs.prompts;
            pluginMcpServers = inputs.mcpServers;
          }
          if (skillDirectories.length > 0) {
            console.log(`[proxy] Added ${skillDirectories.length} skill dir(s) from installed plugins`);
          }
          if (Object.keys(pluginMcpServers).length > 0) {
            console.log(`[proxy] Added ${Object.keys(pluginMcpServers).length} MCP server(s) from installed plugins`);
          }
        } catch (err) {
          console.warn('[proxy] Plugin discovery failed:', err.message);
        }

        // Discover plugin skills once so they can be both surfaced to the browser and
        // injected into the initial system message as a compatibility fallback.
        try {
          if (pluginSkills.length > 0) {
            console.log(`[proxy] Sending ${pluginSkills.length} plugin skill(s) to browser`);
            sendNotification('plugin.skills', { skills: pluginSkills });
            systemMessage = mergePluginSkillsIntoSystemMessage(
              systemMessage,
              pluginSkills,
              disabledSkills
            );
          }
        } catch (err) {
          console.warn('[proxy] Plugin skill notification failed:', err.message);
        }

        // Discover agents from installed Copilot CLI plugins and merge into systemMessage.
        // The SDK's customAgents are VS Code agent-picker entries, NOT auto-applied system
        // prompts — they require explicit @mention. To guarantee plugin agent instructions
        // reach the model, we append them to the systemMessage content instead.
        //
        // We also send a plugin.agents notification so the browser-side agentService can
        // register these agents in the AgentPicker without requiring a manual import.
        const allCustomAgents = applySessionToolAccessToCustomAgents(
          customAgents || [],
          sessionToolNames
        );
        try {
          if (pluginAgents.length > 0) {
            console.log(`[proxy] Merging ${pluginAgents.length} plugin agent(s) into system message`);

            // Notify the browser about discovered plugin agents so they can be shown
            // in the AgentPicker. Send before the session.create response so the browser
            // can update its agentService state promptly.
            sendNotification('plugin.agents', {
              agents: pluginAgents.map(a => ({
                name: a.name,
                description: a.description,
                prompt: a.prompt,
                hosts: a.hosts,
              })),
            });

            // Pick the first plugin agent as the default; additional ones go into customAgents
            // for agent-picker UI (future @mention support).
            const [defaultPluginAgent, ...extraAgents] = pluginAgents;
            if (defaultPluginAgent?.prompt) {
              const existingContent = systemMessage?.content ?? '';
              const merged = existingContent
                ? `${existingContent}\n\n${defaultPluginAgent.prompt}`
                : defaultPluginAgent.prompt;
              systemMessage = {
                mode: systemMessage?.mode ?? 'replace',
                content: merged,
              };
            }
            // Deduplicate: only add extra agents that are not already in allCustomAgents
            // (they may be present if the browser already registered them from a previous
            // plugin.agents notification).
            const existingNames = new Set(allCustomAgents.map(a => a.name));
            const registeredExtraAgents = applySessionToolAccessToCustomAgents(
              extraAgents,
              sessionToolNames
            );
            allCustomAgents.push(
              ...registeredExtraAgents.filter(a => !existingNames.has(a.name))
            );
          }
        } catch (err) {
          console.warn('[proxy] Plugin agent notification failed:', err.message);
        }

        // Discover plugin prompts (slash commands) and notify the browser.
        try {
          if (pluginPrompts.length > 0) {
            console.log(`[proxy] Sending ${pluginPrompts.length} plugin prompt(s) to browser`);
            sendNotification('plugin.prompts', { prompts: pluginPrompts });
          }
        } catch (err) {
          console.warn('[proxy] Plugin prompt notification failed:', err.message);
        }

        const resolvedMcpServers = mergePluginMcpServers(mcpServers || {}, pluginMcpServers);

        // Emit 'starting' status for each configured MCP server
        const mcpServerNames = Object.keys(resolvedMcpServers || {});
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
            mcpServers: Object.keys(resolvedMcpServers).length > 0 ? resolvedMcpServers : undefined,
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
      if (!isTrustedRequestOrigin(request.headers.origin, request.socket?.remoteAddress)) {
        socket.write('HTTP/1.1 403 Forbidden\r\nConnection: close\r\n\r\n');
        socket.destroy();
        return;
      }

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
