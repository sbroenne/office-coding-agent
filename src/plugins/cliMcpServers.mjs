import { spawn } from 'node:child_process';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';

const DEFAULT_TIMEOUT_MS = 30_000;
const DEFAULT_INSTALLED_PLUGINS_DIR = path.join(os.homedir(), '.copilot', 'installed-plugins');

export function runCopilotMcpListCommand(options = {}) {
  const timeoutMs = options.timeoutMs ?? DEFAULT_TIMEOUT_MS;
  const command = options.command ?? 'copilot';

  return new Promise(resolve => {
    const child = spawn(command, ['mcp', 'list', '--json'], {
      stdio: ['ignore', 'pipe', 'pipe'],
      windowsHide: true,
      env: process.env,
    });

    let stdout = '';
    let stderr = '';
    let settled = false;
    const timer = setTimeout(() => {
      settled = true;
      child.kill();
      resolve({
        success: false,
        stdout,
        stderr,
        message: `copilot mcp list --json timed out after ${timeoutMs}ms`,
      });
    }, timeoutMs);

    child.stdout?.on('data', chunk => {
      stdout += chunk.toString();
    });
    child.stderr?.on('data', chunk => {
      stderr += chunk.toString();
    });

    child.on('error', error => {
      if (settled) return;
      settled = true;
      clearTimeout(timer);
      resolve({
        success: false,
        stdout,
        stderr,
        message: error.message,
      });
    });

    child.on('close', code => {
      if (settled) return;
      settled = true;
      clearTimeout(timer);
      resolve({
        success: code === 0,
        stdout,
        stderr,
        message: (stdout || stderr).trim(),
        code,
      });
    });
  });
}

function normalizeTransport(type) {
  if (type === 'stdio' || type === 'local') return 'stdio';
  if (type === 'http' || type === 'sse') return type;
  return undefined;
}

function inferTransport(server) {
  const explicit = normalizeTransport(server.type ?? server.transport);
  if (explicit) return explicit;
  if (typeof server.command === 'string') return 'stdio';
  if (typeof server.url === 'string') return 'http';
  return undefined;
}

function parseMcpServerEntries(mcpServers, options = {}) {
  if (!mcpServers || typeof mcpServers !== 'object') return [];

  return Object.entries(mcpServers)
    .map(([name, server]) => {
      if (!server || typeof server !== 'object') return undefined;
      const transport = inferTransport(server);
      if (!transport) return undefined;
      const source = typeof server.source === 'string' ? server.source : options.source;

      const base = {
        name,
        description:
          typeof server.description === 'string'
            ? server.description
            : source
              ? `Source: ${source}`
              : undefined,
        transport,
        source,
      };

      if (transport === 'stdio') {
        return {
          ...base,
          command: typeof server.command === 'string' ? server.command : '',
          args: Array.isArray(server.args)
            ? server.args.filter(arg => typeof arg === 'string')
            : [],
          ...(server.env && typeof server.env === 'object' ? { env: server.env } : {}),
        };
      }

      return {
        ...base,
        url: typeof server.url === 'string' ? server.url : '',
        ...(server.headers && typeof server.headers === 'object'
          ? { headers: server.headers }
          : {}),
      };
    })
    .filter(Boolean);
}

export function parseCopilotMcpListJson(stdout, options = {}) {
  const parsed = JSON.parse(stdout);
  return parseMcpServerDocument(parsed, options);
}

export function parseMcpServerDocument(parsed, options = {}) {
  return parseMcpServerEntries(parsed?.mcpServers ?? parsed?.servers, options);
}

async function findPluginManifestPaths(root) {
  let entries;
  try {
    entries = await fs.readdir(root, { withFileTypes: true });
  } catch (error) {
    if (error?.code === 'ENOENT') return [];
    throw error;
  }

  const manifests = [];
  for (const entry of entries) {
    const entryPath = path.join(root, entry.name);
    if (entry.isFile() && entry.name === 'plugin.json') {
      manifests.push(entryPath);
    } else if (entry.isDirectory() && entry.name !== '.git' && entry.name !== 'node_modules') {
      manifests.push(...(await findPluginManifestPaths(entryPath)));
    }
  }
  return manifests;
}

async function readPluginMcpServers(manifestPath) {
  const manifest = JSON.parse(await fs.readFile(manifestPath, 'utf8'));
  const pluginDir = path.dirname(manifestPath);
  const pluginName = typeof manifest.name === 'string' ? manifest.name : path.basename(pluginDir);
  const source = `plugin:${pluginName}`;
  const mcpServers = manifest.mcpServers;

  if (typeof mcpServers === 'string') {
    const mcpPath = path.resolve(pluginDir, mcpServers);
    const document = JSON.parse(await fs.readFile(mcpPath, 'utf8'));
    return parseMcpServerDocument(document, { source });
  }

  if (mcpServers && typeof mcpServers === 'object') {
    return parseMcpServerEntries(mcpServers, { source });
  }

  return [];
}

export async function getPluginMcpServers(options = {}) {
  const installedPluginsDir = options.installedPluginsDir ?? DEFAULT_INSTALLED_PLUGINS_DIR;
  const manifestPaths = await findPluginManifestPaths(installedPluginsDir);
  const servers = [];
  const errors = [];

  for (const manifestPath of manifestPaths) {
    try {
      servers.push(...(await readPluginMcpServers(manifestPath)));
    } catch (error) {
      errors.push(`${manifestPath}: ${error instanceof Error ? error.message : String(error)}`);
    }
  }

  return { servers, errors };
}

function mergeServers(cliServers, pluginServers) {
  const seen = new Set();
  const merged = [];
  for (const server of [...cliServers, ...pluginServers]) {
    if (seen.has(server.name)) continue;
    seen.add(server.name);
    merged.push(server);
  }
  return merged;
}

export async function getCliMcpServers(options = {}) {
  const result = await (options.runCommand ?? runCopilotMcpListCommand)(options);
  if (!result.success) {
    return {
      servers: [],
      error: result.message || 'Failed to list Copilot CLI MCP servers',
    };
  }

  try {
    const cliServers = parseCopilotMcpListJson(result.stdout);
    const pluginResult = await getPluginMcpServers(options);
    return {
      servers: mergeServers(cliServers, pluginResult.servers),
      error:
        pluginResult.errors.length > 0
          ? `Failed to load plugin MCP servers: ${pluginResult.errors.join('; ')}`
          : undefined,
    };
  } catch (error) {
    return {
      servers: [],
      error: error instanceof Error ? error.message : String(error),
    };
  }
}
