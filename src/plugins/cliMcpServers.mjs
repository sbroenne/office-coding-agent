import { spawn } from 'node:child_process';

const DEFAULT_TIMEOUT_MS = 30_000;

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

export function parseCopilotMcpListJson(stdout) {
  const parsed = JSON.parse(stdout);
  const mcpServers = parsed?.mcpServers;
  if (!mcpServers || typeof mcpServers !== 'object') return [];

  return Object.entries(mcpServers)
    .map(([name, server]) => {
      if (!server || typeof server !== 'object') return undefined;
      const transport = normalizeTransport(server.type);
      if (!transport) return undefined;

      const base = {
        name,
        description: server.source ? `Source: ${server.source}` : undefined,
        transport,
        source: server.source,
      };

      if (transport === 'stdio') {
        return {
          ...base,
          command: typeof server.command === 'string' ? server.command : '',
          args: Array.isArray(server.args) ? server.args.filter(arg => typeof arg === 'string') : [],
          ...(server.env && typeof server.env === 'object' ? { env: server.env } : {}),
        };
      }

      return {
        ...base,
        url: typeof server.url === 'string' ? server.url : '',
        ...(server.headers && typeof server.headers === 'object' ? { headers: server.headers } : {}),
      };
    })
    .filter(Boolean);
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
    return { servers: parseCopilotMcpListJson(result.stdout), error: undefined };
  } catch (error) {
    return {
      servers: [],
      error: error instanceof Error ? error.message : String(error),
    };
  }
}
