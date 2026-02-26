import type {
  MCPLocalServerConfig,
  MCPRemoteServerConfig,
  MCPServerConfig,
} from '@github/copilot-sdk';
import type { McpServerConfig, McpTransportType } from '@/types';

/** Shape accepted from both Claude Desktop and VS Code mcp.json formats */
interface RawMcpEntry {
  url?: string;
  type?: string;
  transport?: string;
  headers?: Record<string, string>;
  description?: string;
  // stdio fields
  command?: string;
  args?: string[];
  env?: Record<string, string>;
}

/**
 * Parse a `mcp.json` File into a normalized McpServerConfig array.
 *
 * Supports:
 *   - Claude Desktop format: `{ mcpServers: { name: { url, type, headers } } }`
 *   - VS Code format:        `{ servers:    { name: { url, type, headers } } }`
 *
 * Includes both HTTP/SSE remote entries and stdio entries (e.g. npx-based MCP servers).
 * Stdio servers require a local proxy to spawn the subprocess — they are forwarded to
 * the Copilot SDK's MCPLocalServerConfig.
 */
export async function parseMcpJsonFile(file: File): Promise<McpServerConfig[]> {
  if (!file.name.endsWith('.json')) {
    throw new Error('File must be a .json file.');
  }

  const text = await file.text();
  let parsed: unknown;
  try {
    parsed = JSON.parse(text) as unknown;
  } catch {
    throw new Error('Invalid JSON: could not parse mcp.json.');
  }

  if (typeof parsed !== 'object' || parsed === null || Array.isArray(parsed)) {
    throw new Error('mcp.json must be a JSON object.');
  }

  // Support both { mcpServers: {...} } and { servers: {...} }
  const raw = parsed as Record<string, unknown>;
  const serversMap =
    (raw.mcpServers as Record<string, RawMcpEntry> | undefined) ??
    (raw.servers as Record<string, RawMcpEntry> | undefined);

  if (!serversMap || typeof serversMap !== 'object' || Array.isArray(serversMap)) {
    throw new Error('mcp.json must contain a "mcpServers" or "servers" object.');
  }

  const configs: McpServerConfig[] = [];

  for (const [name, entry] of Object.entries(serversMap)) {
    if (typeof entry !== 'object' || entry === null) continue;

    const description = typeof entry.description === 'string' ? entry.description : undefined;

    // Stdio entry: has command but no url
    if (typeof entry.command === 'string' && entry.command) {
      configs.push({
        name,
        description,
        transport: 'stdio',
        command: entry.command,
        args: Array.isArray(entry.args) ? entry.args : [],
        env: entry.env,
      });
      continue;
    }

    const url = entry.url;
    if (typeof url !== 'string' || !url) continue; // skip entries with no url and no command

    const rawTransport = (entry.type ?? entry.transport ?? 'http').toLowerCase();
    if (rawTransport !== 'http' && rawTransport !== 'sse') continue; // skip unknown transport

    configs.push({
      name,
      description,
      url,
      transport: rawTransport as McpTransportType,
      headers: entry.headers,
    });
  }

  if (configs.length === 0) {
    throw new Error(
      'No valid MCP servers found in mcp.json. ' +
        'Each entry must have either a "url" (for http/sse) or a "command" (for stdio/npx).'
    );
  }

  return configs;
}

/** Return only the servers that are currently active (null = all active). */
export function resolveActiveMcpServers(
  servers: McpServerConfig[],
  activeNames: string[] | null
): McpServerConfig[] {
  if (activeNames === null) return servers;
  return servers.filter(s => activeNames.includes(s.name));
}

/**
 * Convert our internal McpServerConfig format to the SDK's MCPServerConfig record.
 * - HTTP/SSE servers become MCPRemoteServerConfig
 * - stdio servers become MCPLocalServerConfig (proxy spawns the subprocess)
 * All servers get `tools: ['*']` so the model can access every tool each server exports.
 */
export function toSdkMcpServers(configs: McpServerConfig[]): Record<string, MCPServerConfig> {
  const entries: [string, MCPServerConfig][] = configs.map(c => {
    if (c.transport === 'stdio') {
      const local: MCPLocalServerConfig = {
        type: 'stdio',
        command: c.command ?? '',
        args: c.args ?? [],
        ...(c.env !== undefined && { env: c.env }),
        tools: ['*'],
      };
      return [c.name, local];
    }
    const remote: MCPRemoteServerConfig = {
      type: c.transport,
      url: c.url ?? '',
      ...(c.headers !== undefined && { headers: c.headers }),
      tools: ['*'],
    };
    return [c.name, remote];
  });
  return Object.fromEntries(entries);
}
