import type { McpServerConfig, McpTransportType } from '@/types';

/** Shape accepted from both Claude Desktop and VS Code mcp.json formats */
interface RawMcpEntry {
  url?: string;
  type?: string;
  transport?: string;
  headers?: Record<string, string>;
  description?: string;
}

/**
 * Parse a `mcp.json` File into a normalized McpServerConfig array.
 *
 * Supports:
 *   - Claude Desktop format: `{ mcpServers: { name: { url, type, headers } } }`
 *   - VS Code format:        `{ servers:    { name: { url, type, headers } } }`
 *
 * Only HTTP/SSE entries are included (stdio entries are silently skipped because
 * they require a Node.js child process and do not work in a browser runtime).
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

    const url = entry.url;
    if (typeof url !== 'string' || !url) continue; // skip stdio entries (no url)

    const rawTransport = (entry.type ?? entry.transport ?? 'http').toLowerCase();
    if (rawTransport !== 'http' && rawTransport !== 'sse') continue; // skip unknown/stdio

    configs.push({
      name,
      description: typeof entry.description === 'string' ? entry.description : undefined,
      url,
      transport: rawTransport as McpTransportType,
      headers: entry.headers,
    });
  }

  if (configs.length === 0) {
    throw new Error(
      'No valid HTTP/SSE MCP servers found in mcp.json. ' +
        'Make sure each entry has a "url" and an optional "type" of "http" or "sse".'
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
