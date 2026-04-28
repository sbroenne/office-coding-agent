import { getLocalApiBase } from '@/lib/api';
import type { McpServerConfig } from '@/types';

export async function fetchConfiguredMcpServers(): Promise<McpServerConfig[]> {
  try {
    const res = await fetch(`${getLocalApiBase()}/api/mcp-servers`);
    if (!res.ok) {
      console.warn(`[mcp] Failed to load CLI MCP servers: HTTP ${res.status}`);
      return [];
    }
    const data = (await res.json()) as { servers?: McpServerConfig[]; error?: string };
    if (data.error) {
      console.warn(`[mcp] Failed to load CLI MCP servers: ${data.error}`);
    }
    return data.servers ?? [];
  } catch (error) {
    console.warn(
      `[mcp] Failed to load CLI MCP servers: ${
        error instanceof Error ? error.message : String(error)
      }`
    );
    return [];
  }
}
