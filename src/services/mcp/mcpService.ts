import type {
  MCPLocalServerConfig,
  MCPRemoteServerConfig,
  MCPServerConfig,
} from '@github/copilot-sdk';
import type { McpServerConfig } from '@/types';

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
