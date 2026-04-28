import type {
  MCPHTTPServerConfig,
  MCPServerConfig,
  MCPStdioServerConfig,
} from '@github/copilot-sdk';
import type { McpServerConfig } from '@/types';

/**
 * Convert our internal McpServerConfig format to the SDK's MCPServerConfig record.
 * - HTTP/SSE servers become MCPHTTPServerConfig
 * - stdio servers become MCPStdioServerConfig (proxy spawns the subprocess)
 * All servers get `tools: ['*']` so the model can access every tool each server exports.
 */
export function toSdkMcpServers(configs: McpServerConfig[]): Record<string, MCPServerConfig> {
  const entries: [string, MCPServerConfig][] = configs.map(c => {
    if (c.transport === 'stdio') {
      const local: MCPStdioServerConfig = {
        type: 'stdio',
        command: c.command ?? '',
        args: c.args ?? [],
        ...(c.env !== undefined && { env: c.env }),
        tools: ['*'],
      };
      return [c.name, local];
    }
    const remote: MCPHTTPServerConfig = {
      type: c.transport,
      url: c.url ?? '',
      ...(c.headers !== undefined && { headers: c.headers }),
      tools: ['*'],
    };
    return [c.name, remote];
  });
  return Object.fromEntries(entries);
}
