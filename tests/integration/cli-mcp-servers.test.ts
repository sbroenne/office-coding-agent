import { describe, expect, it } from 'vitest';
import { getCliMcpServers, parseCopilotMcpListJson } from '@/../src/plugins/cliMcpServers.mjs';

describe('CLI MCP servers', () => {
  it('parses Copilot CLI MCP list JSON into task pane server configs', () => {
    const servers = parseCopilotMcpListJson(
      JSON.stringify({
        mcpServers: {
          's360-breeze': {
            tools: ['*'],
            type: 'stdio',
            command: 'C:\\Users\\me\\agency.exe',
            args: ['mcp', 'remote', '--url', 'https://mcp.example.com'],
            source: 'user',
          },
          powerbi: {
            tools: ['*'],
            type: 'http',
            url: 'https://api.fabric.microsoft.com/v1/mcp/powerbi',
            source: 'plugin',
          },
        },
      })
    );

    expect(servers).toEqual([
      {
        name: 's360-breeze',
        description: 'Source: user',
        transport: 'stdio',
        command: 'C:\\Users\\me\\agency.exe',
        args: ['mcp', 'remote', '--url', 'https://mcp.example.com'],
        source: 'user',
      },
      {
        name: 'powerbi',
        description: 'Source: plugin',
        transport: 'http',
        url: 'https://api.fabric.microsoft.com/v1/mcp/powerbi',
        source: 'plugin',
      },
    ]);
  });

  it('returns an empty list and error when the CLI command fails', async () => {
    const result = await getCliMcpServers({
      runCommand: async () => ({
        success: false,
        stdout: '',
        stderr: 'nope',
        message: 'copilot failed',
      }),
    });

    expect(result).toEqual({
      servers: [],
      error: 'copilot failed',
    });
  });
});
