import { describe, expect, it } from 'vitest';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import {
  getCliMcpServers,
  getSelectedCliMcpServers,
  getMcpServerSummaries,
  parseCopilotMcpListJson,
  parseMcpServerDocument,
} from '@/../src/plugins/cliMcpServers.mjs';

describe('CLI MCP servers', () => {
  it('loads MCP execution settings from the local CLI by selected server name', async () => {
    const selected = await getSelectedCliMcpServers(['workiq'], {
      installedPluginsDir: await fs.mkdtemp(path.join(os.tmpdir(), 'oca-mcp-selection-')),
      runCommand: async () => ({
        success: true,
        stdout: JSON.stringify({
          mcpServers: {
            workiq: {
              type: 'stdio',
              command: 'npx',
              args: ['--token', 'local-secret'],
              env: { API_KEY: 'local-secret' },
            },
            unused: { type: 'stdio', command: 'node', args: ['unused.js'] },
          },
        }),
        stderr: '',
        message: '',
      }),
    });

    expect(selected.names).toEqual(['workiq']);
    expect(selected.servers).toEqual({
      workiq: {
        type: 'stdio',
        command: 'npx',
        args: ['--token', 'local-secret'],
        env: { API_KEY: 'local-secret' },
        tools: ['*'],
      },
    });
  });

  it('rejects names that are not present in the local CLI configuration', async () => {
    await expect(
      getSelectedCliMcpServers(['client-supplied'], {
        installedPluginsDir: await fs.mkdtemp(path.join(os.tmpdir(), 'oca-mcp-selection-')),
        runCommand: async () => ({
          success: true,
          stdout: JSON.stringify({ mcpServers: {} }),
          stderr: '',
          message: '',
        }),
      })
    ).rejects.toThrow('Selected MCP server configuration is unavailable: client-supplied.');
  });

  it('returns only safe display fields in MCP server summaries', () => {
    const summaries = getMcpServerSummaries([
      {
        name: 'private-local',
        transport: 'stdio',
        command: 'private-executable',
        args: ['--token', 'secret'],
        env: { API_KEY: 'secret' },
      },
      {
        name: 'private-remote',
        transport: 'http',
        url: 'https://user:password@mcp.example.com/private/path?token=secret',
        headers: { Authorization: 'Bearer secret' },
      },
    ]);

    expect(summaries).toEqual([
      {
        name: 'private-local',
        description: undefined,
        transport: 'stdio',
        oauthAlias: undefined,
        requiresOAuth: false,
        command: 'Local process',
      },
      {
        name: 'private-remote',
        description: undefined,
        transport: 'http',
        oauthAlias: undefined,
        requiresOAuth: false,
        url: 'https://mcp.example.com',
      },
    ]);
    expect(JSON.stringify(summaries)).not.toContain('secret');
    expect(JSON.stringify(summaries)).not.toContain('private-executable');
    expect(JSON.stringify(summaries)).not.toContain('private/path');
  });

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

  it('parses MCP documents that use VS Code-style servers without explicit transport types', () => {
    const servers = parseMcpServerDocument({
      servers: {
        'powerbi-remote': {
          url: 'https://api.fabric.microsoft.com/v1/mcp/powerbi',
        },
        workiq: {
          command: 'npx',
          args: ['-y', '@microsoft/workiq@latest', 'mcp'],
        },
      },
    });

    expect(servers).toEqual([
      {
        name: 'powerbi-remote',
        description: undefined,
        transport: 'http',
        url: 'https://api.fabric.microsoft.com/v1/mcp/powerbi',
        source: undefined,
      },
      {
        name: 'workiq',
        description: undefined,
        transport: 'stdio',
        command: 'npx',
        args: ['-y', '@microsoft/workiq@latest', 'mcp'],
        source: undefined,
      },
    ]);
  });

  it('merges plugin-declared MCP servers into the CLI-backed server list', async () => {
    const installedPluginsDir = await fs.mkdtemp(path.join(os.tmpdir(), 'oca-mcp-plugins-'));
    const pluginDir = path.join(installedPluginsDir, 'iq-core');
    await fs.mkdir(pluginDir, { recursive: true });
    await fs.writeFile(
      path.join(pluginDir, 'plugin.json'),
      JSON.stringify({
        name: 'iq-core',
        mcpServers: '.mcp.json',
      })
    );
    await fs.writeFile(
      path.join(pluginDir, '.mcp.json'),
      JSON.stringify({
        mcpServers: {
          'powerbi-remote': {
            type: 'http',
            url: 'https://api.fabric.microsoft.com/v1/mcp/powerbi',
          },
          workiq: {
            command: 'npx',
            args: ['-y', '@microsoft/workiq@latest', 'mcp'],
          },
        },
      })
    );

    const result = await getCliMcpServers({
      installedPluginsDir,
      runCommand: async () => ({
        success: true,
        stdout: JSON.stringify({
          mcpServers: {
            's360-breeze': {
              type: 'stdio',
              command: 'agency.exe',
              args: ['mcp', 'remote'],
              source: 'user',
            },
          },
        }),
        stderr: '',
        message: '',
      }),
    });

    expect(result).toEqual({
      error: undefined,
      servers: [
        {
          name: 's360-breeze',
          description: 'Source: user',
          transport: 'stdio',
          command: 'agency.exe',
          args: ['mcp', 'remote'],
          source: 'user',
        },
        {
          name: 'powerbi-remote',
          description: 'Source: plugin:iq-core',
          transport: 'http',
          url: 'https://api.fabric.microsoft.com/v1/mcp/powerbi',
          source: 'plugin:iq-core',
        },
        {
          name: 'workiq',
          description: 'Source: plugin:iq-core',
          transport: 'stdio',
          command: 'npx',
          args: ['-y', '@microsoft/workiq@latest', 'mcp'],
          source: 'plugin:iq-core',
        },
      ],
    });
  });
});
