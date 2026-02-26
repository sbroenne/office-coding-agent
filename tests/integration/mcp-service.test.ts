import { describe, it, expect } from 'vitest';
import { parseMcpJsonFile, resolveActiveMcpServers, toSdkMcpServers } from '@/services/mcp';
import type { McpServerConfig } from '@/types';

function makeFile(content: unknown, name = 'mcp.json'): File {
  return new File([JSON.stringify(content)], name, { type: 'application/json' });
}

describe('parseMcpJsonFile', () => {
  it('parses Claude Desktop mcpServers format', async () => {
    const file = makeFile({
      mcpServers: {
        'my-server': { url: 'https://example.com/mcp', type: 'http' },
      },
    });
    const result = await parseMcpJsonFile(file);
    expect(result).toHaveLength(1);
    expect(result[0]).toMatchObject({
      name: 'my-server',
      url: 'https://example.com/mcp',
      transport: 'http',
    });
  });

  it('parses VS Code servers format', async () => {
    const file = makeFile({
      servers: {
        'sse-server': { url: 'https://example.com/sse', type: 'sse' },
      },
    });
    const result = await parseMcpJsonFile(file);
    expect(result).toHaveLength(1);
    expect(result[0]).toMatchObject({ name: 'sse-server', transport: 'sse' });
  });

  it('defaults transport to http when type is omitted', async () => {
    const file = makeFile({ mcpServers: { srv: { url: 'https://example.com/mcp' } } });
    const result = await parseMcpJsonFile(file);
    expect(result[0].transport).toBe('http');
  });

  it('preserves headers and description', async () => {
    const file = makeFile({
      mcpServers: {
        srv: {
          url: 'https://example.com/mcp',
          type: 'http',
          headers: { Authorization: 'Bearer tok' },
          description: 'My server',
        },
      },
    });
    const result = await parseMcpJsonFile(file);
    expect(result[0].headers).toEqual({ Authorization: 'Bearer tok' });
    expect(result[0].description).toBe('My server');
  });

  it('includes stdio entries alongside http/sse entries', async () => {
    const file = makeFile({
      mcpServers: {
        stdio: { command: 'node', args: ['server.js'] },
        web: { url: 'https://example.com/mcp', type: 'http' },
      },
    });
    const result = await parseMcpJsonFile(file);
    expect(result).toHaveLength(2);
    expect(result.find(s => s.name === 'stdio')).toMatchObject({
      transport: 'stdio',
      command: 'node',
      args: ['server.js'],
    });
    expect(result.find(s => s.name === 'web')).toMatchObject({ transport: 'http' });
  });

  it('parses a VS Code-style npx stdio entry', async () => {
    const file = makeFile({
      servers: {
        'mcp-docs-server': { command: 'npx', args: ['-y', '@assistant-ui/mcp-docs-server'] },
      },
    });
    const result = await parseMcpJsonFile(file);
    expect(result).toHaveLength(1);
    expect(result[0]).toMatchObject({
      name: 'mcp-docs-server',
      transport: 'stdio',
      command: 'npx',
      args: ['-y', '@assistant-ui/mcp-docs-server'],
    });
  });

  it('passes env to stdio entries when present', async () => {
    const file = makeFile({
      mcpServers: {
        local: { command: 'node', args: ['srv.js'], env: { API_KEY: 'abc' } },
      },
    });
    const result = await parseMcpJsonFile(file);
    expect(result[0]).toMatchObject({ transport: 'stdio', env: { API_KEY: 'abc' } });
  });

  it('skips unknown transport types', async () => {
    const file = makeFile({
      mcpServers: {
        bad: { url: 'https://example.com/mcp', type: 'grpc' },
        good: { url: 'https://example.com/mcp', type: 'sse' },
      },
    });
    const result = await parseMcpJsonFile(file);
    expect(result).toHaveLength(1);
    expect(result[0].name).toBe('good');
  });

  it('throws on invalid JSON', async () => {
    const file = new File(['not-json'], 'mcp.json');
    await expect(parseMcpJsonFile(file)).rejects.toThrow('Invalid JSON');
  });

  it('throws when no mcpServers/servers key exists', async () => {
    const file = makeFile({ foo: 'bar' });
    await expect(parseMcpJsonFile(file)).rejects.toThrow(/mcpServers|servers/);
  });

  it('throws when all entries are skipped (neither url nor command)', async () => {
    const file = makeFile({
      mcpServers: {
        empty: { description: 'no url or command' },
      },
    });
    await expect(parseMcpJsonFile(file)).rejects.toThrow('No valid MCP servers');
  });

  it('throws on non-.json file extension', async () => {
    const file = new File(['{}'], 'mcp.txt');
    await expect(parseMcpJsonFile(file)).rejects.toThrow('.json');
  });

  it.each([
    [
      'mcpServers',
      {
        mcpServers: {
          a: { url: 'https://a.com', type: 'http' },
          b: { url: 'https://b.com', type: 'sse' },
        },
      },
    ],
    [
      'servers',
      {
        servers: {
          a: { url: 'https://a.com', type: 'http' },
          b: { url: 'https://b.com', type: 'sse' },
        },
      },
    ],
  ])('handles multiple servers in %s format', async (_fmt, content) => {
    const file = makeFile(content);
    const result = await parseMcpJsonFile(file);
    expect(result).toHaveLength(2);
  });
});

describe('resolveActiveMcpServers', () => {
  const servers: McpServerConfig[] = [
    { name: 'a', url: 'https://a.com', transport: 'http' },
    { name: 'b', url: 'https://b.com', transport: 'sse' },
    { name: 'c', url: 'https://c.com', transport: 'http' },
  ];

  it('returns all servers when activeNames is null', () => {
    expect(resolveActiveMcpServers(servers, null)).toHaveLength(3);
  });

  it('filters to only active server names', () => {
    const result = resolveActiveMcpServers(servers, ['a', 'c']);
    expect(result.map(s => s.name)).toEqual(['a', 'c']);
  });

  it('returns empty when activeNames is empty array', () => {
    expect(resolveActiveMcpServers(servers, [])).toHaveLength(0);
  });
});

describe('toSdkMcpServers', () => {
  it('converts an HTTP server correctly', () => {
    const result = toSdkMcpServers([{ name: 'my-server', url: 'https://example.com/mcp', transport: 'http' }]);
    expect(result).toHaveProperty('my-server');
    expect(result['my-server']).toMatchObject({ type: 'http', url: 'https://example.com/mcp', tools: ['*'] });
  });

  it('converts an SSE server correctly', () => {
    const result = toSdkMcpServers([{ name: 'sse-srv', url: 'https://sse.example.com', transport: 'sse' }]);
    expect(result['sse-srv'].type).toBe('sse');
    expect(result['sse-srv'].tools).toEqual(['*']);
  });

  it('includes headers when present', () => {
    const result = toSdkMcpServers([{
      name: 'auth-server',
      url: 'https://example.com/mcp',
      transport: 'http',
      headers: { Authorization: 'Bearer tok' },
    }]);
    expect(result['auth-server']).toMatchObject({ headers: { Authorization: 'Bearer tok' } });
  });

  it('omits headers key when headers is undefined', () => {
    const result = toSdkMcpServers([{ name: 'bare', url: 'https://example.com', transport: 'http' }]);
    expect(Object.keys(result['bare'])).not.toContain('headers');
  });

  it('converts a stdio server to MCPLocalServerConfig', () => {
    const result = toSdkMcpServers([{
      name: 'local-srv',
      transport: 'stdio',
      command: 'npx',
      args: ['-y', '@some/mcp-server'],
    }]);
    expect(result['local-srv']).toMatchObject({
      type: 'stdio',
      command: 'npx',
      args: ['-y', '@some/mcp-server'],
      tools: ['*'],
    });
  });

  it('passes env to stdio server when present', () => {
    const result = toSdkMcpServers([{
      name: 'env-srv',
      transport: 'stdio',
      command: 'node',
      args: ['srv.js'],
      env: { TOKEN: 'secret' },
    }]);
    expect(result['env-srv']).toMatchObject({ type: 'stdio', env: { TOKEN: 'secret' } });
  });

  it('omits env key when env is undefined for stdio server', () => {
    const result = toSdkMcpServers([{ name: 'bare-stdio', transport: 'stdio', command: 'node', args: [] }]);
    expect(Object.keys(result['bare-stdio'])).not.toContain('env');
  });

  it('returns a record with one key per server', () => {
    const result = toSdkMcpServers([
      { name: 'a', url: 'https://a.com', transport: 'http' },
      { name: 'b', url: 'https://b.com', transport: 'sse' },
    ]);
    expect(Object.keys(result)).toEqual(['a', 'b']);
  });

  it('returns empty object for empty input', () => {
    expect(toSdkMcpServers([])).toEqual({});
  });
});
