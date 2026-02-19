import { describe, it, expect } from 'vitest';
import { parseMcpJsonFile, resolveActiveMcpServers } from '@/services/mcp';
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
    expect(result[0]).toMatchObject({ name: 'my-server', url: 'https://example.com/mcp', transport: 'http' });
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

  it('skips entries without a url (stdio entries)', async () => {
    const file = makeFile({
      mcpServers: {
        stdio: { command: 'node', args: ['server.js'] },
        web: { url: 'https://example.com/mcp', type: 'http' },
      },
    });
    const result = await parseMcpJsonFile(file);
    expect(result).toHaveLength(1);
    expect(result[0].name).toBe('web');
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
    await expect(parseMcpJsonFile(file)).rejects.toThrow();
  });

  it('throws when all entries are skipped (no valid HTTP/SSE servers)', async () => {
    const file = makeFile({
      mcpServers: {
        stdio: { command: 'node' },
      },
    });
    await expect(parseMcpJsonFile(file)).rejects.toThrow('No valid HTTP/SSE');
  });

  it('throws on non-.json file extension', async () => {
    const file = new File(['{}'], 'mcp.txt');
    await expect(parseMcpJsonFile(file)).rejects.toThrow('.json');
  });

  it.each([
    ['mcpServers', { mcpServers: { a: { url: 'https://a.com', type: 'http' }, b: { url: 'https://b.com', type: 'sse' } } }],
    ['servers', { servers: { a: { url: 'https://a.com', type: 'http' }, b: { url: 'https://b.com', type: 'sse' } } }],
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
