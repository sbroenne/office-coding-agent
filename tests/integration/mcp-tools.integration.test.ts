/**
 * Integration test: MCP client ↔ local test server.
 *
 * Starts a real Streamable HTTP MCP server in-process, then uses
 * @ai-sdk/mcp createMCPClient to connect, discover tools, and execute them.
 *
 * This tests the full round-trip: useMcpTools hook logic (createMCPClient +
 * .tools() + .execute()) against a real server — no public API needed.
 */

import { describe, it, expect, beforeAll, afterAll, afterEach } from 'vitest';
import { createMCPClient } from '@ai-sdk/mcp';
import { startMcpTestServer, type McpTestServer } from '../helpers/mcp-test-server';

let server: McpTestServer;
let client: Awaited<ReturnType<typeof createMCPClient>> | undefined;

beforeAll(async () => {
  server = await startMcpTestServer();
});

afterAll(async () => {
  await server.stop();
});

afterEach(async () => {
  if (client) {
    await client.close();
    client = undefined;
  }
});

describe('MCP integration: local test server', () => {
  it('connects and retrieves the tool list', async () => {
    client = await createMCPClient({
      transport: { type: 'http', url: server.url },
    });

    const tools = await client.tools();
    expect(Object.keys(tools)).toContain('echo');
    expect(Object.keys(tools)).toContain('add');
  });

  it('each tool has a description and execute function', async () => {
    client = await createMCPClient({
      transport: { type: 'http', url: server.url },
    });

    const tools = await client.tools();
    for (const tool of Object.values(tools)) {
      expect(tool).toHaveProperty('description');
      expect(tool).toHaveProperty('execute');
      expect(typeof tool.execute).toBe('function');
    }
  });

  it('echo tool returns the input message', async () => {
    client = await createMCPClient({
      transport: { type: 'http', url: server.url },
    });

    const tools = await client.tools();
    const result = await tools.echo.execute({ message: 'hello mcp' }, { messages: [], toolCallId: 'test-1' });
    expect(result).toMatchObject({ content: [{ type: 'text', text: 'hello mcp' }] });
  });

  it('add tool returns the correct sum', async () => {
    client = await createMCPClient({
      transport: { type: 'http', url: server.url },
    });

    const tools = await client.tools();
    const result = await tools.add.execute({ a: 7, b: 35 }, { messages: [], toolCallId: 'test-2' });
    expect(result).toMatchObject({ content: [{ type: 'text', text: '42' }] });
  });

  it('multiple clients can connect concurrently', async () => {
    const [c1, c2] = await Promise.all([
      createMCPClient({ transport: { type: 'http', url: server.url } }),
      createMCPClient({ transport: { type: 'http', url: server.url } }),
    ]);

    try {
      const [t1, t2] = await Promise.all([c1.tools(), c2.tools()]);
      expect(Object.keys(t1)).toContain('echo');
      expect(Object.keys(t2)).toContain('add');
    } finally {
      await Promise.all([c1.close(), c2.close()]);
    }
  });
});
