/**
 * Local HTTP MCP test server.
 *
 * Starts a minimal Streamable HTTP MCP server on a random port for use
 * in integration tests. The server exposes two deterministic test tools:
 *   - echo(message)  → returns the message as text
 *   - add(a, b)      → returns the sum of two numbers
 *
 * Usage:
 *   const server = await startMcpTestServer();
 *   // server.url → 'http://127.0.0.1:<port>/mcp'
 *   await server.stop();
 */

import * as http from 'http';
import type { AddressInfo } from 'net';
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { StreamableHTTPServerTransport } from '@modelcontextprotocol/sdk/server/streamableHttp.js';
import { z } from 'zod';

export interface McpTestServer {
  /** Full URL to POST MCP requests to, e.g. http://127.0.0.1:12345/mcp */
  url: string;
  stop: () => Promise<void>;
}

function buildMcpServer(): McpServer {
  const server = new McpServer({ name: 'test-server', version: '1.0.0' });

  server.registerTool(
    'echo',
    { description: 'Returns the input message unchanged', inputSchema: { message: z.string() } },
    async ({ message }) => ({ content: [{ type: 'text' as const, text: message }] })
  );

  server.registerTool(
    'add',
    { description: 'Adds two numbers and returns the result', inputSchema: { a: z.number(), b: z.number() } },
    async ({ a, b }) => ({ content: [{ type: 'text' as const, text: String(a + b) }] })
  );

  return server;
}

async function readBody(req: http.IncomingMessage): Promise<unknown> {
  const chunks: Buffer[] = [];
  for await (const chunk of req) {
    chunks.push(chunk as Buffer);
  }
  return JSON.parse(Buffer.concat(chunks).toString()) as unknown;
}

export async function startMcpTestServer(): Promise<McpTestServer> {
  const httpServer = http.createServer(async (req, res) => {
    if (req.method !== 'POST') {
      res.writeHead(405).end(
        JSON.stringify({ jsonrpc: '2.0', error: { code: -32000, message: 'Method not allowed' }, id: null })
      );
      return;
    }

    try {
      const body = await readBody(req);
      const mcpServer = buildMcpServer();
      const transport = new StreamableHTTPServerTransport({ sessionIdGenerator: undefined });
      await mcpServer.connect(transport);
      await transport.handleRequest(req, res, body);
      res.on('close', () => {
        void transport.close();
        void mcpServer.close();
      });
    } catch (err) {
      if (!res.headersSent) {
        res.writeHead(500).end(
          JSON.stringify({ jsonrpc: '2.0', error: { code: -32603, message: 'Internal error' }, id: null })
        );
      }
    }
  });

  await new Promise<void>(resolve => {
    httpServer.listen(0, '127.0.0.1', resolve);
  });

  const { port } = httpServer.address() as AddressInfo;

  return {
    url: `http://127.0.0.1:${port}/mcp`,
    stop: () =>
      new Promise<void>((resolve, reject) => {
        httpServer.close(err => (err ? reject(err) : resolve()));
      }),
  };
}
