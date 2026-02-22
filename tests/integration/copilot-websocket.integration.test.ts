// @vitest-environment node
/**
 * Live integration test: WebSocket → Copilot proxy → GitHub Copilot API
 *
 * Requires `npm run dev` to be running on https://localhost:3000.
 *
 * Run manually:
 *   npm run dev             # terminal 1
 *   npm test -- copilot-websocket   # terminal 2
 */

import { describe, it, expect } from 'vitest';
import WS from 'ws';
import type { SystemMessageConfig } from '@github/copilot-sdk';
import { createWebSocketClient } from '@/lib/websocket-client';

const SERVER_URL = 'wss://localhost:3000/api/copilot';
const TIMEOUT_MS = 30_000;

// Node doesn't trust the office-addin-dev-certs CA by default.
// Patch the global WebSocket to use `ws` with rejectUnauthorized: false
// so we can connect to the local HTTPS dev server.
global.WebSocket = class PatchedWebSocket extends WS {
  constructor(url: string | URL, protocols?: string | string[]) {
    super(url, typeof protocols === 'string' ? protocols : (protocols ?? []), {
      rejectUnauthorized: false,
    });
  }
} as unknown as typeof WebSocket;



const SYSTEM: SystemMessageConfig = {
  mode: 'append',
  content: 'You are a helpful assistant. Answer briefly.',
};

describe('Copilot WebSocket integration', () => {
  it(
    'connects to the proxy server',
    async () => {
      const client = await createWebSocketClient(SERVER_URL);
      expect(client).toBeTruthy();
      await client.stop();
    },
    TIMEOUT_MS
  );

  it(
    'creates a session and gets a response to a simple prompt',
    async () => {
      const client = await createWebSocketClient(SERVER_URL);
      try {
        const session = await client.createSession({ systemMessage: SYSTEM });
        expect(session.sessionId).toBeTruthy();

        const events: string[] = [];
        let fullText = '';

        for await (const event of session.query({ prompt: 'Reply with exactly: PONG' })) {
          events.push(event.type);
          if (event.type === 'assistant.message_delta') {
            fullText += event.data.deltaContent;
          }
          if (event.type === 'assistant.message') {
            // Some models emit a complete message instead of streaming deltas
            fullText += event.data.content;
          }
          if (event.type === 'session.idle') break;
        }

        expect(events).toContain('session.idle');
        expect(fullText.length).toBeGreaterThan(0);
        expect(fullText.toLowerCase()).toContain('pong');
      } finally {
        await client.stop();
      }
    },
    TIMEOUT_MS
  );

  it(
    'lists available models from the Copilot API',
    async () => {
      const client = await createWebSocketClient(SERVER_URL);
      try {
        // listModels requires an active session (CLI must be initialized)
        await client.createSession({ systemMessage: SYSTEM });
        const models = await client.listModels();
        console.log('>>> listModels returned', models.length, 'models:');
        for (const m of models) {
          console.log(`  - ${m.id} (${m.name})`);
        }

        expect(Array.isArray(models)).toBe(true);
        expect(models.length).toBeGreaterThan(0);
        expect(models[0]).toHaveProperty('id');
        expect(models[0]).toHaveProperty('name');
      } finally {
        await client.stop();
      }
    },
    TIMEOUT_MS
  );

  it(
    'executes a tool call and returns the result to the model',
    async () => {
      const client = await createWebSocketClient(SERVER_URL);
      try {
        let toolWasCalled = false;

        const session = await client.createSession({
          systemMessage: {
            mode: 'append',
            content:
              'You must call the echo tool with the text "hello". Do not answer without calling it.',
          },
          tools: [
            {
              name: 'echo',
              description: 'Echoes a message back',
              parameters: {
                type: 'object' as const,
                properties: {
                  text: { type: 'string', description: 'Text to echo' },
                },
                required: ['text'],
              },
              handler: (args: unknown) => {
                toolWasCalled = true;
                const { text } = args as { text: string };
                return Promise.resolve({
                  textResultForLlm: `Echo: ${text}`,
                  resultType: 'success' as const,
                  toolTelemetry: {},
                });
              },
            },
          ],
        });

        for await (const event of session.query({
          prompt: 'Please call the echo tool with "hello".',
        })) {
          if (event.type === 'session.idle') break;
        }

        expect(toolWasCalled).toBe(true);
      } finally {
        await client.stop();
      }
    },
    TIMEOUT_MS
  );
});
