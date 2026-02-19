/**
 * Live integration test: WebSocket → Copilot proxy → GitHub Copilot API
 *
 * Requires `npm run server` to be running on https://localhost:3000.
 * Skips automatically when the server is unreachable.
 *
 * Run manually:
 *   npm run server          # terminal 1
 *   npm test -- copilot-websocket   # terminal 2
 */

import { describe, it, expect, beforeAll } from 'vitest';
import type { SystemMessageConfig } from '@github/copilot-sdk';
import { createWebSocketClient } from '@/lib/websocket-client';

const SERVER_URL = 'ws://localhost:3000/api/copilot';
const TIMEOUT_MS = 30_000;

let serverAvailable = false;

beforeAll(async () => {
  try {
    await new Promise<void>((resolve, reject) => {
      const ws = new WebSocket(SERVER_URL);
      const t = setTimeout(() => {
        ws.close();
        reject(new Error('timeout'));
      }, 3000);
      ws.addEventListener('open', () => {
        clearTimeout(t);
        ws.close();
        resolve();
      });
      ws.addEventListener('error', () => {
        clearTimeout(t);
        reject(new Error('connection refused'));
      });
    });
    serverAvailable = true;
  } catch {
    serverAvailable = false;
  }
});

const SYSTEM: SystemMessageConfig = {
  mode: 'append',
  content: 'You are a helpful assistant. Answer briefly.',
};

describe('Copilot WebSocket integration', () => {
  it(
    'connects to the proxy server',
    async () => {
      if (!serverAvailable) {
        console.log('Skipping — start `npm run server` to run live Copilot tests');
        return;
      }
      const client = await createWebSocketClient(SERVER_URL);
      expect(client).toBeTruthy();
      await client.stop();
    },
    TIMEOUT_MS
  );

  it(
    'creates a session and gets a response to a simple prompt',
    async () => {
      if (!serverAvailable) {
        console.log('Skipping — start `npm run server` to run live Copilot tests');
        return;
      }

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
    'executes a tool call and returns the result to the model',
    async () => {
      if (!serverAvailable) {
        console.log('Skipping — start `npm run server` to run live Copilot tests');
        return;
      }

      const client = await createWebSocketClient(SERVER_URL);
      try {
        let toolWasCalled = false;

        const session = await client.createSession({
          systemMessage: {
            mode: 'append',
            content: 'You must call the echo tool with the text "hello". Do not answer without calling it.',
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
              handler: async (args: unknown) => {
                toolWasCalled = true;
                const { text } = args as { text: string };
                return {
                  textResultForLlm: `Echo: ${text}`,
                  resultType: 'success' as const,
                  toolTelemetry: {},
                };
              },
            },
          ],
        });

        for await (const event of session.query({ prompt: 'Please call the echo tool with "hello".' })) {
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
