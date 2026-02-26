// @vitest-environment node
/**
 * Live integration test: custom agent instructions via WebSocket → Copilot API.
 *
 * Requires `npm run dev` to be running on https://localhost:3000.
 *
 * These tests verify that custom agent instructions actually change the model's
 * behaviour — i.e. the system prompt wiring works end-to-end through the proxy.
 */

import { describe, it, expect } from 'vitest';
import WS from 'ws';
import type { SystemMessageConfig } from '@github/copilot-sdk';
import { createWebSocketClient } from '@/lib/websocket-client';
import { buildSystemPrompt } from '@/services/ai/systemPrompt';
import { buildSkillContext } from '@/services/skills';

const SERVER_URL = 'wss://localhost:3000/api/copilot';
const TIMEOUT_MS = 45_000;
/** Extra time for tests that run npm install inside the proxy before the session opens. */
const TIMEOUT_SKILLPM_MS = 90_000;

global.WebSocket = class PatchedWebSocket extends WS {
  constructor(url: string | URL, protocols?: string | string[]) {
    super(url, typeof protocols === 'string' ? protocols : (protocols ?? []), {
      rejectUnauthorized: false,
    });
  }
} as unknown as typeof WebSocket;



describe('Copilot custom agent integration', () => {
  it(
    'custom agent instructions influence the model response',
    async () => {
      const customInstructions =
        'You are the Pirate Agent. You MUST respond to every message entirely in pirate speak. ' +
        'Use words like "ahoy", "matey", "arrr", "ye", "treasure", "sail". ' +
        'Never respond in normal English.';

      const systemContent = `${buildSystemPrompt('excel')}\n\n${customInstructions}`;
      const systemMessage: SystemMessageConfig = {
        mode: 'replace',
        content: systemContent,
      };

      const client = await createWebSocketClient(SERVER_URL);
      try {
        const session = await client.createSession({ systemMessage });

        let fullText = '';
        for await (const event of session.query({ prompt: 'Say hello to me' })) {
          if (event.type === 'assistant.message_delta') {
            fullText += event.data.deltaContent;
          }
          if (event.type === 'assistant.message') {
            fullText = event.data.content;
          }
          if (event.type === 'session.idle') break;
        }

        const lower = fullText.toLowerCase();
        const pirateTerms = ['ahoy', 'matey', 'arrr', 'ye', 'pirate', 'sail', 'treasure', 'aye'];
        const hasPirateSpeak = pirateTerms.some(term => lower.includes(term));
        expect(hasPirateSpeak).toBe(true);
      } finally {
        await client.stop();
      }
    },
    TIMEOUT_MS
  );

  it(
    'skill context injected into system prompt influences the response',
    async () => {
      // Build a system prompt with a synthetic skill that provides a known fact
      const syntheticSkillContext =
        '\n\n# Agent Skills\n' +
        'The following agent skills provide domain-specific knowledge.\n\n' +
        '---\n## Agent Skill: Secret Code\n' +
        'This skill provides a secret code the user may ask about.\n\n' +
        'The secret code is: AZURE-FALCON-42. ' +
        'When the user asks for the secret code, reply with exactly this code.';

      const systemContent = `${buildSystemPrompt('excel')}${syntheticSkillContext}`;
      const systemMessage: SystemMessageConfig = {
        mode: 'replace',
        content: systemContent,
      };

      const client = await createWebSocketClient(SERVER_URL);
      try {
        const session = await client.createSession({ systemMessage });

        let fullText = '';
        for await (const event of session.query({ prompt: 'What is the secret code?' })) {
          if (event.type === 'assistant.message_delta') {
            fullText += event.data.deltaContent;
          }
          if (event.type === 'assistant.message') {
            fullText = event.data.content;
          }
          if (event.type === 'session.idle') break;
        }

        expect(fullText).toContain('AZURE-FALCON-42');
      } finally {
        await client.stop();
      }
    },
    TIMEOUT_MS
  );

  it(
    'buildSkillContext with active skill names filters correctly',
    async () => {
      // Verify that buildSkillContext with empty array produces no skill injection
      const noSkills = buildSkillContext([]);
      expect(noSkills).toBe('');

      // With undefined (all skills), should include bundled Excel skill
      const allSkills = buildSkillContext();
      expect(allSkills).toContain('Agent Skill');

      // Use the "all skills" system prompt and ask something Excel-specific
      const systemContent = `${buildSystemPrompt('excel')}${allSkills}`;
      const systemMessage: SystemMessageConfig = {
        mode: 'replace',
        content: systemContent,
      };

      const client = await createWebSocketClient(SERVER_URL);
      try {
        const session = await client.createSession({ systemMessage });

        let fullText = '';
        for await (const event of session.query({
          prompt: 'Reply with exactly one word: PONG',
        })) {
          if (event.type === 'assistant.message_delta') {
            fullText += event.data.deltaContent;
          }
          if (event.type === 'assistant.message') {
            fullText = event.data.content;
          }
          if (event.type === 'session.idle') break;
        }

        expect(fullText.toLowerCase()).toContain('pong');
      } finally {
        await client.stop();
      }
    },
    TIMEOUT_MS
  );

  it(
    'tool scoping via availableTools restricts which tools are offered',
    async () => {
      // Register two tools but restrict availableTools to only one
      const echoWasCalled = { value: false };
      const blockedWasCalled = { value: false };

      const tools = [
        {
          name: 'echo_tool',
          description: 'Echoes text back',
          parameters: {
            type: 'object' as const,
            properties: { text: { type: 'string', description: 'Text to echo' } },
            required: ['text'],
          },
          handler: (args: unknown) => {
            echoWasCalled.value = true;
            return Promise.resolve({
              textResultForLlm: `Echo: ${(args as { text: string }).text}`,
              resultType: 'success' as const,
              toolTelemetry: {},
            });
          },
        },
        {
          name: 'blocked_tool',
          description: 'This tool should not be called',
          parameters: {
            type: 'object' as const,
            properties: { data: { type: 'string', description: 'Data' } },
            required: ['data'],
          },
          handler: () => {
            blockedWasCalled.value = true;
            return Promise.resolve({
              textResultForLlm: 'Should not be called',
              resultType: 'success' as const,
              toolTelemetry: {},
            });
          },
        },
      ];

      const client = await createWebSocketClient(SERVER_URL);
      try {
        const session = await client.createSession({
          systemMessage: {
            mode: 'replace',
            content:
              'You must call the echo_tool with the text "test". Do not answer without calling it. Do not call blocked_tool.',
          },
          tools,
          availableTools: ['echo_tool'],
        });

        for await (const event of session.query({
          prompt: 'Call the echo tool with "test".',
        })) {
          if (event.type === 'session.idle') break;
        }

        expect(echoWasCalled.value).toBe(true);
        expect(blockedWasCalled.value).toBe(false);
      } finally {
        await client.stop();
      }
    },
    TIMEOUT_MS
  );

  it(
    'report_intent events are emitted during tool-calling turns',
    async () => {
      const client = await createWebSocketClient(SERVER_URL);
      try {
        const session = await client.createSession({
          systemMessage: {
            mode: 'replace',
            content:
              'You must call the echo tool with the text "hello". Always report your intent before acting.',
          },
          tools: [
            {
              name: 'echo',
              description: 'Echoes a message back',
              parameters: {
                type: 'object' as const,
                properties: { text: { type: 'string', description: 'Text to echo' } },
                required: ['text'],
              },
              handler: (args: unknown) => {
                return Promise.resolve({
                  textResultForLlm: `Echo: ${(args as { text: string }).text}`,
                  resultType: 'success' as const,
                  toolTelemetry: {},
                });
              },
            },
          ],
        });

        const intentTexts: string[] = [];
        const eventTypes: string[] = [];

        for await (const event of session.query({
          prompt: 'Please call the echo tool with "hello".',
        })) {
          eventTypes.push(event.type);
          // Capture report_intent tool calls (before they're filtered by the hook)
          if (event.type === 'tool.execution_start') {
            const data = event.data as { toolName: string; arguments?: Record<string, unknown> };
            if (data.toolName === 'report_intent' && typeof data.arguments?.intent === 'string') {
              intentTexts.push(data.arguments.intent);
            }
          }
          if (event.type === 'session.idle') break;
        }

        // report_intent should fire at least once during a tool-calling turn
        expect(intentTexts.length).toBeGreaterThanOrEqual(1);
        // Intent text should be a non-empty descriptive string
        expect(intentTexts[0].length).toBeGreaterThan(0);
      } finally {
        await client.stop();
      }
    },
    TIMEOUT_MS
  );

  it(
    'npmSkillPackages: skillpm-skill is installed by proxy and its content reaches the model',
    async () => {
      // This test verifies the full npmSkillPackages pipeline end-to-end:
      //   1. The proxy receives npmSkillPackages = ['skillpm-skill']
      //   2. It runs `npm install skillpm-skill` into a temp dir
      //   3. It discovers skills/skillpm/SKILL.md via findSkillDirs
      //   4. It passes that dir to the SDK's skillDirectories
      //   5. The model can answer a question that requires knowledge from that SKILL.md
      //
      // skillpm-skill's SKILL.md teaches agents how to use the skillpm CLI.
      // The install command `npx skillpm install <skill-name>` is a concrete,
      // unambiguous fact from that file that the base model is unlikely to know
      // confidently without the skill context being injected.
      const client = await createWebSocketClient(SERVER_URL);
      try {
        const session = await client.createSession({
          systemMessage: {
            mode: 'replace',
            content:
              'You are a helpful assistant. Answer questions concisely. ' +
              'Use any agent skills loaded in your context.',
          },
          npmSkillPackages: ['skillpm-skill'],
        });

        let fullText = '';
        for await (const event of session.query({
          prompt:
            'I have the skillpm skill loaded. ' +
            'What is the exact CLI command to install a skill package called "my-skill" using skillpm? ' +
            'Just give me the command.',
        })) {
          if (event.type === 'assistant.message_delta') {
            fullText += event.data.deltaContent;
          }
          if (event.type === 'assistant.message') {
            fullText = event.data.content;
          }
          if (event.type === 'session.idle') break;
        }

        // The skillpm SKILL.md documents: `npx skillpm install <skill-name>`
        // or `skillpm install <skill-name>`. Verify the model produced the right command.
        const lower = fullText.toLowerCase();
        expect(lower).toContain('skillpm');
        expect(lower).toContain('install');
        expect(lower).toContain('my-skill');
      } finally {
        await client.stop();
      }
    },
    TIMEOUT_SKILLPM_MS
  );
});
