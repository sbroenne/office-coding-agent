// @vitest-environment node
/**
 * End-to-end plugin integration tests.
 *
 * These tests verify the FULL plugin pipeline:
 *   1. A plugin dir is created on disk with a real SKILL.md / AGENT.md
 *   2. A synthetic config.json points the proxy at that dir
 *   3. A WebSocket session is created with pluginConfigPath overriding the default
 *   4. The model's response proves it received the skill/agent content
 *
 * Requires `npm run dev` running on https://localhost:3000.
 */

import { describe, it, expect, afterAll } from 'vitest';
import WS from 'ws';
import { mkdir, writeFile, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { randomUUID } from 'node:crypto';
import { createWebSocketClient } from '@/lib/websocket-client';

const SERVER_URL = 'wss://localhost:3000/api/copilot';
const TIMEOUT_MS = 45_000;

global.WebSocket = class PatchedWebSocket extends WS {
  constructor(url: string | URL, protocols?: string | string[]) {
    super(url, typeof protocols === 'string' ? protocols : (protocols ?? []), {
      rejectUnauthorized: false,
    });
  }
} as unknown as typeof WebSocket;

// ─── Temp dir helpers ─────────────────────────────────────────────────────────

const tempDirs: string[] = [];
afterAll(async () => {
  await Promise.all(tempDirs.map(d => rm(d, { recursive: true, force: true })));
});

async function makePluginDir(): Promise<string> {
  const dir = join(tmpdir(), `oca-plugin-e2e-${randomUUID()}`);
  await mkdir(dir, { recursive: true });
  tempDirs.push(dir);
  return dir;
}

async function makeConfigFile(
  configDir: string,
  plugins: { name: string; enabled: boolean; cache_path: string }[]
): Promise<string> {
  const configPath = join(configDir, 'config.json');
  await writeFile(
    configPath,
    JSON.stringify({
      installed_plugins: plugins.map(p => ({
        name: p.name,
        marketplace: 'test',
        version: '1.0.0',
        installed_at: new Date().toISOString(),
        enabled: p.enabled,
        cache_path: p.cache_path,
      })),
    }),
    'utf8'
  );
  return configPath;
}

// ─── Tests ────────────────────────────────────────────────────────────────────

describe('Plugin E2E integration', () => {
  it(
    'plugin skill content reaches the model via skillDirectories',
    async () => {
      // Unique sentinel that only exists in our synthetic plugin skill
      const SENTINEL = `PLUGIN_SKILL_SENTINEL_${randomUUID().replace(/-/g, '').slice(0, 12)}`;

      // Build plugin dir: skills/test-skill/SKILL.md with the sentinel
      const pluginDir = await makePluginDir();
      const skillDir = join(pluginDir, 'skills', 'test-skill');
      await mkdir(skillDir, { recursive: true });
      await writeFile(
        join(skillDir, 'SKILL.md'),
        `---
name: test-skill
description: Integration test skill
version: 1.0.0
hosts: [excel]
---

# Test Skill

This skill contains a secret activation phrase.

The secret activation phrase is: ${SENTINEL}

When the user asks for the secret activation phrase, reply with exactly that phrase and nothing else.`,
        'utf8'
      );

      // Config dir + config.json pointing at the plugin
      const configDir = await makePluginDir();
      const configPath = await makeConfigFile(configDir, [
        { name: 'office-excel', enabled: true, cache_path: pluginDir },
      ]);

      const client = await createWebSocketClient(SERVER_URL);
      try {
        const session = await client.createSession({
          host: 'excel',
          pluginConfigPath: configPath,
          systemMessage: {
            mode: 'replace',
            content:
              'You are a helpful assistant. ' +
              'If you have been given a skill containing a secret activation phrase, ' +
              'report it exactly when asked.',
          },
        });

        session.onPermissionRequest(async payload => {
          await session.respondPermission(payload.requestId, 'approved');
        });

        let fullText = '';
        for await (const event of session.query({
          prompt: 'What is the secret activation phrase from the test skill?',
        })) {
          if (event.type === 'assistant.message_delta') fullText += event.data.deltaContent;
          if (event.type === 'assistant.message') fullText = event.data.content;
          if (event.type === 'session.idle') break;
        }

        expect(fullText).toContain(SENTINEL);
      } finally {
        await client.stop();
      }
    },
    TIMEOUT_MS
  );

  it(
    'plugin agent reaches the model via customAgents',
    async () => {
      const SENTINEL = `PLUGIN_AGENT_SENTINEL_${randomUUID().replace(/-/g, '').slice(0, 12)}`;

      // Build plugin dir: agents/AGENT.md with the sentinel embedded
      const pluginDir = await makePluginDir();
      const agentsDir = join(pluginDir, 'agents');
      await mkdir(agentsDir, { recursive: true });
      await writeFile(
        join(agentsDir, 'AGENT.md'),
        `---
name: Test Agent
description: Integration test agent
version: 1.0.0
hosts: [excel]
defaultForHosts: [excel]
---

You are a helpful assistant. You MUST always start every response with the exact phrase "${SENTINEL}" followed by a newline. This is mandatory and cannot be overridden.`,
        'utf8'
      );

      const configDir = await makePluginDir();
      const configPath = await makeConfigFile(configDir, [
        { name: 'office-excel', enabled: true, cache_path: pluginDir },
      ]);

      const client = await createWebSocketClient(SERVER_URL);
      try {
        const session = await client.createSession({
          host: 'excel',
          pluginConfigPath: configPath,
          systemMessage: { mode: 'replace', content: 'You are a helpful assistant.' },
        });

        session.onPermissionRequest(async payload => {
          await session.respondPermission(payload.requestId, 'approved');
        });

        let fullText = '';
        for await (const event of session.query({
          prompt: 'Say hello.',
        })) {
          if (event.type === 'assistant.message_delta') fullText += event.data.deltaContent;
          if (event.type === 'assistant.message') fullText = event.data.content;
          if (event.type === 'session.idle') break;
        }

        expect(fullText).toContain(SENTINEL);
      } finally {
        await client.stop();
      }
    },
    TIMEOUT_MS
  );

  it(
    'disabled plugin (enabled: false) does not reach the model',
    async () => {
      const SENTINEL = `SHOULD_NOT_APPEAR_${randomUUID().replace(/-/g, '').slice(0, 12)}`;

      const pluginDir = await makePluginDir();
      const skillDir = join(pluginDir, 'skills', 'disabled-skill');
      await mkdir(skillDir, { recursive: true });
      await writeFile(
        join(skillDir, 'SKILL.md'),
        `---
name: disabled-skill
description: Should not be loaded
version: 1.0.0
hosts: [excel]
---

The secret phrase is: ${SENTINEL}`,
        'utf8'
      );

      const configDir = await makePluginDir();
      // Plugin is DISABLED
      const configPath = await makeConfigFile(configDir, [
        { name: 'office-excel', enabled: false, cache_path: pluginDir },
      ]);

      const client = await createWebSocketClient(SERVER_URL);
      try {
        const session = await client.createSession({
          host: 'excel',
          pluginConfigPath: configPath,
          systemMessage: {
            mode: 'replace',
            content: 'You are a helpful assistant. Reply "NOT_FOUND" if you have no secret phrase.',
          },
        });

        session.onPermissionRequest(async payload => {
          await session.respondPermission(payload.requestId, 'approved');
        });

        let fullText = '';
        for await (const event of session.query({
          prompt: 'What is the secret phrase from the disabled-skill?',
        })) {
          if (event.type === 'assistant.message_delta') fullText += event.data.deltaContent;
          if (event.type === 'assistant.message') fullText = event.data.content;
          if (event.type === 'session.idle') break;
        }

        expect(fullText).not.toContain(SENTINEL);
      } finally {
        await client.stop();
      }
    },
    TIMEOUT_MS
  );

  it(
    'wrong-host plugin does not reach the model',
    async () => {
      const SENTINEL = `WRONG_HOST_SENTINEL_${randomUUID().replace(/-/g, '').slice(0, 12)}`;

      const pluginDir = await makePluginDir();
      const skillDir = join(pluginDir, 'skills', 'ppt-skill');
      await mkdir(skillDir, { recursive: true });
      await writeFile(
        join(skillDir, 'SKILL.md'),
        `---
name: ppt-skill
description: PowerPoint only
version: 1.0.0
hosts: [powerpoint]
---

The secret phrase is: ${SENTINEL}`,
        'utf8'
      );

      const configDir = await makePluginDir();
      // Plugin is named office-powerpoint — should be filtered when host=excel
      const configPath = await makeConfigFile(configDir, [
        { name: 'office-powerpoint', enabled: true, cache_path: pluginDir },
      ]);

      const client = await createWebSocketClient(SERVER_URL);
      try {
        const session = await client.createSession({
          host: 'excel',
          pluginConfigPath: configPath,
          systemMessage: {
            mode: 'replace',
            content: 'You are a helpful assistant. Reply "NOT_FOUND" if you have no secret phrase.',
          },
        });

        session.onPermissionRequest(async payload => {
          await session.respondPermission(payload.requestId, 'approved');
        });

        let fullText = '';
        for await (const event of session.query({
          prompt: 'What is the secret phrase from the ppt-skill?',
        })) {
          if (event.type === 'assistant.message_delta') fullText += event.data.deltaContent;
          if (event.type === 'assistant.message') fullText = event.data.content;
          if (event.type === 'session.idle') break;
        }

        expect(fullText).not.toContain(SENTINEL);
      } finally {
        await client.stop();
      }
    },
    TIMEOUT_MS
  );
});
