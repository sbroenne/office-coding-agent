/**
 * Integration tests: full plugin pipeline regression suite.
 *
 * Verifies all 4 plugin asset types (agents, skills, prompts, MCP servers) are:
 *   - Discoverable from the fixture test-plugin directory
 *   - Stored in the Zustand settings store
 *   - Visible in the UI components (AgentPicker, SkillPicker, ChatComposer)
 *
 * No live server required — uses store manipulation + fixture files on disk.
 */

// @vitest-environment jsdom
import { describe, it, expect, beforeEach, afterAll } from 'vitest';
import { screen } from '@testing-library/react';
import userEvent from '@testing-library/user-event';
import { renderWithProviders } from '../test-utils';
import { AgentPicker } from '@/components/AgentPicker';
import { SkillPicker } from '@/components/SkillPicker';
import { ChatComposer } from '@/components/chat/ChatComposer';
import { useSettingsStore } from '@/stores/settingsStore';
import { parseAgentFrontmatter } from '@/services/agents';
import type { McpServerConfig } from '@/types/mcp';
import { writeFile, mkdir, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import { randomUUID } from 'node:crypto';

// Path to the persistent fixture test-plugin
const FIXTURE_PLUGIN_DIR = resolve(__dirname, '../fixtures/test-plugin');

beforeEach(() => {
  useSettingsStore.getState().reset();
});

// ─── [A] AgentPicker shows plugin agent ───────────────────────────────────────

describe('[A] Plugin agent visible in AgentPicker', () => {
  it('custom plugin agent appears in the dropdown after setPluginAgents', async () => {
    const customAgent = parseAgentFrontmatter(`---
name: Test Excel Agent
description: A regression-test agent that helps with Excel analysis tasks.
version: 1.0.0
hosts: [excel]
defaultForHosts: []
---
You are a specialised Excel analysis agent used in regression tests.`);

    useSettingsStore.getState().setPluginAgents([customAgent]);
    renderWithProviders(<AgentPicker />);

    await userEvent.click(screen.getByLabelText('Select agent'));

    // Plugin agent name must appear in the picker dropdown
    expect(screen.getByText('Test Excel Agent')).toBeInTheDocument();
  });

  it('custom plugin agent is selectable and updates the store', async () => {
    const customAgent = parseAgentFrontmatter(`---
name: Test Excel Agent
description: A regression-test agent that helps with Excel analysis tasks.
version: 1.0.0
hosts: [excel]
defaultForHosts: []
---
You are a specialised Excel analysis agent.`);

    useSettingsStore.getState().setPluginAgents([customAgent]);
    renderWithProviders(<AgentPicker />);

    await userEvent.click(screen.getByLabelText('Select agent'));
    // Click the agent by its text label (existing pattern from agent-picker.test.tsx)
    await userEvent.click(screen.getByText('Test Excel Agent'));

    expect(useSettingsStore.getState().activeAgentId).toBe('Test Excel Agent');
  });

  it('clearing pluginAgents removes the custom agent from the dropdown', async () => {
    const customAgent = parseAgentFrontmatter(`---
name: TemporaryPluginAgent
description: Will be removed.
version: 1.0.0
hosts: [excel]
defaultForHosts: []
---
Temp.`);

    useSettingsStore.getState().setPluginAgents([customAgent]);
    const { unmount } = renderWithProviders(<AgentPicker />);
    unmount();

    useSettingsStore.getState().setPluginAgents([]);
    renderWithProviders(<AgentPicker />);

    await userEvent.click(screen.getByLabelText('Select agent'));
    expect(screen.queryByText('TemporaryPluginAgent')).not.toBeInTheDocument();
  });
});

// ─── [B] SkillPicker shows plugin skill ───────────────────────────────────────

describe('[B] Plugin skill visible in SkillPicker', () => {
  it('plugin skill appears under "Plugin" section after setPluginSkills', async () => {
    useSettingsStore.getState().setPluginSkills([
      {
        metadata: {
          name: 'test-excel-skill',
          description: 'A regression-test skill providing Excel best-practice guidance.',
          version: '1.2.0',
          tags: [],
          hosts: ['excel'],
        },
        content: '# Test Excel Skill\n\nGuidance content.',
      },
    ]);

    renderWithProviders(<SkillPicker />);

    await userEvent.click(screen.getByLabelText('Agent skills'));

    // Skill name must be visible
    expect(screen.getByText('test-excel-skill')).toBeInTheDocument();
  });

  it('clearing pluginSkills removes the skill from the picker', async () => {
    useSettingsStore.getState().setPluginSkills([
      {
        metadata: {
          name: 'temp-plugin-skill',
          description: 'Temporary skill',
          version: '1.0.0',
          tags: [],
          hosts: [],
        },
        content: '# Temp',
      },
    ]);

    const { unmount } = renderWithProviders(<SkillPicker />);
    unmount();

    useSettingsStore.getState().setPluginSkills([]);
    renderWithProviders(<SkillPicker />);

    await userEvent.click(screen.getByLabelText('Agent skills'));
    expect(screen.queryByText('temp-plugin-skill')).not.toBeInTheDocument();
  });
});

// ─── [C] ChatComposer slash menu shows plugin prompts ─────────────────────────

describe('[C] Plugin prompts visible in ChatComposer slash menu', () => {
  it('plugin prompt appears in the slash command menu when "/" is typed', async () => {
    const prompts = [
      {
        name: 'test-excel-prompt',
        description: 'Analyse a named Excel range and summarise the data.',
        agent: 'Test Excel Agent',
        argumentHint: 'Range name (e.g. SalesData)',
        body: 'Please analyse the Excel range "${input:rangeName}".',
      },
    ];

    renderWithProviders(
      <ChatComposer
        onSend={async () => {}}
        onCancel={() => {}}
        isRunning={false}
        slashCommands={prompts}
      />
    );

    const input = screen.getByLabelText('Message input');
    await userEvent.type(input, '/');

    // Prompt name appears prefixed with "/" in the slash command dropdown
    expect(screen.getByRole('listbox', { name: /slash commands/i })).toBeInTheDocument();
    expect(screen.getByText('/test-excel-prompt')).toBeInTheDocument();
  });
});

// ─── [D] Store holds plugin MCP servers ───────────────────────────────────────

describe('[D] pluginMcpServers store field', () => {
  it('setPluginMcpServers populates pluginMcpServers in the store', () => {
    const server: McpServerConfig = {
      name: 'test-excel-mcp',
      description: 'Regression-test MCP server',
      transport: 'stdio',
      command: 'node',
      args: ['--version'],
    };

    useSettingsStore.getState().setPluginMcpServers([server]);

    const stored = useSettingsStore.getState().pluginMcpServers;
    expect(stored).toHaveLength(1);
    expect(stored[0].name).toBe('test-excel-mcp');
    expect(stored[0].command).toBe('node');
    expect(stored[0].transport).toBe('stdio');
  });

  it('reset() clears pluginMcpServers back to empty array', () => {
    useSettingsStore.getState().setPluginMcpServers([
      { name: 'my-mcp', transport: 'http', url: 'http://localhost:3001' },
    ]);

    expect(useSettingsStore.getState().pluginMcpServers).toHaveLength(1);

    useSettingsStore.getState().reset();
    expect(useSettingsStore.getState().pluginMcpServers).toHaveLength(0);
  });

  it('initial pluginMcpServers is an empty array', () => {
    expect(useSettingsStore.getState().pluginMcpServers).toEqual([]);
  });
});

// ─── [E] discoverPluginMcpServers reads fixture mcp.json ──────────────────────

describe('[E] discoverPluginMcpServers reads fixture test-plugin', () => {
  const tempDirs: string[] = [];

  afterAll(async () => {
    await Promise.all(tempDirs.map(d => rm(d, { recursive: true, force: true })));
  });

  it('discovers test-excel-mcp server from fixture plugin mcp.json', async () => {
    const configDir = join(tmpdir(), `oca-plugin-pipe-${randomUUID()}`);
    await mkdir(configDir, { recursive: true });
    tempDirs.push(configDir);

    const configPath = join(configDir, 'config.json');
    await writeFile(
      configPath,
      JSON.stringify({
        installed_plugins: [
          {
            name: 'office-excel-test-plugin',
            marketplace: 'test',
            version: '1.0.0',
            installed_at: new Date().toISOString(),
            enabled: true,
            cache_path: FIXTURE_PLUGIN_DIR,
          },
        ],
      }),
      'utf8'
    );

    const { discoverPluginMcpServers } = await import('../../src/pluginDiscovery.mjs');
    const servers = await discoverPluginMcpServers('excel', configPath);

    expect(servers).toHaveLength(1);
    const srv = servers[0];
    expect(srv.name).toBe('test-excel-mcp');
    expect(srv.transport).toBe('stdio');
    expect(srv.command).toBe('node');
    expect(srv.args).toEqual(['--version']);
  });

  it('returns empty array when plugin is disabled', async () => {
    const configDir = join(tmpdir(), `oca-plugin-pipe-${randomUUID()}`);
    await mkdir(configDir, { recursive: true });
    tempDirs.push(configDir);

    const configPath = join(configDir, 'config.json');
    await writeFile(
      configPath,
      JSON.stringify({
        installed_plugins: [
          {
            name: 'office-excel-test-plugin',
            marketplace: 'test',
            version: '1.0.0',
            installed_at: new Date().toISOString(),
            enabled: false,
            cache_path: FIXTURE_PLUGIN_DIR,
          },
        ],
      }),
      'utf8'
    );

    const { discoverPluginMcpServers } = await import('../../src/pluginDiscovery.mjs');
    const servers = await discoverPluginMcpServers('excel', configPath);
    expect(servers).toHaveLength(0);
  });

  it('returns empty array when host does not match plugin name', async () => {
    const configDir = join(tmpdir(), `oca-plugin-pipe-${randomUUID()}`);
    await mkdir(configDir, { recursive: true });
    tempDirs.push(configDir);

    const configPath = join(configDir, 'config.json');
    await writeFile(
      configPath,
      JSON.stringify({
        installed_plugins: [
          {
            name: 'office-excel-test-plugin',
            marketplace: 'test',
            version: '1.0.0',
            installed_at: new Date().toISOString(),
            enabled: true,
            cache_path: FIXTURE_PLUGIN_DIR,
          },
        ],
      }),
      'utf8'
    );

    const { discoverPluginMcpServers } = await import('../../src/pluginDiscovery.mjs');
    // Host is 'powerpoint' but plugin name contains 'excel' → should be skipped
    const servers = await discoverPluginMcpServers('powerpoint', configPath);
    expect(servers).toHaveLength(0);
  });
});
