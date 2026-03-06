// @vitest-environment node
/**
 * Integration tests for pluginDiscovery.mjs.
 *
 * Uses real temp directories on disk — no mocks. Each test builds a realistic
 * plugin layout, writes a synthetic config.json pointing at it, and calls the
 * exported discovery functions to verify they find the right dirs/agents.
 */

import { describe, it, expect, afterEach } from 'vitest';
import { mkdir, writeFile, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { randomUUID } from 'node:crypto';
import {
  readCopilotConfig,
  readPluginManifest,
  discoverPluginSkillDirs,
  discoverPluginAgents,
  isPluginForHost,
  slugify,
} from '@/../src/pluginDiscovery.mjs';

// ─── Helpers ─────────────────────────────────────────────────────────────────

/** Create a temp dir for a test and return its path. */
async function makeTempDir(): Promise<string> {
  const dir = join(tmpdir(), `oca-test-${randomUUID()}`);
  await mkdir(dir, { recursive: true });
  return dir;
}

/** Write a synthetic config.json and return its path. */
async function writeConfig(
  configDir: string,
  plugins: {
    name: string;
    enabled: boolean;
    cache_path: string;
  }[]
): Promise<string> {
  const configPath = join(configDir, 'config.json');
  const config = {
    installed_plugins: plugins.map(p => ({
      name: p.name,
      marketplace: 'test',
      version: '1.0.0',
      installed_at: new Date().toISOString(),
      enabled: p.enabled,
      cache_path: p.cache_path,
    })),
  };
  await writeFile(configPath, JSON.stringify(config, null, 2), 'utf8');
  return configPath;
}

/** Create a plugin dir with a skills/<name>/SKILL.md layout. */
async function makePluginWithSkill(
  baseDir: string,
  pluginName: string,
  skillName: string,
  skillContent = `# ${skillName} skill`
): Promise<string> {
  const pluginDir = join(baseDir, pluginName);
  const skillDir = join(pluginDir, 'skills', skillName);
  await mkdir(skillDir, { recursive: true });
  await writeFile(join(skillDir, 'SKILL.md'), skillContent, 'utf8');
  return pluginDir;
}

/** Create a plugin dir with an agents/AGENT.md or agents/<name>.agent.md. */
async function makePluginWithAgent(
  baseDir: string,
  pluginName: string,
  agentFileName: string,
  agentContent: string
): Promise<string> {
  const pluginDir = join(baseDir, pluginName);
  const agentsDir = join(pluginDir, 'agents');
  await mkdir(agentsDir, { recursive: true });
  await writeFile(join(agentsDir, agentFileName), agentContent, 'utf8');
  return pluginDir;
}

// Collect temp dirs to clean up after each test
const tempDirs: string[] = [];
afterEach(async () => {
  await Promise.all(tempDirs.map(d => rm(d, { recursive: true, force: true })));
  tempDirs.length = 0;
});

// ─── slugify ─────────────────────────────────────────────────────────────────

describe('slugify', () => {
  it('lowercases and replaces spaces with hyphens', () => {
    expect(slugify('Office Excel')).toBe('office-excel');
  });

  it('strips leading and trailing hyphens', () => {
    expect(slugify('  --excel--  ')).toBe('excel');
  });

  it('returns "skill" for empty/whitespace-only input', () => {
    expect(slugify('')).toBe('skill');
    expect(slugify('   ')).toBe('skill');
  });

  it('handles already-slug strings unchanged', () => {
    expect(slugify('office-excel')).toBe('office-excel');
  });
});

// ─── isPluginForHost ─────────────────────────────────────────────────────────

describe('isPluginForHost', () => {
  it('returns true for a plugin whose name contains the host', () => {
    expect(isPluginForHost('office-excel', 'excel')).toBe(true);
  });

  it('returns false for a plugin whose name targets a different host', () => {
    expect(isPluginForHost('office-powerpoint', 'excel')).toBe(false);
    expect(isPluginForHost('office-word', 'excel')).toBe(false);
  });

  it('returns true for a universal plugin (no host in name)', () => {
    expect(isPluginForHost('my-universal-plugin', 'excel')).toBe(true);
    expect(isPluginForHost('my-universal-plugin', 'powerpoint')).toBe(true);
  });

  it('returns true when plugin name matches the host exactly', () => {
    expect(isPluginForHost('excel', 'excel')).toBe(true);
  });

  it('handles all four supported hosts', () => {
    for (const host of ['excel', 'powerpoint', 'word', 'outlook'] as const) {
      expect(isPluginForHost(`office-${host}`, host)).toBe(true);
    }
  });
});

// ─── readCopilotConfig ───────────────────────────────────────────────────────

describe('readCopilotConfig', () => {
  it('returns empty object when file does not exist', async () => {
    const result = await readCopilotConfig('/nonexistent/path/config.json');
    expect(result).toEqual({});
  });

  it('returns parsed JSON from a real config file', async () => {
    const dir = await makeTempDir();
    tempDirs.push(dir);

    const configPath = join(dir, 'config.json');
    await writeFile(configPath, JSON.stringify({ installed_plugins: [] }), 'utf8');

    const result = await readCopilotConfig(configPath);
    expect(result).toEqual({ installed_plugins: [] });
  });

  it('returns empty object for malformed JSON', async () => {
    const dir = await makeTempDir();
    tempDirs.push(dir);

    const configPath = join(dir, 'config.json');
    await writeFile(configPath, '{ bad json !!', 'utf8');

    const result = await readCopilotConfig(configPath);
    expect(result).toEqual({});
  });
});

// ─── readPluginManifest ──────────────────────────────────────────────────────

describe('readPluginManifest', () => {
  it('returns null when no manifest exists', async () => {
    const dir = await makeTempDir();
    tempDirs.push(dir);
    expect(await readPluginManifest(dir)).toBeNull();
  });

  it('reads plugin.json at root', async () => {
    const dir = await makeTempDir();
    tempDirs.push(dir);

    const manifest = { name: 'test-plugin', version: '1.0.0' };
    await writeFile(join(dir, 'plugin.json'), JSON.stringify(manifest), 'utf8');

    expect(await readPluginManifest(dir)).toEqual(manifest);
  });

  it('reads plugin.json from .github/plugin/ subdirectory', async () => {
    const dir = await makeTempDir();
    tempDirs.push(dir);

    const manifestDir = join(dir, '.github', 'plugin');
    await mkdir(manifestDir, { recursive: true });
    const manifest = { name: 'nested-plugin', version: '2.0.0' };
    await writeFile(join(manifestDir, 'plugin.json'), JSON.stringify(manifest), 'utf8');

    expect(await readPluginManifest(dir)).toEqual(manifest);
  });

  it('prefers root plugin.json over .github/plugin/plugin.json', async () => {
    const dir = await makeTempDir();
    tempDirs.push(dir);

    await writeFile(join(dir, 'plugin.json'), JSON.stringify({ name: 'root' }), 'utf8');
    const manifestDir = join(dir, '.github', 'plugin');
    await mkdir(manifestDir, { recursive: true });
    await writeFile(join(manifestDir, 'plugin.json'), JSON.stringify({ name: 'nested' }), 'utf8');

    const result = await readPluginManifest(dir);
    expect(result?.name).toBe('root');
  });
});

// ─── discoverPluginSkillDirs ─────────────────────────────────────────────────

describe('discoverPluginSkillDirs', () => {
  it('returns empty array when config has no plugins', async () => {
    const dir = await makeTempDir();
    tempDirs.push(dir);
    const configPath = await writeConfig(dir, []);

    const result = await discoverPluginSkillDirs(undefined, configPath);
    expect(result).toEqual([]);
  });

  it('returns empty array when config file does not exist', async () => {
    const result = await discoverPluginSkillDirs(undefined, '/nonexistent/config.json');
    expect(result).toEqual([]);
  });

  it('discovers a skill dir from an enabled plugin', async () => {
    const dir = await makeTempDir();
    tempDirs.push(dir);

    const pluginDir = await makePluginWithSkill(dir, 'office-excel', 'excel');
    const configPath = await writeConfig(dir, [
      { name: 'office-excel', enabled: true, cache_path: pluginDir },
    ]);

    const result = await discoverPluginSkillDirs(undefined, configPath);
    expect(result).toHaveLength(1);
    expect(result[0]).toContain('excel');
  });

  it('skips disabled plugins', async () => {
    const dir = await makeTempDir();
    tempDirs.push(dir);

    const pluginDir = await makePluginWithSkill(dir, 'office-excel', 'excel');
    const configPath = await writeConfig(dir, [
      { name: 'office-excel', enabled: false, cache_path: pluginDir },
    ]);

    const result = await discoverPluginSkillDirs(undefined, configPath);
    expect(result).toEqual([]);
  });

  it('skips plugins where cache_path does not exist on disk', async () => {
    const dir = await makeTempDir();
    tempDirs.push(dir);

    const configPath = await writeConfig(dir, [
      { name: 'office-excel', enabled: true, cache_path: '/does/not/exist' },
    ]);

    const result = await discoverPluginSkillDirs(undefined, configPath);
    expect(result).toEqual([]);
  });

  it('filters out plugins for a different host', async () => {
    const dir = await makeTempDir();
    tempDirs.push(dir);

    const pptDir = await makePluginWithSkill(dir, 'office-powerpoint', 'powerpoint');
    const configPath = await writeConfig(dir, [
      { name: 'office-powerpoint', enabled: true, cache_path: pptDir },
    ]);

    const result = await discoverPluginSkillDirs('excel', configPath);
    expect(result).toEqual([]);
  });

  it('includes universal plugins (no host in name) for any host', async () => {
    const dir = await makeTempDir();
    tempDirs.push(dir);

    const universalDir = await makePluginWithSkill(dir, 'my-universal-plugin', 'shared-skill');
    const configPath = await writeConfig(dir, [
      { name: 'my-universal-plugin', enabled: true, cache_path: universalDir },
    ]);

    expect(await discoverPluginSkillDirs('excel', configPath)).toHaveLength(1);
    expect(await discoverPluginSkillDirs('powerpoint', configPath)).toHaveLength(1);
  });

  it('discovers multiple skills from a single plugin', async () => {
    const dir = await makeTempDir();
    tempDirs.push(dir);

    const pluginDir = join(dir, 'multi-skill-plugin');
    for (const skillName of ['skill-a', 'skill-b', 'skill-c']) {
      const skillDir = join(pluginDir, 'skills', skillName);
      await mkdir(skillDir, { recursive: true });
      await writeFile(join(skillDir, 'SKILL.md'), `# ${skillName}`, 'utf8');
    }
    const configPath = await writeConfig(dir, [
      { name: 'multi-skill-plugin', enabled: true, cache_path: pluginDir },
    ]);

    const result = await discoverPluginSkillDirs(undefined, configPath);
    expect(result).toHaveLength(3);
  });

  it('only returns dirs that contain SKILL.md', async () => {
    const dir = await makeTempDir();
    tempDirs.push(dir);

    const pluginDir = join(dir, 'office-excel');
    // One dir with SKILL.md, one without
    const withSkill = join(pluginDir, 'skills', 'real-skill');
    const withoutSkill = join(pluginDir, 'skills', 'empty-dir');
    await mkdir(withSkill, { recursive: true });
    await mkdir(withoutSkill, { recursive: true });
    await writeFile(join(withSkill, 'SKILL.md'), '# Real skill', 'utf8');

    const configPath = await writeConfig(dir, [
      { name: 'office-excel', enabled: true, cache_path: pluginDir },
    ]);

    const result = await discoverPluginSkillDirs(undefined, configPath);
    expect(result).toHaveLength(1);
    expect(result[0]).toContain('real-skill');
  });

  it('discovers skills from multiple enabled plugins', async () => {
    const dir = await makeTempDir();
    tempDirs.push(dir);

    const excelDir = await makePluginWithSkill(dir, 'office-excel', 'excel');
    const wordDir = await makePluginWithSkill(dir, 'office-word', 'word');
    const configPath = await writeConfig(dir, [
      { name: 'office-excel', enabled: true, cache_path: excelDir },
      { name: 'office-word', enabled: true, cache_path: wordDir },
    ]);

    const result = await discoverPluginSkillDirs(undefined, configPath);
    expect(result).toHaveLength(2);
  });
});

// ─── discoverPluginAgents ────────────────────────────────────────────────────

describe('discoverPluginAgents', () => {
  it('returns empty array when config has no plugins', async () => {
    const dir = await makeTempDir();
    tempDirs.push(dir);
    const configPath = await writeConfig(dir, []);

    const result = await discoverPluginAgents(undefined, configPath);
    expect(result).toEqual([]);
  });

  it('discovers an agent from AGENT.md and uses plugin name as agent name', async () => {
    const dir = await makeTempDir();
    tempDirs.push(dir);

    const agentContent = `---
name: Excel
description: Default Excel agent
version: 1.0.0
hosts: [excel]
defaultForHosts: [excel]
---
Excel agent instructions.`;

    const pluginDir = await makePluginWithAgent(dir, 'office-excel', 'AGENT.md', agentContent);
    const configPath = await writeConfig(dir, [
      { name: 'office-excel', enabled: true, cache_path: pluginDir },
    ]);

    const result = await discoverPluginAgents(undefined, configPath);
    expect(result).toHaveLength(1);
    expect(result[0].name).toBe('office-excel');
    expect(result[0].description).toBe('Default Excel agent');
    expect(result[0].prompt).toContain('Excel agent instructions.');
  });

  it('discovers an agent from a named <name>.agent.md file', async () => {
    const dir = await makeTempDir();
    tempDirs.push(dir);

    const agentContent = `---
description: A named agent
---
Named agent instructions.`;

    const pluginDir = await makePluginWithAgent(dir, 'office-excel', 'my-agent.agent.md', agentContent);
    const configPath = await writeConfig(dir, [
      { name: 'office-excel', enabled: true, cache_path: pluginDir },
    ]);

    const result = await discoverPluginAgents(undefined, configPath);
    expect(result).toHaveLength(1);
    expect(result[0].name).toBe('my-agent');
  });

  it('skips disabled plugins', async () => {
    const dir = await makeTempDir();
    tempDirs.push(dir);

    const pluginDir = await makePluginWithAgent(dir, 'office-excel', 'AGENT.md', '# Agent');
    const configPath = await writeConfig(dir, [
      { name: 'office-excel', enabled: false, cache_path: pluginDir },
    ]);

    const result = await discoverPluginAgents(undefined, configPath);
    expect(result).toEqual([]);
  });

  it('filters out agents from plugins targeting a different host', async () => {
    const dir = await makeTempDir();
    tempDirs.push(dir);

    const pptDir = await makePluginWithAgent(dir, 'office-powerpoint', 'AGENT.md', '# PPT agent');
    const configPath = await writeConfig(dir, [
      { name: 'office-powerpoint', enabled: true, cache_path: pptDir },
    ]);

    const result = await discoverPluginAgents('excel', configPath);
    expect(result).toEqual([]);
  });

  it('uses a fallback description when frontmatter has none', async () => {
    const dir = await makeTempDir();
    tempDirs.push(dir);

    const pluginDir = await makePluginWithAgent(
      dir,
      'office-excel',
      'AGENT.md',
      'No frontmatter here.'
    );
    const configPath = await writeConfig(dir, [
      { name: 'office-excel', enabled: true, cache_path: pluginDir },
    ]);

    const result = await discoverPluginAgents(undefined, configPath);
    expect(result[0].description).toContain('office-excel');
  });

  it('discovers agents from multiple plugins', async () => {
    const dir = await makeTempDir();
    tempDirs.push(dir);

    const excelDir = await makePluginWithAgent(dir, 'office-excel', 'AGENT.md', '# Excel');
    const wordDir = await makePluginWithAgent(dir, 'office-word', 'AGENT.md', '# Word');
    const configPath = await writeConfig(dir, [
      { name: 'office-excel', enabled: true, cache_path: excelDir },
      { name: 'office-word', enabled: true, cache_path: wordDir },
    ]);

    const result = await discoverPluginAgents(undefined, configPath);
    expect(result).toHaveLength(2);
    const names = result.map((a: { name: string }) => a.name);
    expect(names).toContain('office-excel');
    expect(names).toContain('office-word');
  });

  it('ignores non-.agent.md files in agents/ directory', async () => {
    const dir = await makeTempDir();
    tempDirs.push(dir);

    const pluginDir = join(dir, 'office-excel');
    const agentsDir = join(pluginDir, 'agents');
    await mkdir(agentsDir, { recursive: true });
    await writeFile(join(agentsDir, 'README.md'), '# Not an agent', 'utf8');
    await writeFile(join(agentsDir, 'config.json'), '{}', 'utf8');

    const configPath = await writeConfig(dir, [
      { name: 'office-excel', enabled: true, cache_path: pluginDir },
    ]);

    const result = await discoverPluginAgents(undefined, configPath);
    expect(result).toEqual([]);
  });
});
