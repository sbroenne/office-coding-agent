/**
 * Discovers agents, skills, prompts, and MCP servers from sandboxed Copilot CLI plugins.
 *
 * The default config path is the Office Coding Agent sandbox, not the user's
 * terminal Copilot home. Tests can still pass an explicit configPath.
 */

import { existsSync } from 'node:fs';
import { readdir, readFile } from 'node:fs/promises';
import { basename, isAbsolute, join, resolve } from 'node:path';
import { getCopilotConfigPath } from './plugins/pluginPaths.mjs';

/** Path to the sandboxed Copilot CLI config file. */
export const COPILOT_CONFIG_PATH = getCopilotConfigPath();

export class ConfigReadError extends Error {
  constructor(message, cause) {
    super(message);
    this.name = 'ConfigReadError';
    this.cause = cause;
  }
}

export function slugify(name) {
  const slug = String(name)
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '-')
    .replace(/^-+|-+$/g, '');
  return slug || 'skill';
}

export async function readCopilotConfig(configPath = COPILOT_CONFIG_PATH) {
  let raw;
  try {
    raw = await readFile(configPath, 'utf8');
  } catch (err) {
    const code = err && typeof err === 'object' && typeof err.code === 'string' ? err.code : undefined;
    if (code === 'ENOENT') return {};
    throw new ConfigReadError(`Failed to read ${configPath}: ${String(err)}`, err);
  }

  try {
    const stripped = raw.replace(/^\s*\/\/[^\n]*\n/gm, '');
    const parsed = JSON.parse(stripped);
    return parsed && typeof parsed === 'object' && !Array.isArray(parsed) ? parsed : {};
  } catch (err) {
    throw new ConfigReadError(`Failed to parse ${configPath}: ${String(err)}`, err);
  }
}

export async function listAllPluginConfigs(configPath = COPILOT_CONFIG_PATH) {
  const config = await readCopilotConfig(configPath);
  const plugins = Array.isArray(config.installedPlugins)
    ? config.installedPlugins
    : Array.isArray(config.installed_plugins)
      ? config.installed_plugins
      : [];
  return plugins.filter(plugin => plugin && typeof plugin === 'object');
}

export async function readPluginManifest(pluginDir) {
  const candidates = [
    join(pluginDir, 'plugin.json'),
    join(pluginDir, '.github', 'plugin', 'plugin.json'),
    join(pluginDir, '.claude-plugin', 'plugin.json'),
  ];
  const results = await Promise.all(
    candidates.map(async p => {
      try {
        const raw = await readFile(p, 'utf8');
        const parsed = JSON.parse(raw);
        return parsed && typeof parsed === 'object' && !Array.isArray(parsed) ? parsed : null;
      } catch {
        return null;
      }
    })
  );
  return results.find(Boolean) ?? null;
}

function parseFrontmatterListValue(value) {
  const trimmed = value.trim();
  if (!trimmed) return [];
  if (trimmed.startsWith('[') && trimmed.endsWith(']')) {
    return trimmed
      .slice(1, -1)
      .split(',')
      .map(item => item.trim().replace(/^['"]|['"]$/g, ''))
      .filter(Boolean);
  }
  return [trimmed.replace(/^['"]|['"]$/g, '')];
}

function parseFrontmatterScalarValue(lines, index, line, value) {
  if (value !== '>' && value !== '|' && value !== '>-' && value !== '|-') {
    return { value: value.replace(/^['"]|['"]$/g, ''), nextIndex: index };
  }

  const collected = [];
  const baseIndent = line.match(/^\s*/)?.[0].length ?? 0;
  let nextIndex = index;
  while (nextIndex + 1 < lines.length) {
    const next = lines[nextIndex + 1];
    if (!next.trim()) {
      nextIndex++;
      collected.push('');
      continue;
    }
    const nextIndent = next.match(/^\s*/)?.[0].length ?? 0;
    if (nextIndent <= baseIndent) break;
    collected.push(next.trim());
    nextIndex++;
  }

  return {
    value: collected.join(value.startsWith('|') ? '\n' : ' ').trim(),
    nextIndex,
  };
}

export function parsePluginAgentFrontmatter(raw) {
  const trimmed = raw.trimStart();
  const metadata = { description: '', hosts: [], tools: undefined };
  const match = /^---\r?\n([\s\S]*?)\r?\n---/.exec(trimmed);
  if (!match) return metadata;

  let currentListKey = null;
  const lines = match[1].split(/\r?\n/);
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    const trimmedLine = line.trim();
    if (!trimmedLine) continue;

    if (trimmedLine.startsWith('- ') && currentListKey) {
      const values = parseFrontmatterListValue(trimmedLine.slice(2));
      if (currentListKey === 'hosts') metadata.hosts.push(...values);
      if (currentListKey === 'tools') metadata.tools = [...(metadata.tools ?? []), ...values];
      continue;
    }

    currentListKey = null;
    const colonIdx = trimmedLine.indexOf(':');
    if (colonIdx === -1) continue;

    const key = trimmedLine.slice(0, colonIdx).trim();
    const value = trimmedLine.slice(colonIdx + 1).trim();

    if (key === 'description') {
      const parsed = parseFrontmatterScalarValue(lines, i, line, value);
      metadata.description = parsed.value;
      i = parsed.nextIndex;
      continue;
    }

    if (key === 'hosts' || key === 'tools') {
      if (!value) {
        currentListKey = key;
        continue;
      }
      const values = parseFrontmatterListValue(value);
      if (key === 'hosts') metadata.hosts.push(...values);
      if (key === 'tools') metadata.tools = [...(metadata.tools ?? []), ...values];
    }
  }

  metadata.hosts = Array.from(new Set(metadata.hosts));
  if (metadata.tools) metadata.tools = Array.from(new Set(metadata.tools));
  return metadata;
}

function parsePluginPromptFrontmatter(raw) {
  const trimmed = raw.trimStart();
  const metadata = { description: '', agent: '', argumentHint: '' };
  const match = /^---\r?\n([\s\S]*?)\r?\n---/.exec(trimmed);
  if (!match) return metadata;

  const lines = match[1].split(/\r?\n/);
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    const trimmedLine = line.trim();
    if (!trimmedLine) continue;
    const colonIdx = trimmedLine.indexOf(':');
    if (colonIdx === -1) continue;
    const key = trimmedLine.slice(0, colonIdx).trim();
    const value = trimmedLine.slice(colonIdx + 1).trim();
    const parsed = parseFrontmatterScalarValue(lines, i, line, value);
    i = parsed.nextIndex;
    if (key === 'description') metadata.description = parsed.value;
    if (key === 'agent') metadata.agent = parsed.value;
    if (key === 'argument-hint' || key === 'argumentHint') metadata.argumentHint = parsed.value;
  }
  return metadata;
}

function parsePluginSkillFrontmatter(raw, fallbackName) {
  const trimmed = raw.trimStart();
  const metadata = { name: fallbackName, description: '', version: '0.0.0', hosts: [] };
  const match = /^---\r?\n([\s\S]*?)\r?\n---/.exec(trimmed);
  if (!match) return { metadata, content: trimmed };

  const lines = match[1].split(/\r?\n/);
  let currentListKey = null;
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    const trimmedLine = line.trim();
    if (!trimmedLine) continue;

    if (trimmedLine.startsWith('- ') && currentListKey === 'hosts') {
      metadata.hosts.push(...parseFrontmatterListValue(trimmedLine.slice(2)));
      continue;
    }

    currentListKey = null;
    const colonIdx = trimmedLine.indexOf(':');
    if (colonIdx === -1) continue;
    const key = trimmedLine.slice(0, colonIdx).trim();
    const value = trimmedLine.slice(colonIdx + 1).trim();
    if (key === 'hosts' && !value) {
      currentListKey = 'hosts';
      continue;
    }

    const parsed = parseFrontmatterScalarValue(lines, i, line, value);
    i = parsed.nextIndex;
    if (key === 'name') metadata.name = parsed.value || fallbackName;
    if (key === 'description') metadata.description = parsed.value;
    if (key === 'version') metadata.version = parsed.value || '0.0.0';
    if (key === 'hosts') metadata.hosts.push(...parseFrontmatterListValue(parsed.value));
  }

  metadata.hosts = Array.from(new Set(metadata.hosts));
  return { metadata, content: trimmed.slice(match[0].length).trim() };
}

function pluginComponentPaths(manifestValue, fallback) {
  if (!manifestValue) return [fallback];
  return (Array.isArray(manifestValue) ? manifestValue : [manifestValue]).filter(
    value => typeof value === 'string' && value.length > 0
  );
}

function resolvePluginPath(pluginDir, candidate) {
  return isAbsolute(candidate) ? candidate : resolve(pluginDir, candidate);
}

export async function discoverPluginSkillDirs(host, configPath = COPILOT_CONFIG_PATH) {
  const plugins = await listAllPluginConfigs(configPath);
  const skillDirs = [];

  for (const plugin of plugins) {
    if (!plugin.enabled || !plugin.cache_path || !existsSync(plugin.cache_path)) continue;
    if (host && !isPluginForHost(plugin.name, host)) continue;
    const manifest = await readPluginManifest(plugin.cache_path);
    skillDirs.push(...(await skillDirectoriesForPlugin(plugin, manifest)));
  }
  return Array.from(new Set(skillDirs));
}

export async function discoverPluginAgents(host, configPath = COPILOT_CONFIG_PATH) {
  const plugins = await listAllPluginConfigs(configPath);
  const agents = [];
  for (const plugin of plugins) {
    if (!plugin.enabled || !plugin.cache_path || !existsSync(plugin.cache_path)) continue;
    if (host && !isPluginForHost(plugin.name, host)) continue;
    const manifest = await readPluginManifest(plugin.cache_path);
    agents.push(...(await discoverAgentsForPlugin(plugin, manifest)));
  }
  return agents;
}

export async function discoverPluginSkillObjects(host, configPath = COPILOT_CONFIG_PATH) {
  const plugins = await listAllPluginConfigs(configPath);
  const skills = [];
  for (const plugin of plugins) {
    if (!plugin.enabled || !plugin.cache_path || !existsSync(plugin.cache_path)) continue;
    if (host && !isPluginForHost(plugin.name, host)) continue;
    const manifest = await readPluginManifest(plugin.cache_path);
    skills.push(...(await discoverSkillsForPlugin(plugin, manifest)));
  }
  return skills;
}

export async function discoverPluginPrompts(host, configPath = COPILOT_CONFIG_PATH) {
  const plugins = await listAllPluginConfigs(configPath);
  const prompts = [];
  for (const plugin of plugins) {
    if (!plugin.enabled || !plugin.cache_path || !existsSync(plugin.cache_path)) continue;
    if (host && !isPluginForHost(plugin.name, host)) continue;
    const manifest = await readPluginManifest(plugin.cache_path);
    prompts.push(...(await discoverPromptsForPlugin(plugin, manifest)));
  }
  return prompts;
}

export async function discoverAgentsForPlugin(plugin, manifest = null) {
  if (!plugin.cache_path || !existsSync(plugin.cache_path)) return [];
  const agents = [];
  const roots = pluginComponentPaths(manifest?.agents, 'agents');

  for (const root of roots) {
    const agentsDir = resolvePluginPath(plugin.cache_path, root);
    let entries;
    try {
      entries = await readdir(agentsDir, { withFileTypes: true });
    } catch {
      continue;
    }

    for (const entry of entries) {
      if (!entry.isFile()) continue;
      if (!entry.name.endsWith('.agent.md') && entry.name !== 'AGENT.md') continue;
      try {
        const content = await readFile(join(agentsDir, entry.name), 'utf8');
        const agentName = entry.name === 'AGENT.md' ? plugin.name : entry.name.replace(/\.agent\.md$/, '');
        const metadata = parsePluginAgentFrontmatter(content);
        agents.push({
          name: agentName,
          description: metadata.description || `Agent from plugin ${plugin.name}`,
          prompt: content,
          hosts: metadata.hosts,
          tools: metadata.tools,
        });
      } catch {
        // skip unreadable agent files
      }
    }
  }

  return agents;
}

export async function discoverSkillsForPlugin(plugin, manifest = null) {
  const skillDirs = await skillDirectoriesForPlugin(plugin, manifest);
  const skills = [];
  for (const skillDir of skillDirs) {
    try {
      const raw = await readFile(join(skillDir, 'SKILL.md'), 'utf8');
      const { metadata, content } = parsePluginSkillFrontmatter(raw, basename(skillDir));
      skills.push({ ...metadata, content, dir: skillDir });
    } catch {
      // skip unreadable skill files
    }
  }
  return skills;
}

export async function discoverPromptsForPlugin(plugin, manifest = null) {
  if (!plugin.cache_path || !existsSync(plugin.cache_path)) return [];
  const prompts = [];
  const roots = pluginComponentPaths(manifest?.commands ?? manifest?.prompts, 'prompts');

  for (const root of roots) {
    const promptsDir = resolvePluginPath(plugin.cache_path, root);
    let entries;
    try {
      entries = await readdir(promptsDir, { withFileTypes: true });
    } catch {
      continue;
    }

    for (const entry of entries) {
      if (!entry.isFile() || !entry.name.endsWith('.prompt.md')) continue;
      try {
        const raw = await readFile(join(promptsDir, entry.name), 'utf8');
        const trimmed = raw.trimStart();
        const metadata = parsePluginPromptFrontmatter(trimmed);
        const body = trimmed.startsWith('---')
          ? trimmed.slice(trimmed.indexOf('---', 3) + 3).trim()
          : trimmed;
        prompts.push({
          name: entry.name.replace(/\.prompt\.md$/, ''),
          description: metadata.description,
          agent: metadata.agent,
          argumentHint: metadata.argumentHint,
          body,
        });
      } catch {
        // skip unreadable prompt files
      }
    }
  }

  return prompts;
}

export async function skillDirectoriesForPlugin(plugin, manifest = null) {
  if (!plugin.cache_path || !existsSync(plugin.cache_path)) return [];
  const roots = pluginComponentPaths(manifest?.skills, 'skills');
  const skillDirs = [];
  for (const root of roots) {
    const skillsRoot = resolvePluginPath(plugin.cache_path, root);
    let entries;
    try {
      entries = await readdir(skillsRoot, { withFileTypes: true });
    } catch {
      continue;
    }
    for (const entry of entries) {
      if (!entry.isDirectory()) continue;
      const skillDir = join(skillsRoot, entry.name);
      if (existsSync(join(skillDir, 'SKILL.md'))) skillDirs.push(skillDir);
    }
  }
  return skillDirs;
}

export async function readMcpServersForPlugin(plugin, manifest = null) {
  if (!plugin.cache_path || !existsSync(plugin.cache_path)) return {};

  const declaration = manifest?.mcpServers;
  if (declaration && typeof declaration === 'object' && !Array.isArray(declaration)) {
    return declaration;
  }

  const candidates = [];
  if (typeof declaration === 'string') candidates.push(resolvePluginPath(plugin.cache_path, declaration));
  candidates.push(join(plugin.cache_path, '.mcp.json'));

  for (const candidate of candidates) {
    try {
      const parsed = JSON.parse(await readFile(candidate, 'utf8'));
      const servers = parsed?.mcpServers ?? parsed?.servers ?? parsed;
      if (servers && typeof servers === 'object' && !Array.isArray(servers)) return servers;
    } catch {
      continue;
    }
  }

  return {};
}

const HOST_PREFIXES = ['excel', 'powerpoint', 'word', 'outlook'];

export function isPluginForHost(pluginName, host) {
  const pluginSlug = slugify(pluginName);
  const hostSlug = slugify(host);
  const pluginHostTarget = HOST_PREFIXES.find(h => pluginSlug.includes(h));
  if (!pluginHostTarget) return true;
  return pluginHostTarget === hostSlug;
}
