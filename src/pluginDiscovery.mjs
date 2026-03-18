/**
 * pluginDiscovery.mjs — discovers skills and agents from installed Copilot CLI plugins.
 *
 * Extracted from copilotProxy.mjs so these functions can be tested independently.
 *
 * Plugin layout (under ~/.copilot/installed-plugins/<marketplace>/<name>/):
 *   skills/<skill-name>/SKILL.md   — skill files loaded via skillDirectories
 *   agents/AGENT.md                — agent definition (name = plugin name)
 *   agents/<name>.agent.md         — named agent definition
 */

import { readdir, readFile } from 'node:fs/promises';
import { existsSync } from 'node:fs';
import { join, basename } from 'node:path';
import { homedir } from 'node:os';

/** Path to the Copilot CLI config file. */
export const COPILOT_CONFIG_PATH = join(homedir(), '.copilot', 'config.json');

/** Convert a name to a safe lowercase directory slug. */
export function slugify(name) {
  const slug = name
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '-')
    .replace(/^-+|-+$/g, '');
  return slug || 'skill';
}

/**
 * Read a Copilot CLI config file and return its parsed contents.
 * Returns a safe default if the file doesn't exist or is malformed.
 *
 * @param {string} [configPath] - path to config.json (defaults to COPILOT_CONFIG_PATH)
 * @returns {Promise<{installed_plugins?: Array<{name: string, marketplace: string, version: string, installed_at: string, enabled: boolean, cache_path: string}>}>}
 */
export async function readCopilotConfig(configPath = COPILOT_CONFIG_PATH) {
  try {
    const raw = await readFile(configPath, 'utf8');
    return JSON.parse(raw);
  } catch {
    return {};
  }
}

/**
 * Find the plugin.json manifest in a plugin cache directory.
 * Checks standard locations: plugin.json, .github/plugin/plugin.json.
 *
 * @param {string} pluginDir
 * @returns {Promise<object|null>}
 */
export async function readPluginManifest(pluginDir) {
  const candidates = [
    join(pluginDir, 'plugin.json'),
    join(pluginDir, '.github', 'plugin', 'plugin.json'),
  ];
  for (const p of candidates) {
    try {
      const raw = await readFile(p, 'utf8');
      return JSON.parse(raw);
    } catch {
      continue;
    }
  }
  return null;
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

/**
 * Parse the subset of AGENT.md frontmatter needed by plugin discovery.
 *
 * @param {string} raw
 * @returns {{description: string, hosts: string[], tools?: string[]}}
 */
export function parsePluginAgentFrontmatter(raw) {
  const trimmed = raw.trimStart();
  const metadata = { description: '', hosts: [], tools: undefined };
  const match = trimmed.match(/^---\r?\n([\s\S]*?)\r?\n---/);
  if (!match) return metadata;

  /** @type {'hosts' | 'tools' | null} */
  let currentListKey = null;

  for (const line of match[1].split(/\r?\n/)) {
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
      metadata.description = value.replace(/^['"]|['"]$/g, '');
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
  if (metadata.tools) {
    metadata.tools = Array.from(new Set(metadata.tools));
  }

  return metadata;
}

/**
 * Discover skill directories from installed Copilot CLI plugins.
 *
 * Reads config → installed_plugins[], then for each enabled plugin with a
 * cache_path, scans <cache_path>/skills/ for subdirectories containing SKILL.md
 * files (same layout as bundled skills and skillpm packages).
 *
 * Host filtering: plugins whose name contains a different host slug (e.g.
 * "office-powerpoint" when host is "excel") are skipped. Universal plugins
 * (no recognisable host in their name) are always included.
 *
 * @param {string} [host] - Office host slug (e.g. 'excel', 'powerpoint')
 * @param {string} [configPath] - path to config.json (defaults to COPILOT_CONFIG_PATH)
 * @returns {Promise<string[]>} array of skill directory paths (each containing SKILL.md)
 */
export async function discoverPluginSkillDirs(host, configPath = COPILOT_CONFIG_PATH) {
  const config = await readCopilotConfig(configPath);
  const plugins = config.installed_plugins || [];
  const skillDirs = [];

  for (const plugin of plugins) {
    if (!plugin.enabled || !plugin.cache_path) continue;
    if (!existsSync(plugin.cache_path)) continue;

    if (host && !isPluginForHost(plugin.name, host)) continue;

    const skillsRoot = join(plugin.cache_path, 'skills');
    let entries;
    try {
      entries = await readdir(skillsRoot, { withFileTypes: true });
    } catch {
      continue;
    }

    for (const entry of entries) {
      if (entry.isDirectory()) {
        const skillDir = join(skillsRoot, entry.name);
        if (existsSync(join(skillDir, 'SKILL.md'))) {
          skillDirs.push(skillDir);
        }
      }
    }
  }
  return skillDirs;
}

/**
 * Discover custom agent definitions from installed Copilot CLI plugins.
 *
 * For each enabled plugin with a cache_path, scans <cache_path>/agents/ for
 * *.agent.md and AGENT.md files. Returns an array of
 * {name, description, prompt, hosts, tools} objects compatible with the SDK's customAgents.
 *
 * The `hosts` field is extracted from the AGENT.md frontmatter when present.
 * An empty `hosts` array means "all hosts" — the caller (browser-side agentService)
 * treats agents with no hosts as universal agents visible for every host.
 *
 * @param {string} [host] - Office host slug for filtering
 * @param {string} [configPath] - path to config.json (defaults to COPILOT_CONFIG_PATH)
 * @returns {Promise<Array<{name: string, description: string, prompt: string, hosts: string[], tools?: string[]}>>}
 */
export async function discoverPluginAgents(host, configPath = COPILOT_CONFIG_PATH) {
  const config = await readCopilotConfig(configPath);
  const plugins = config.installed_plugins || [];
  const agents = [];

  for (const plugin of plugins) {
    if (!plugin.enabled || !plugin.cache_path) continue;
    if (!existsSync(plugin.cache_path)) continue;

    if (host && !isPluginForHost(plugin.name, host)) continue;

    const agentsDir = join(plugin.cache_path, 'agents');
    let entries;
    try {
      entries = await readdir(agentsDir, { withFileTypes: true });
    } catch {
      continue;
    }

    for (const entry of entries) {
      if (!entry.isFile()) continue;
      const isAgentMd = entry.name.endsWith('.agent.md') || entry.name === 'AGENT.md';
      if (!isAgentMd) continue;

      try {
        const content = await readFile(join(agentsDir, entry.name), 'utf8');
        const agentName =
          entry.name === 'AGENT.md'
            ? plugin.name
            : entry.name.replace(/\.agent\.md$/, '');

        const metadata = parsePluginAgentFrontmatter(content);
        const description = metadata.description || `Agent from plugin ${plugin.name}`;

        agents.push({
          name: agentName,
          description,
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

/**
 * Discover plugin skills as parsed objects (name, description, hosts, content).
 * Reads SKILL.md from each skill directory found via discoverPluginSkillDirs.
 *
 * @param {string} [host]
 * @param {string} [configPath]
 * @returns {Promise<Array<{name: string, description: string, version: string, hosts: string[], content: string}>>}
 */
export async function discoverPluginSkillObjects(host, configPath = COPILOT_CONFIG_PATH) {
  const skillDirs = await discoverPluginSkillDirs(host, configPath);
  const skills = [];

  for (const skillDir of skillDirs) {
    try {
      const raw = await readFile(join(skillDir, 'SKILL.md'), 'utf8');
      const trimmed = raw.trimStart();

      let name = basename(skillDir);
      let description = '';
      let version = '0.0.0';
      let hosts = [];
      let content = trimmed;

      if (trimmed.startsWith('---')) {
        const endIdx = trimmed.indexOf('---', 3);
        if (endIdx !== -1) {
          const yamlBlock = trimmed.slice(3, endIdx).trim();
          content = trimmed.slice(endIdx + 3).trim();

          for (const line of yamlBlock.split('\n')) {
            const colonIdx = line.indexOf(':');
            if (colonIdx === -1) continue;
            const key = line.slice(0, colonIdx).trim();
            const value = line.slice(colonIdx + 1).trim();
            if (key === 'name') name = value;
            else if (key === 'description')
              description = value.replace(/^['"]|['"]$/g, '');
            else if (key === 'version') version = value;
            else if (key === 'hosts' && value.startsWith('[') && value.endsWith(']')) {
              hosts = value
                .slice(1, -1)
                .split(',')
                .map(h => h.trim())
                .filter(Boolean);
            }
          }
        }
      }

      skills.push({ name, description, version, hosts, content });
    } catch {
      // skip unreadable skill files
    }
  }

  return skills;
}

/**
 * Discover prompt templates from installed plugin prompts/ directories.
 * Each .prompt.md file becomes a slash command in the ChatComposer.
 *
 * @param {string} [host]
 * @param {string} [configPath]
 * @returns {Promise<Array<{name: string, description: string, agent: string, argumentHint: string, body: string}>>}
 */
export async function discoverPluginPrompts(host, configPath = COPILOT_CONFIG_PATH) {
  const config = await readCopilotConfig(configPath);
  const plugins = config.installed_plugins || [];
  const prompts = [];

  for (const plugin of plugins) {
    if (!plugin.enabled || !plugin.cache_path) continue;
    if (!existsSync(plugin.cache_path)) continue;
    if (host && !isPluginForHost(plugin.name, host)) continue;

    const promptsDir = join(plugin.cache_path, 'prompts');
    let entries;
    try {
      entries = await readdir(promptsDir, { withFileTypes: true });
    } catch {
      continue; // no prompts dir — that's fine
    }

    for (const entry of entries) {
      if (!entry.isFile() || !entry.name.endsWith('.prompt.md')) continue;

      try {
        const raw = await readFile(join(promptsDir, entry.name), 'utf8');
        const trimmed = raw.trimStart();

        let name = entry.name.replace(/\.prompt\.md$/, '');
        let description = '';
        let agent = '';
        let argumentHint = '';
        let body = trimmed;

        if (trimmed.startsWith('---')) {
          const endIdx = trimmed.indexOf('---', 3);
          if (endIdx !== -1) {
            const yamlBlock = trimmed.slice(3, endIdx).trim();
            body = trimmed.slice(endIdx + 3).trim();

            for (const line of yamlBlock.split('\n')) {
              const colonIdx = line.indexOf(':');
              if (colonIdx === -1) continue;
              const key = line.slice(0, colonIdx).trim();
              const value = line.slice(colonIdx + 1).trim().replace(/^['"]|['"]$/g, '');
              if (key === 'name') name = value;
              else if (key === 'description') description = value;
              else if (key === 'agent') agent = value;
              else if (key === 'argument-hint') argumentHint = value;
            }
          }
        }

        prompts.push({ name, description, agent, argumentHint, body });
      } catch {
        // skip unreadable files
      }
    }
  }

  return prompts;
}

// ─── Internal helpers ────────────────────────────────────────────────────────

const HOST_PREFIXES = ['excel', 'powerpoint', 'word', 'outlook'];

/**
 * Returns true if the plugin should be included for the given host.
 * A plugin is included when:
 *  - its name contains the host slug (e.g. "office-excel" matches "excel"), OR
 *  - its name contains no recognised host slug at all (universal plugin).
 * A plugin is excluded when its name targets a DIFFERENT recognised host.
 *
 * @param {string} pluginName
 * @param {string} host
 * @returns {boolean}
 */
export function isPluginForHost(pluginName, host) {
  const pluginSlug = slugify(pluginName);
  const hostSlug = slugify(host);
  const pluginHostTarget = HOST_PREFIXES.find(h => pluginSlug.includes(h));
  if (!pluginHostTarget) return true; // universal plugin
  return pluginHostTarget === hostSlug;
}
