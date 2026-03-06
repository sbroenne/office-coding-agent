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
import { join } from 'node:path';
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
 * {name, description, prompt, hosts} objects compatible with the SDK's customAgents.
 *
 * The `hosts` field is extracted from the AGENT.md frontmatter when present.
 * An empty `hosts` array means "all hosts" — the caller (browser-side agentService)
 * treats agents with no hosts as universal agents visible for every host.
 *
 * @param {string} [host] - Office host slug for filtering
 * @param {string} [configPath] - path to config.json (defaults to COPILOT_CONFIG_PATH)
 * @returns {Promise<Array<{name: string, description: string, prompt: string, hosts: string[]}>>}
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

        let description = `Agent from plugin ${plugin.name}`;
        let hosts = [];
        const fmMatch = content.match(/^---\n([\s\S]*?)\n---/);
        if (fmMatch) {
          const descMatch = fmMatch[1].match(/description:\s*(.+)/);
          if (descMatch) description = descMatch[1].trim();

          // Extract hosts from frontmatter inline array: hosts: [excel, word]
          const hostsMatch = fmMatch[1].match(/hosts:\s*\[([^\]]*)\]/);
          if (hostsMatch) {
            hosts = hostsMatch[1]
              .split(',')
              .map(h => h.trim())
              .filter(Boolean);
          }
        }

        agents.push({ name: agentName, description, prompt: content, hosts });
      } catch {
        // skip unreadable agent files
      }
    }
  }
  return agents;
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
