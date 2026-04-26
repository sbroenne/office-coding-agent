import { EventEmitter } from 'node:events';
import { watch } from 'node:fs';
import { dirname } from 'node:path';
import {
  ConfigReadError,
  discoverAgentsForPlugin,
  discoverPromptsForPlugin,
  discoverSkillsForPlugin,
  isPluginForHost,
  listAllPluginConfigs,
  readMcpServersForPlugin,
  readPluginManifest,
  skillDirectoriesForPlugin,
} from '../pluginDiscovery.mjs';
import { getCopilotConfigPath } from './pluginPaths.mjs';

export class PluginManager extends EventEmitter {
  constructor(opts = {}) {
    super();
    this.configPath = opts.configPath ?? getCopilotConfigPath();
    this.debounceMs = opts.debounceMs ?? 250;
    this.disableWatcher = opts.disableWatcher ?? false;
    this.plugins = [];
    this.refreshPromise = null;
    this.watcher = null;
    this.debounceTimer = null;
    this.started = false;
  }

  async start() {
    if (this.started) return;
    this.started = true;
    await this.refresh();
    if (!this.disableWatcher) this.attachWatcher();
  }

  stop() {
    if (this.debounceTimer) clearTimeout(this.debounceTimer);
    this.debounceTimer = null;
    if (this.watcher) {
      try {
        this.watcher.close();
      } catch {
        // watcher may already be closed
      }
    }
    this.watcher = null;
    this.started = false;
  }

  list() {
    return this.plugins.map(plugin => {
      const { skillDirectoriesResolved, mcpServersResolved, ...publicPlugin } = plugin;
      return publicPlugin;
    });
  }

  listEffective() {
    return this.list().filter(plugin => plugin.enabled);
  }

  getMcpServerConfigs(host) {
    const servers = [];
    for (const plugin of this.plugins) {
      if (!plugin.enabled) continue;
      if (host && !isPluginForHost(plugin.name, host)) continue;
      for (const [name, cfg] of Object.entries(plugin.mcpServersResolved)) {
        const server = { name, description: `From plugin: ${plugin.name}`, ...cfg };
        if (cfg && typeof cfg === 'object' && 'command' in cfg) server.transport = 'stdio';
        else if (cfg && typeof cfg === 'object' && 'url' in cfg) server.transport = cfg.type === 'sse' ? 'sse' : 'http';
        servers.push(server);
      }
    }
    return servers;
  }

  getSessionInputs(host, disabledMcpServerNames = []) {
    const customAgents = [];
    const skillDirectories = [];
    const prompts = [];
    const skills = [];
    const mcpServers = {};
    const agentNames = new Set();
    const skillDirs = new Set();
    const disabledMcp = new Set(disabledMcpServerNames);

    for (const plugin of this.plugins) {
      if (!plugin.enabled) continue;
      if (host && !isPluginForHost(plugin.name, host)) continue;

      for (const agent of plugin.components.agents) {
        if (host && agent.hosts?.length > 0 && !agent.hosts.includes(host)) continue;
        if (agentNames.has(agent.name)) continue;
        agentNames.add(agent.name);
        customAgents.push({
          name: agent.name,
          description: agent.description,
          prompt: agent.prompt,
          hosts: agent.hosts,
          tools: agent.tools?.length > 0 ? agent.tools.slice() : undefined,
        });
      }

      for (const skill of plugin.components.skills) {
        if (host && skill.hosts?.length > 0 && !skill.hosts.includes(host)) continue;
        skills.push(skill);
      }
      for (const prompt of plugin.components.prompts) prompts.push(prompt);

      for (const dir of plugin.skillDirectoriesResolved) {
        if (skillDirs.has(dir)) continue;
        skillDirs.add(dir);
        skillDirectories.push(dir);
      }

      for (const [name, config] of Object.entries(plugin.mcpServersResolved)) {
        if (disabledMcp.has(name) || mcpServers[name]) continue;
        mcpServers[name] = config;
      }
    }

    return { customAgents, skillDirectories, skills, prompts, mcpServers };
  }

  async refresh() {
    const prior = this.refreshPromise;
    const next = (async () => {
      if (prior) await prior.catch(() => undefined);
      try {
        await this.refreshImpl();
      } catch (err) {
        this.emit('error', err);
      }
    })();
    this.refreshPromise = next.finally(() => {
      if (this.refreshPromise === next) this.refreshPromise = null;
    });
    return this.refreshPromise;
  }

  attachWatcher() {
    const onChange = () => {
      if (this.debounceTimer) clearTimeout(this.debounceTimer);
      this.debounceTimer = setTimeout(() => {
        this.debounceTimer = null;
        this.refresh()
          .then(() => this.emit('changed'))
          .catch(() => {});
      }, this.debounceMs);
    };

    try {
      this.watcher = watch(this.configPath, { persistent: false }, onChange);
    } catch {
      try {
        this.watcher = watch(dirname(this.configPath), { persistent: false }, (_evt, filename) => {
          if (filename && !String(filename).endsWith('config.json')) return;
          onChange();
        });
      } catch {
        // Some environments cannot watch the config; API callers can refresh manually.
      }
    }
  }

  async refreshImpl() {
    let configs;
    try {
      configs = await listAllPluginConfigs(this.configPath);
    } catch (err) {
      if (err instanceof ConfigReadError) {
        this.emit('error', err);
        return;
      }
      throw err;
    }

    const enriched = await Promise.all(
      configs.map(async cfg => {
        const manifest = cfg.cache_path ? await readPluginManifest(cfg.cache_path) : null;
        const [agents, skills, prompts, mcpServers, skillDirs] = await Promise.all([
          discoverAgentsForPlugin(cfg, manifest),
          discoverSkillsForPlugin(cfg, manifest),
          discoverPromptsForPlugin(cfg, manifest),
          readMcpServersForPlugin(cfg, manifest),
          skillDirectoriesForPlugin(cfg, manifest),
        ]);
        return { cfg, manifest, agents, skills, prompts, mcpServers, skillDirs };
      })
    );

    this.plugins = enriched.map(e => this.toInstalledPlugin(e));
  }

  toInstalledPlugin(e) {
    return {
      name: e.cfg.name,
      marketplace: e.cfg.marketplace,
      version: e.cfg.version,
      enabled: e.cfg.enabled,
      installedAt: e.cfg.installed_at,
      cachePath: e.cfg.cache_path,
      source: e.cfg.source,
      manifest: e.manifest,
      components: {
        agents: e.agents,
        skills: e.skills,
        prompts: e.prompts,
        agentCount: e.agents.length,
        agentNames: e.agents.map(agent => agent.name),
        skillCount: e.skills.length,
        skillNames: e.skills.map(skill => skill.name),
        commandCount: e.prompts.length,
        hookCount: countHooks(e.manifest?.hooks),
        mcpServerCount: Object.keys(e.mcpServers).length,
        mcpServerNames: Object.keys(e.mcpServers),
      },
      skillDirectoriesResolved: e.skillDirs,
      mcpServersResolved: e.mcpServers,
    };
  }
}

function countHooks(hooks) {
  if (!hooks) return 0;
  if (Array.isArray(hooks)) return hooks.length;
  if (typeof hooks === 'object') {
    if (hooks.hooks && typeof hooks.hooks === 'object') return countHooks(hooks.hooks);
    return Object.values(hooks).reduce((sum, value) => sum + (Array.isArray(value) ? value.length : 0), 0);
  }
  return 0;
}

export const pluginManager = new PluginManager();

export async function getPluginManager() {
  await pluginManager.start();
  return pluginManager;
}
