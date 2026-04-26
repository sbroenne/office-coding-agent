/**
 * Management tools for plugins and user memory.
 *
 * These are general-purpose tools — not tied to any Office host.
 * manage_plugins wraps the /api/plugins REST endpoints.
 * manage_memory stores per-user facts in the Zustand memory store.
 */

import type { Tool, ToolInvocation, ToolResultObject } from '@github/copilot-sdk';
import * as pluginService from '@/services/plugins/pluginService';
import { useMemoryStore } from '@/stores/memoryStore';

// ─── manage_plugins ───────────────────────────────────────────────────────────

export const managePluginsTool: Tool = {
  name: 'manage_plugins',
  description:
    'Manage sandboxed Copilot plugins — install, uninstall, list, browse, and update plugins. Uninstalling is the disable path. Also manage plugin marketplaces. Actions: "list" (all installed plugins), "browse" (browse a marketplace), "install" (by spec: owner/repo, name@marketplace, or path), "uninstall" (by name), "update" (by name), "update_all", "marketplaces" (list registered marketplaces), "add_marketplace" (register a new marketplace), "remove_marketplace" (remove a marketplace).',
  parameters: {
    type: 'object',
    properties: {
      action: {
        type: 'string',
        description: 'Operation to perform',
        enum: [
          'list',
          'browse',
          'install',
          'uninstall',
          'update',
          'update_all',
          'marketplaces',
          'add_marketplace',
          'remove_marketplace',
        ],
      },
      name: {
        type: 'string',
        description: 'Plugin or marketplace name (uninstall, update, remove_marketplace).',
      },
      spec: {
        type: 'string',
        description:
          'Install spec: owner/repo, name@marketplace, URL, or path (install, add_marketplace).',
      },
      marketplace: { type: 'string', description: 'Marketplace name to browse (browse action).' },
    },
    required: ['action'],
  },
  handler: async (
    args: unknown,
    _invocation: ToolInvocation
  ): Promise<ToolResultObject | string> => {
    const { action, name, spec, marketplace } = args as {
      action: string;
      name?: string;
      spec?: string;
      marketplace?: string;
    };

    try {
      if (action === 'list') {
        const plugins = await pluginService.getInstalledPlugins();
        return JSON.stringify({
          plugins: plugins.map(p => ({
            name: p.name,
            version: p.version,
            enabled: p.enabled,
            marketplace: p.marketplace || 'direct',
            agents: p.components.agentCount,
            skills: p.components.skillCount,
            mcpServers: p.components.mcpServerCount,
          })),
          count: plugins.length,
        });
      }

      if (action === 'browse') {
        const mp = marketplace ?? 'awesome-copilot';
        const plugins = await pluginService.browseMarketplace(mp);
        return JSON.stringify({
          marketplace: mp,
          plugins: plugins.map(p => ({
            name: p.name,
            description: p.description,
            version: p.version,
            installed: p.installed,
          })),
          count: plugins.length,
        });
      }

      if (action === 'install') {
        if (!spec) return JSON.stringify({ error: 'spec is required for install' });
        const result = await pluginService.installPlugin(spec);
        return JSON.stringify(result);
      }

      if (action === 'uninstall') {
        if (!name) return JSON.stringify({ error: 'name is required for uninstall' });
        const result = await pluginService.uninstallPlugin(name);
        return JSON.stringify(result);
      }

      if (action === 'update') {
        if (!name) return JSON.stringify({ error: 'name is required for update' });
        const result = await pluginService.updatePlugin(name);
        return JSON.stringify(result);
      }

      if (action === 'update_all') {
        const result = await pluginService.updateAllPlugins();
        return JSON.stringify(result);
      }

      if (action === 'marketplaces') {
        const mps = await pluginService.getMarketplaces();
        return JSON.stringify({
          marketplaces: mps.map(m => ({
            name: m.name,
            source: m.source,
            isBuiltIn: m.isBuiltIn,
            pluginCount: m.pluginCount,
          })),
          count: mps.length,
        });
      }

      if (action === 'add_marketplace') {
        if (!spec) return JSON.stringify({ error: 'spec is required for add_marketplace' });
        const result = await pluginService.addMarketplace(spec);
        return JSON.stringify(result);
      }

      if (action === 'remove_marketplace') {
        if (!name) return JSON.stringify({ error: 'name is required for remove_marketplace' });
        const result = await pluginService.removeMarketplace(name);
        return JSON.stringify(result);
      }

      return JSON.stringify({ error: `Unknown action: ${action}` });
    } catch (err) {
      return JSON.stringify({
        error: err instanceof Error ? err.message : 'Plugin operation failed',
      });
    }
  },
};

// ─── manage_memory ────────────────────────────────────────────────────────────

export const manageMemoryTool: Tool = {
  name: 'manage_memory',
  description:
    'Remember facts and preferences about the user across conversations. Actions: "save" (store a new fact), "list" (show all memories), "search" (find memories by keyword), "remove" (delete a memory by ID), "clear" (delete all memories). Use this proactively to remember user preferences (colors, fonts, formatting), project context (team names, data sources), and recurring patterns.',
  parameters: {
    type: 'object',
    properties: {
      action: {
        type: 'string',
        description: 'Operation to perform',
        enum: ['save', 'list', 'search', 'remove', 'clear'],
      },
      content: {
        type: 'string',
        description: 'The fact or preference to remember (save action).',
      },
      category: {
        type: 'string',
        description:
          'Optional category: "preference", "context", "style", "correction", or custom (save action).',
      },
      query: { type: 'string', description: 'Search keyword (search action).' },
      id: { type: 'string', description: 'Memory ID to remove (remove action).' },
    },
    required: ['action'],
  },
  handler: (args: unknown): string => {
    const { action, content, category, query, id } = args as {
      action: string;
      content?: string;
      category?: string;
      query?: string;
      id?: string;
    };

    const store = useMemoryStore.getState();

    if (action === 'save') {
      if (!content) return JSON.stringify({ error: 'content is required' });
      const memId = store.addMemory(content, category);
      return JSON.stringify({
        saved: true,
        id: memId,
        message: `Remembered: "${content}"${category ? ` (${category})` : ''}`,
      });
    }

    if (action === 'list') {
      const memories = store.listMemories(category);
      return JSON.stringify({
        count: memories.length,
        memories: memories.map(m => ({
          id: m.id,
          content: m.content,
          category: m.category ?? 'general',
        })),
      });
    }

    if (action === 'search') {
      if (!query) return JSON.stringify({ error: 'query is required' });
      const results = store.searchMemories(query);
      return JSON.stringify({
        count: results.length,
        results: results.map(m => ({
          id: m.id,
          content: m.content,
          category: m.category ?? 'general',
        })),
      });
    }

    if (action === 'remove') {
      if (!id) return JSON.stringify({ error: 'id is required' });
      store.removeMemory(id);
      return JSON.stringify({ removed: true, id });
    }

    if (action === 'clear') {
      store.clearMemories();
      return JSON.stringify({ cleared: true, message: 'All memories deleted.' });
    }

    return JSON.stringify({ error: `Unknown action: ${action}` });
  },
};

/** All management tools — included for every host */
export const managementTools: Tool[] = [managePluginsTool, manageMemoryTool];
