/**
 * General-purpose tools included for every Office host.
 */

import type { Tool } from '@github/copilot-sdk';
import { useMemoryStore } from '@/stores/memoryStore';

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

export const managementTools: Tool[] = [manageMemoryTool];
