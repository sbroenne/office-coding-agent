/**
 * Integration tests for management tools (manage_plugins, manage_memory).
 *
 * manage_plugins calls pluginService REST endpoints — we mock global fetch.
 * manage_memory exercises real Zustand store operations — no mocks.
 */

import { describe, it, expect, beforeEach, vi, type Mock } from 'vitest';
import Ajv from 'ajv';
import {
  managePluginsTool,
  manageMemoryTool,
  managementTools,
} from '@/tools/management';

const ajv = new Ajv({ allErrors: true });

function validate(schema: unknown, data: unknown): boolean {
  return !!ajv.compile(schema as object)(data);
}

/** Call the async manage_plugins handler and parse JSON result */
async function callAsync(
  tool: { handler?: unknown },
  args: Record<string, unknown>,
): Promise<unknown> {
  const handler = (tool as { handler: (a: unknown, i: unknown) => Promise<string> }).handler;
  const raw = await handler(args, {});
  return JSON.parse(typeof raw === 'string' ? raw : JSON.stringify(raw));
}

/** Call the sync manage_memory handler and parse JSON result */
function callSync(tool: { handler?: unknown }, args: Record<string, unknown>): unknown {
  const handler = (tool as { handler: (a: unknown, i: unknown) => string }).handler;
  const raw = handler(args, {});
  return JSON.parse(raw as string);
}

// ─── Mock fetch for pluginService calls ─────────────────────────────────────

let fetchMock: Mock;

beforeEach(() => {
  fetchMock = vi.fn();
  vi.stubGlobal('fetch', fetchMock);
});

/** Helper: make fetch return a successful JSON response */
function mockFetchOk(body: unknown) {
  fetchMock.mockResolvedValueOnce({
    ok: true,
    json: () => Promise.resolve(body),
  });
}

/** Helper: make fetch reject (network failure) */
function mockFetchReject(message: string) {
  fetchMock.mockRejectedValueOnce(new Error(message));
}

// ─── Schema validation ──────────────────────────────────────────────────────

describe('management tool schemas', () => {
  describe('manage_plugins', () => {
    it('accepts list action', () => {
      expect(validate(managePluginsTool.parameters, { action: 'list' })).toBe(true);
    });

    it('accepts install with spec param', () => {
      expect(
        validate(managePluginsTool.parameters, {
          action: 'install',
          spec: 'owner/repo',
        }),
      ).toBe(true);
    });

    it('accepts browse with marketplace param', () => {
      expect(
        validate(managePluginsTool.parameters, {
          action: 'browse',
          marketplace: 'awesome-copilot',
        }),
      ).toBe(true);
    });

    it('accepts uninstall with name param', () => {
      expect(
        validate(managePluginsTool.parameters, {
          action: 'uninstall',
          name: 'my-plugin',
        }),
      ).toBe(true);
    });

    it('accepts update all', () => {
      expect(validate(managePluginsTool.parameters, { action: 'update_all' })).toBe(true);
    });

    it('accepts update with name', () => {
      expect(
        validate(managePluginsTool.parameters, { action: 'update', name: 'my-plugin' }),
      ).toBe(true);
    });

    it('accepts marketplace-management actions', () => {
      expect(validate(managePluginsTool.parameters, { action: 'marketplaces' })).toBe(true);
      expect(
        validate(managePluginsTool.parameters, {
          action: 'add_marketplace',
          spec: 'https://example.com/marketplace.json',
        }),
      ).toBe(true);
      expect(
        validate(managePluginsTool.parameters, {
          action: 'remove_marketplace',
          name: 'custom-mp',
        }),
      ).toBe(true);
    });

    it('rejects missing action', () => {
      expect(validate(managePluginsTool.parameters, { name: 'test' })).toBe(false);
    });

    it('has the expected action enum values', () => {
      const actionProp = (managePluginsTool.parameters as { properties: { action: { enum: string[] } } })
        .properties.action;
      expect(actionProp.enum).toEqual([
        'list',
        'browse',
        'install',
        'uninstall',
        'update',
        'update_all',
        'marketplaces',
        'add_marketplace',
        'remove_marketplace',
      ]);
    });

    it('has name, spec, and marketplace optional parameters', () => {
      const props = (managePluginsTool.parameters as { properties: Record<string, unknown> }).properties;
      expect(props).toHaveProperty('name');
      expect(props).toHaveProperty('spec');
      expect(props).toHaveProperty('marketplace');
    });
  });

  describe('manage_memory', () => {
    it('accepts list action', () => {
      expect(validate(manageMemoryTool.parameters, { action: 'list' })).toBe(true);
    });

    it('accepts save with content', () => {
      expect(
        validate(manageMemoryTool.parameters, { action: 'save', content: 'User prefers dark mode' }),
      ).toBe(true);
    });

    it('accepts search with query', () => {
      expect(
        validate(manageMemoryTool.parameters, { action: 'search', query: 'dark mode' }),
      ).toBe(true);
    });

    it('rejects missing action', () => {
      expect(validate(manageMemoryTool.parameters, { content: 'test' })).toBe(false);
    });
  });
});

// ─── managementTools array ──────────────────────────────────────────────────

describe('managementTools array', () => {
  it('contains exactly manage_plugins and manage_memory', () => {
    const names = managementTools.map(t => t.name);
    expect(names).toEqual(['manage_plugins', 'manage_memory']);
  });

  it('has length 2', () => {
    expect(managementTools).toHaveLength(2);
  });
});

// ─── manage_plugins handler ─────────────────────────────────────────────────

describe('manage_plugins handler', () => {
  it('list returns installed plugins', async () => {
    mockFetchOk({ plugins: [
      {
        name: 'my-plugin',
        version: '1.0.0',
        enabled: true,
        marketplace: 'awesome-copilot',
        components: { agentCount: 1, skillCount: 2, mcpServerCount: 0 },
      },
    ] });

    const result = (await callAsync(managePluginsTool, { action: 'list' })) as {
      plugins: { name: string; enabled: boolean }[];
      count: number;
    };
    expect(result.count).toBe(1);
    expect(result.plugins[0].name).toBe('my-plugin');
    expect(result.plugins[0].enabled).toBe(true);
  });

  it('browse returns marketplace plugins', async () => {
    mockFetchOk({ marketplace: 'my-mp', plugins: [
      { name: 'cool-plugin', description: 'A cool plugin', version: '2.0.0', installed: false },
    ] });

    const result = (await callAsync(managePluginsTool, {
      action: 'browse',
      marketplace: 'my-mp',
    })) as { marketplace: string; plugins: { name: string }[]; count: number };
    expect(result.marketplace).toBe('my-mp');
    expect(result.count).toBe(1);
    expect(result.plugins[0].name).toBe('cool-plugin');
  });

  it('browse defaults to awesome-copilot marketplace', async () => {
    mockFetchOk({ marketplace: 'awesome-copilot', plugins: [] });
    const result = (await callAsync(managePluginsTool, { action: 'browse' })) as {
      marketplace: string;
    };
    expect(result.marketplace).toBe('awesome-copilot');
  });

  it('install requires spec', async () => {
    const result = (await callAsync(managePluginsTool, { action: 'install' })) as {
      error: string;
    };
    expect(result.error).toContain('spec');
  });

  it('install calls pluginService with spec', async () => {
    mockFetchOk({ installed: true, name: 'new-plugin' });
    const result = (await callAsync(managePluginsTool, {
      action: 'install',
      spec: 'owner/repo',
    })) as { installed: boolean; name: string };
    expect(result.installed).toBe(true);
    expect(result.name).toBe('new-plugin');
  });

  it('uninstall requires name', async () => {
    const result = (await callAsync(managePluginsTool, { action: 'uninstall' })) as {
      error: string;
    };
    expect(result.error).toContain('name');
  });

  it('update requires name', async () => {
    const result = (await callAsync(managePluginsTool, { action: 'update' })) as {
      error: string;
    };
    expect(result.error).toContain('name');
  });

  it('add_marketplace requires spec', async () => {
    const result = (await callAsync(managePluginsTool, { action: 'add_marketplace' })) as {
      error: string;
    };
    expect(result.error).toContain('spec');
  });

  it('remove_marketplace requires name', async () => {
    const result = (await callAsync(managePluginsTool, { action: 'remove_marketplace' })) as {
      error: string;
    };
    expect(result.error).toContain('name');
  });

  it('marketplaces returns registered marketplaces', async () => {
    mockFetchOk({ marketplaces: [
      { name: 'awesome-copilot', source: 'built-in', isBuiltIn: true, pluginCount: 42 },
    ] });
    const result = (await callAsync(managePluginsTool, { action: 'marketplaces' })) as {
      marketplaces: { name: string; isBuiltIn: boolean }[];
      count: number;
    };
    expect(result.count).toBe(1);
    expect(result.marketplaces[0].name).toBe('awesome-copilot');
    expect(result.marketplaces[0].isBuiltIn).toBe(true);
  });

  it('unknown action returns error', async () => {
    const result = (await callAsync(managePluginsTool, { action: 'purge' })) as {
      error: string;
    };
    expect(result.error).toContain('Unknown action');
  });

  it('network failure returns error JSON', async () => {
    mockFetchReject('fetch failed');
    const result = (await callAsync(managePluginsTool, { action: 'list' })) as {
      error: string;
    };
    expect(result.error).toContain('fetch failed');
  });

  it('non-ok response returns error JSON', async () => {
    fetchMock.mockResolvedValueOnce({
      ok: false,
      status: 500,
      statusText: 'Internal Server Error',
      text: () => Promise.resolve('server error'),
    });
    const result = (await callAsync(managePluginsTool, { action: 'list' })) as {
      error: string;
    };
    expect(result.error).toContain('500');
  });
});

// ─── manage_memory handler ──────────────────────────────────────────────────

describe('manage_memory handler', () => {
  it('list returns empty initially', () => {
    const result = callSync(manageMemoryTool, { action: 'list' }) as {
      count: number;
      memories: unknown[];
    };
    expect(result.count).toBe(0);
    expect(result.memories).toEqual([]);
  });

  it('save → list → remove roundtrip', () => {
    const saved = callSync(manageMemoryTool, {
      action: 'save',
      content: 'User prefers blue headers',
      category: 'preference',
    }) as { saved: boolean; id: string };
    expect(saved.saved).toBe(true);
    expect(saved.id).toBeTruthy();

    const listed = callSync(manageMemoryTool, { action: 'list' }) as {
      count: number;
      memories: { id: string; content: string; category: string }[];
    };
    expect(listed.count).toBe(1);
    expect(listed.memories[0].content).toBe('User prefers blue headers');
    expect(listed.memories[0].category).toBe('preference');

    const removed = callSync(manageMemoryTool, {
      action: 'remove',
      id: saved.id,
    }) as { removed: boolean };
    expect(removed.removed).toBe(true);

    const afterRemove = callSync(manageMemoryTool, { action: 'list' }) as {
      count: number;
    };
    expect(afterRemove.count).toBe(0);
  });

  it('save requires content', () => {
    const result = callSync(manageMemoryTool, { action: 'save' }) as { error: string };
    expect(result.error).toContain('content');
  });

  it('search requires query', () => {
    const result = callSync(manageMemoryTool, { action: 'search' }) as { error: string };
    expect(result.error).toContain('query');
  });

  it('search finds matching memories', () => {
    callSync(manageMemoryTool, { action: 'save', content: 'Team name is Contoso' });
    callSync(manageMemoryTool, { action: 'save', content: 'Prefers dark mode' });

    const result = callSync(manageMemoryTool, {
      action: 'search',
      query: 'Contoso',
    }) as { count: number; results: { content: string }[] };
    expect(result.count).toBe(1);
    expect(result.results[0].content).toContain('Contoso');
  });

  it('remove requires id', () => {
    const result = callSync(manageMemoryTool, { action: 'remove' }) as { error: string };
    expect(result.error).toContain('id');
  });

  it('clear deletes all memories', () => {
    callSync(manageMemoryTool, { action: 'save', content: 'Fact 1' });
    callSync(manageMemoryTool, { action: 'save', content: 'Fact 2' });

    const cleared = callSync(manageMemoryTool, { action: 'clear' }) as { cleared: boolean };
    expect(cleared.cleared).toBe(true);

    const listed = callSync(manageMemoryTool, { action: 'list' }) as { count: number };
    expect(listed.count).toBe(0);
  });

  it('unknown action returns error', () => {
    const result = callSync(manageMemoryTool, { action: 'purge' }) as { error: string };
    expect(result.error).toContain('Unknown action');
  });
});

// ─── getToolsForHost includes management tools ──────────────────────────────

describe('management tools wiring', () => {
  it('getToolsForHost("excel") includes management tools', async () => {
    const { getToolsForHost } = await import('@/tools');
    const tools = getToolsForHost('excel');
    const names = tools.map(t => t.name);
    expect(names).toContain('manage_plugins');
    expect(names).toContain('manage_memory');
  });

  it('getToolsForHost("powerpoint") includes management tools', async () => {
    const { getToolsForHost } = await import('@/tools');
    const tools = getToolsForHost('powerpoint');
    const names = tools.map(t => t.name);
    expect(names).toContain('manage_plugins');
    expect(names).toContain('manage_memory');
  });

  it('getToolsForHost("word") includes management tools', async () => {
    const { getToolsForHost } = await import('@/tools');
    const tools = getToolsForHost('word');
    const names = tools.map(t => t.name);
    expect(names).toContain('manage_plugins');
    expect(names).toContain('manage_memory');
  });
});
