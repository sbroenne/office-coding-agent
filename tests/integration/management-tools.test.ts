import { describe, it, expect, beforeEach } from 'vitest';
import Ajv from 'ajv';
import { manageMemoryTool, managementTools } from '@/tools/management';
import { useMemoryStore } from '@/stores/memoryStore';

const ajv = new Ajv({ allErrors: true });

function validate(schema: unknown, data: unknown): boolean {
  return !!ajv.compile(schema as object)(data);
}

function callSync(tool: { handler?: unknown }, args: Record<string, unknown>): unknown {
  const handler = (tool as { handler: (a: unknown) => string }).handler;
  return JSON.parse(handler(args));
}

describe('management tools', () => {
  beforeEach(() => {
    useMemoryStore.getState().clearMemories();
  });

  it('only includes memory management', () => {
    expect(managementTools.map(tool => tool.name)).toEqual(['manage_memory']);
  });

  it('validates manage_memory schema', () => {
    expect(validate(manageMemoryTool.parameters, { action: 'save', content: 'Prefer concise replies' })).toBe(true);
    expect(validate(manageMemoryTool.parameters, { action: 'list' })).toBe(true);
    expect(validate(manageMemoryTool.parameters, { action: 'search', query: 'concise' })).toBe(true);
    expect(validate(manageMemoryTool.parameters, { action: 'remove', id: 'mem-1' })).toBe(true);
    expect(validate(manageMemoryTool.parameters, { action: 'clear' })).toBe(true);
    expect(validate(manageMemoryTool.parameters, { action: 'install' })).toBe(false);
  });

  it('saves, lists, searches, removes, and clears memories', () => {
    const saved = callSync(manageMemoryTool, {
      action: 'save',
      content: 'Prefer concise replies',
      category: 'preference',
    }) as { id: string; saved: boolean };
    expect(saved.saved).toBe(true);

    const listed = callSync(manageMemoryTool, { action: 'list' }) as { count: number };
    expect(listed.count).toBe(1);

    const searched = callSync(manageMemoryTool, { action: 'search', query: 'concise' }) as {
      count: number;
    };
    expect(searched.count).toBe(1);

    expect(callSync(manageMemoryTool, { action: 'remove', id: saved.id })).toEqual({
      removed: true,
      id: saved.id,
    });

    callSync(manageMemoryTool, { action: 'save', content: 'Temporary fact' });
    expect(callSync(manageMemoryTool, { action: 'clear' })).toEqual({
      cleared: true,
      message: 'All memories deleted.',
    });
  });
});
