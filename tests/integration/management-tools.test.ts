/**
 * Integration tests for management tools (manage_skills, manage_agents, manage_mcp_servers).
 *
 * These exercise real Zustand store operations through the tool handlers — no mocks.
 */

import { describe, it, expect, beforeEach } from 'vitest';
import Ajv from 'ajv';
import { useSettingsStore } from '@/stores';
import { setImportedSkills } from '@/services/skills/skillService';
import { manageSkillsTool, manageAgentsTool, manageMcpServersTool } from '@/tools/management';

const ajv = new Ajv({ allErrors: true });

function validate(schema: unknown, data: unknown): boolean {
  return !!ajv.compile(schema as object)(data);
}

/** Call a tool handler and parse the JSON result */
function call(tool: { handler?: unknown }, args: Record<string, unknown>): unknown {
  const handler = (tool as { handler: (a: unknown, i: unknown) => unknown }).handler;
  const raw = handler(args, {});
  return JSON.parse(raw as string);
}

beforeEach(() => {
  useSettingsStore.getState().reset();
  setImportedSkills([]);
});

// ─── Schema validation ──────────────────────────────────────────────────────

describe('management tool schemas', () => {
  it('manage_skills accepts list action', () => {
    expect(validate(manageSkillsTool.parameters, { action: 'list' })).toBe(true);
  });

  it('manage_skills accepts install with all params', () => {
    expect(
      validate(manageSkillsTool.parameters, {
        action: 'install',
        name: 'test',
        description: 'desc',
        version: '1.0.0',
        hosts: ['excel'],
        tags: ['tag1'],
        content: 'body',
      })
    ).toBe(true);
  });

  it('manage_skills rejects missing action', () => {
    expect(validate(manageSkillsTool.parameters, { name: 'test' })).toBe(false);
  });

  it('manage_agents accepts list action', () => {
    expect(validate(manageAgentsTool.parameters, { action: 'list' })).toBe(true);
  });

  it('manage_agents accepts install with structured params', () => {
    expect(
      validate(manageAgentsTool.parameters, {
        action: 'install',
        name: 'My Agent',
        hosts: ['excel'],
        instructions: 'Do something.',
      })
    ).toBe(true);
  });

  it('manage_mcp_servers accepts list action', () => {
    expect(validate(manageMcpServersTool.parameters, { action: 'list' })).toBe(true);
  });

  it('manage_mcp_servers accepts install', () => {
    expect(
      validate(manageMcpServersTool.parameters, {
        action: 'install',
        name: 'my-server',
        url: 'https://example.com/mcp',
        transport: 'http',
      })
    ).toBe(true);
  });

  it('manage_mcp_servers rejects invalid transport', () => {
    expect(
      validate(manageMcpServersTool.parameters, {
        action: 'install',
        name: 'server',
        url: 'https://example.com',
        transport: 'grpc',
      })
    ).toBe(false);
  });
});

// ─── manage_skills handler ──────────────────────────────────────────────────

describe('manage_skills handler', () => {
  it('list returns bundled skills', () => {
    const result = call(manageSkillsTool, { action: 'list' }) as {
      skills: { name: string; active: boolean }[];
      count: number;
    };
    expect(result.count).toBeGreaterThan(0);
    expect(result.skills.some(s => s.name === 'excel')).toBe(true);
  });

  it('install → list → remove roundtrip', () => {
    // Install
    const installed = call(manageSkillsTool, {
      action: 'install',
      name: 'New Skill',
      content: '## Guide\nUse this for testing.',
      hosts: ['excel'],
    }) as { installed: boolean; name: string };
    expect(installed.installed).toBe(true);
    expect(installed.name).toBe('New Skill');

    // List — should include the new skill
    const listed = call(manageSkillsTool, { action: 'list' }) as {
      skills: { name: string }[];
    };
    expect(listed.skills.some(s => s.name === 'New Skill')).toBe(true);

    // Remove
    const removed = call(manageSkillsTool, {
      action: 'remove',
      name: 'New Skill',
    }) as { removed: boolean };
    expect(removed.removed).toBe(true);

    // List again — should be gone
    const afterRemove = call(manageSkillsTool, { action: 'list' }) as {
      skills: { name: string }[];
    };
    expect(afterRemove.skills.some(s => s.name === 'New Skill')).toBe(false);
  });

  it('install requires name', () => {
    const result = call(manageSkillsTool, {
      action: 'install',
      content: 'body',
    }) as { error: string };
    expect(result.error).toContain('name');
  });

  it('install requires content', () => {
    const result = call(manageSkillsTool, {
      action: 'install',
      name: 'No Content',
    }) as { error: string };
    expect(result.error).toContain('content');
  });

  it('toggle changes active state', () => {
    // Install → toggle OFF → check → toggle ON → check
    call(manageSkillsTool, {
      action: 'install',
      name: 'ToggleSkill',
      content: 'body',
    });

    const toggled = call(manageSkillsTool, {
      action: 'toggle',
      name: 'ToggleSkill',
    }) as { toggled: boolean; name: string; active: boolean };
    expect(toggled.toggled).toBe(true);
    expect(toggled.name).toBe('ToggleSkill');
    expect(toggled.active).toBe(false);
  });

  it('toggle requires name', () => {
    const result = call(manageSkillsTool, { action: 'toggle' }) as { error: string };
    expect(result.error).toContain('name');
  });

  it('remove requires name', () => {
    const result = call(manageSkillsTool, { action: 'remove' }) as { error: string };
    expect(result.error).toContain('name');
  });

  it('unknown action returns error', () => {
    const result = call(manageSkillsTool, { action: 'delete' }) as { error: string };
    expect(result.error).toContain('Unknown action');
  });
});

// ─── manage_agents handler ──────────────────────────────────────────────────

describe('manage_agents handler', () => {
  it('list returns bundled agents', () => {
    const result = call(manageAgentsTool, { action: 'list' }) as {
      agents: { name: string; active: boolean }[];
      count: number;
    };
    expect(result.count).toBeGreaterThan(0);
    expect(result.agents.some(a => a.name === 'Excel')).toBe(true);
  });

  it('install → list → remove roundtrip', () => {
    const installed = call(manageAgentsTool, {
      action: 'install',
      name: 'Test Agent',
      hosts: ['excel'],
      instructions: 'Be helpful.',
    }) as { installed: boolean; name: string };
    expect(installed.installed).toBe(true);

    const listed = call(manageAgentsTool, { action: 'list' }) as {
      agents: { name: string }[];
    };
    expect(listed.agents.some(a => a.name === 'Test Agent')).toBe(true);

    const removed = call(manageAgentsTool, {
      action: 'remove',
      name: 'Test Agent',
    }) as { removed: boolean };
    expect(removed.removed).toBe(true);

    const afterRemove = call(manageAgentsTool, { action: 'list' }) as {
      agents: { name: string }[];
    };
    expect(afterRemove.agents.some(a => a.name === 'Test Agent')).toBe(false);
  });

  it('install requires name', () => {
    const result = call(manageAgentsTool, {
      action: 'install',
      hosts: ['excel'],
      instructions: 'body',
    }) as { error: string };
    expect(result.error).toContain('name');
  });

  it('install requires hosts', () => {
    const result = call(manageAgentsTool, {
      action: 'install',
      name: 'Agent',
      instructions: 'body',
    }) as { error: string };
    expect(result.error).toContain('hosts');
  });

  it('install requires instructions', () => {
    const result = call(manageAgentsTool, {
      action: 'install',
      name: 'Agent',
      hosts: ['excel'],
    }) as { error: string };
    expect(result.error).toContain('instructions');
  });

  it('set_active changes active agent', () => {
    // Install a new agent and set it active
    call(manageAgentsTool, {
      action: 'install',
      name: 'Custom Agent',
      hosts: ['excel'],
      instructions: 'Be awesome.',
    });

    const result = call(manageAgentsTool, {
      action: 'set_active',
      name: 'Custom Agent',
    }) as { activeAgentId: string; message: string };
    expect(result.activeAgentId).toBe('Custom Agent');
    expect(result.message).toContain('Custom Agent');
  });

  it('set_active requires name', () => {
    const result = call(manageAgentsTool, { action: 'set_active' }) as { error: string };
    expect(result.error).toContain('name');
  });

  it('remove requires name', () => {
    const result = call(manageAgentsTool, { action: 'remove' }) as { error: string };
    expect(result.error).toContain('name');
  });

  it('unknown action returns error', () => {
    const result = call(manageAgentsTool, { action: 'update' }) as { error: string };
    expect(result.error).toContain('Unknown action');
  });
});

// ─── manage_mcp_servers handler ─────────────────────────────────────────────

describe('manage_mcp_servers handler', () => {
  it('list returns empty initially', () => {
    const result = call(manageMcpServersTool, { action: 'list' }) as {
      servers: unknown[];
      count: number;
    };
    expect(result.count).toBe(0);
    expect(result.servers).toEqual([]);
  });

  it('install → list → remove roundtrip', () => {
    const installed = call(manageMcpServersTool, {
      action: 'install',
      name: 'my-server',
      url: 'https://example.com/mcp',
      transport: 'http',
      description: 'Test MCP server',
    }) as { installed: boolean; name: string };
    expect(installed.installed).toBe(true);

    const listed = call(manageMcpServersTool, { action: 'list' }) as {
      servers: { name: string; url: string; transport: string }[];
    };
    expect(listed.servers).toHaveLength(1);
    expect(listed.servers[0].name).toBe('my-server');
    expect(listed.servers[0].url).toBe('https://example.com/mcp');

    const removed = call(manageMcpServersTool, {
      action: 'remove',
      name: 'my-server',
    }) as { removed: boolean };
    expect(removed.removed).toBe(true);

    const afterRemove = call(manageMcpServersTool, { action: 'list' }) as {
      servers: unknown[];
    };
    expect(afterRemove.servers).toHaveLength(0);
  });

  it('install requires name', () => {
    const result = call(manageMcpServersTool, {
      action: 'install',
      url: 'https://example.com',
    }) as { error: string };
    expect(result.error).toContain('name');
  });

  it('install requires url', () => {
    const result = call(manageMcpServersTool, {
      action: 'install',
      name: 'server',
    }) as { error: string };
    expect(result.error).toContain('url');
  });

  it('toggle changes active state', () => {
    call(manageMcpServersTool, {
      action: 'install',
      name: 'toggle-server',
      url: 'https://example.com/mcp',
    });

    const toggled = call(manageMcpServersTool, {
      action: 'toggle',
      name: 'toggle-server',
    }) as { toggled: boolean; name: string; active: boolean };
    expect(toggled.toggled).toBe(true);
    expect(toggled.active).toBe(false);
  });

  it('toggle requires name', () => {
    const result = call(manageMcpServersTool, { action: 'toggle' }) as { error: string };
    expect(result.error).toContain('name');
  });

  it('remove requires name', () => {
    const result = call(manageMcpServersTool, { action: 'remove' }) as { error: string };
    expect(result.error).toContain('name');
  });

  it('unknown action returns error', () => {
    const result = call(manageMcpServersTool, { action: 'purge' }) as { error: string };
    expect(result.error).toContain('Unknown action');
  });
});

// ─── getToolsForHost includes management tools ──────────────────────────────

describe('management tools wiring', () => {
  it('getToolsForHost("excel") includes management tools', async () => {
    const { getToolsForHost } = await import('@/tools');
    const tools = getToolsForHost('excel');
    const names = tools.map(t => t.name);
    expect(names).toContain('manage_skills');
    expect(names).toContain('manage_agents');
    expect(names).toContain('manage_mcp_servers');
    expect(names).toContain('web_fetch');
  });

  it('getToolsForHost("powerpoint") includes management tools', async () => {
    const { getToolsForHost } = await import('@/tools');
    const tools = getToolsForHost('powerpoint');
    const names = tools.map(t => t.name);
    expect(names).toContain('manage_skills');
    expect(names).toContain('manage_agents');
    expect(names).toContain('manage_mcp_servers');
  });

  it('getToolsForHost("word") includes management tools', async () => {
    const { getToolsForHost } = await import('@/tools');
    const tools = getToolsForHost('word');
    const names = tools.map(t => t.name);
    expect(names).toContain('manage_skills');
    expect(names).toContain('manage_agents');
    expect(names).toContain('manage_mcp_servers');
  });
});
