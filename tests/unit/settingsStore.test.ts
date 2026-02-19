import { describe, it, expect, beforeEach } from 'vitest';
import { useSettingsStore } from '@/stores/settingsStore';
import { COPILOT_MODELS } from '@/types';
import type { AgentConfig, AgentSkill, McpServerConfig } from '@/types';

beforeEach(() => {
  useSettingsStore.getState().reset();
});

// ─── Model management ───

describe('settingsStore — model', () => {
  it('starts with the default model (claude-sonnet-4.5)', () => {
    expect(useSettingsStore.getState().activeModel).toBe('claude-sonnet-4.5');
  });

  it('setActiveModel accepts a valid COPILOT_MODELS ID', () => {
    const id = COPILOT_MODELS[1].id;
    useSettingsStore.getState().setActiveModel(id);
    expect(useSettingsStore.getState().activeModel).toBe(id);
  });

  it('setActiveModel ignores unknown model IDs', () => {
    useSettingsStore.getState().setActiveModel('unknown-model-xyz');
    expect(useSettingsStore.getState().activeModel).toBe('claude-sonnet-4.5');
  });

  it('reset restores the default model', () => {
    useSettingsStore.getState().setActiveModel('gpt-4.1');
    useSettingsStore.getState().reset();
    expect(useSettingsStore.getState().activeModel).toBe('claude-sonnet-4.5');
  });
});

// ─── Skill management ───

describe('settingsStore — skills', () => {
  it('starts with all skills enabled (null)', () => {
    expect(useSettingsStore.getState().activeSkillNames).toBeNull();
  });

  it('toggleSkill on null materializes list minus toggled skill', () => {
    useSettingsStore.getState().toggleSkill('xa2');
    const names = useSettingsStore.getState().activeSkillNames;
    expect(Array.isArray(names)).toBe(true);
    expect(names).not.toContain('xa2');
  });

  it('toggleSkill adds a skill back after removal', () => {
    useSettingsStore.getState().toggleSkill('xa2'); // remove
    useSettingsStore.getState().toggleSkill('xa2'); // re-add
    expect(useSettingsStore.getState().activeSkillNames).toContain('xa2');
  });

  it('toggleSkill handles multiple skills independently', () => {
    useSettingsStore.getState().setActiveSkills(['xa2', 'another']);
    useSettingsStore.getState().toggleSkill('xa2');
    expect(useSettingsStore.getState().activeSkillNames).toEqual(['another']);
  });

  it('setActiveSkills replaces the full list', () => {
    useSettingsStore.getState().setActiveSkills(['a', 'b']);
    expect(useSettingsStore.getState().activeSkillNames).toEqual(['a', 'b']);
  });

  it('setActiveSkills(null) restores all-on default', () => {
    useSettingsStore.getState().setActiveSkills(['a']);
    useSettingsStore.getState().setActiveSkills(null);
    expect(useSettingsStore.getState().activeSkillNames).toBeNull();
  });

  it('getActiveSkillNames returns null when all on', () => {
    expect(useSettingsStore.getState().getActiveSkillNames()).toBeNull();
  });

  it('reset restores null (all skills on)', () => {
    useSettingsStore.getState().setActiveSkills(['xa2']);
    useSettingsStore.getState().reset();
    expect(useSettingsStore.getState().activeSkillNames).toBeNull();
  });

  it('importSkills stores imported skills', () => {
    const skill: AgentSkill = {
      metadata: {
        name: 'Imported Skill',
        description: 'Imported from zip.',
        version: '1.0.0',
        tags: [],
      },
      content: 'Skill body',
    };

    useSettingsStore.getState().importSkills([skill]);

    expect(useSettingsStore.getState().importedSkills).toHaveLength(1);
    expect(useSettingsStore.getState().importedSkills[0].metadata.name).toBe('Imported Skill');
  });

  it('removeImportedSkill removes imported skill and prunes active list', () => {
    const skill: AgentSkill = {
      metadata: {
        name: 'Imported Skill',
        description: 'Imported from zip.',
        version: '1.0.0',
        tags: [],
      },
      content: 'Skill body',
    };

    useSettingsStore.getState().importSkills([skill]);
    useSettingsStore.getState().setActiveSkills(['Imported Skill']);
    useSettingsStore.getState().removeImportedSkill('Imported Skill');

    expect(useSettingsStore.getState().importedSkills).toEqual([]);
    expect(useSettingsStore.getState().activeSkillNames).toEqual([]);
  });
});

// ─── Agent management ───

describe('settingsStore — agents', () => {
  it('starts with "Excel" as the default active agent', () => {
    expect(useSettingsStore.getState().activeAgentId).toBe('Excel');
  });

  it('setActiveAgent changes the active agent', () => {
    useSettingsStore.getState().setActiveAgent('Excel');
    expect(useSettingsStore.getState().activeAgentId).toBe('Excel');
  });

  it('setActiveAgent ignores invalid agent names', () => {
    useSettingsStore.getState().setActiveAgent('NonExistentAgent');
    expect(useSettingsStore.getState().activeAgentId).toBe('Excel');
  });

  it('getActiveAgent returns the current agent id', () => {
    expect(useSettingsStore.getState().getActiveAgent()).toBe('Excel');
  });

  it('reset restores the default agent', () => {
    useSettingsStore.getState().reset();
    expect(useSettingsStore.getState().activeAgentId).toBe('Excel');
  });

  it('importAgents stores imported agents', () => {
    const agent: AgentConfig = {
      metadata: {
        name: 'Imported Agent',
        description: 'Imported from zip.',
        version: '1.0.0',
        hosts: ['excel'],
        defaultForHosts: [],
      },
      instructions: 'Do imported work.',
    };

    useSettingsStore.getState().importAgents([agent]);

    expect(useSettingsStore.getState().importedAgents).toHaveLength(1);
    expect(useSettingsStore.getState().importedAgents[0].metadata.name).toBe('Imported Agent');
  });

  it('removeImportedAgent resets active agent when removed agent was selected', () => {
    const agent: AgentConfig = {
      metadata: {
        name: 'Imported Agent',
        description: 'Imported from zip.',
        version: '1.0.0',
        hosts: ['excel'],
        defaultForHosts: [],
      },
      instructions: 'Do imported work.',
    };

    useSettingsStore.getState().importAgents([agent]);
    useSettingsStore.getState().setActiveAgent('Imported Agent');
    useSettingsStore.getState().removeImportedAgent('Imported Agent');

    expect(useSettingsStore.getState().importedAgents).toEqual([]);
    expect(useSettingsStore.getState().activeAgentId).toBe('Excel');
  });
});

// ─── MCP server management ───

describe('settingsStore — MCP servers', () => {
  const server1: McpServerConfig = { name: 'srv1', url: 'https://s1.com/mcp', transport: 'http' };
  const server2: McpServerConfig = { name: 'srv2', url: 'https://s2.com/mcp', transport: 'sse' };

  it('imports MCP servers', () => {
    useSettingsStore.getState().importMcpServers([server1, server2]);
    expect(useSettingsStore.getState().importedMcpServers).toHaveLength(2);
  });

  it('renames duplicate server names on import', () => {
    useSettingsStore.getState().importMcpServers([server1]);
    useSettingsStore.getState().importMcpServers([server1]);
    const names = useSettingsStore.getState().importedMcpServers.map(s => s.name);
    expect(names[0]).toBe('srv1');
    expect(names[1]).not.toBe('srv1');
  });

  it('removes a MCP server by name', () => {
    useSettingsStore.getState().importMcpServers([server1, server2]);
    useSettingsStore.getState().removeMcpServer('srv1');
    expect(useSettingsStore.getState().importedMcpServers.map(s => s.name)).toEqual(['srv2']);
  });

  it('removes from activeMcpServerNames on server removal', () => {
    useSettingsStore.getState().importMcpServers([server1, server2]);
    useSettingsStore.setState({ activeMcpServerNames: ['srv1', 'srv2'] });
    useSettingsStore.getState().removeMcpServer('srv1');
    expect(useSettingsStore.getState().activeMcpServerNames).toEqual(['srv2']);
  });

  it('activeMcpServerNames is null (all on) by default', () => {
    useSettingsStore.getState().importMcpServers([server1]);
    expect(useSettingsStore.getState().activeMcpServerNames).toBeNull();
  });

  it('toggleMcpServer off materializes full list minus toggled server', () => {
    useSettingsStore.getState().importMcpServers([server1, server2]);
    useSettingsStore.getState().toggleMcpServer('srv1');
    expect(useSettingsStore.getState().activeMcpServerNames).toEqual(['srv2']);
  });

  it('toggleMcpServer on adds server back to active list', () => {
    useSettingsStore.getState().importMcpServers([server1, server2]);
    useSettingsStore.setState({ activeMcpServerNames: ['srv2'] });
    useSettingsStore.getState().toggleMcpServer('srv1');
    expect(useSettingsStore.getState().activeMcpServerNames).toContain('srv1');
  });

  it('reset clears imported MCP servers', () => {
    useSettingsStore.getState().importMcpServers([server1]);
    useSettingsStore.getState().reset();
    expect(useSettingsStore.getState().importedMcpServers).toEqual([]);
    expect(useSettingsStore.getState().activeMcpServerNames).toBeNull();
  });
});
