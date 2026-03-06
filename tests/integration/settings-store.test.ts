import { describe, it, expect, beforeEach } from 'vitest';
import { useSettingsStore } from '@/stores/settingsStore';
import type { CopilotModel } from '@/types';
import type { AgentSkill } from '@/types/skill';
import type { AgentConfig } from '@/types/agent';
import type { McpServerConfig } from '@/types/mcp';

const TEST_MODELS: CopilotModel[] = [
  { id: 'claude-sonnet-4.6', name: 'Claude Sonnet 4.6', provider: 'Anthropic' },
  { id: 'gpt-4.1', name: 'GPT-4.1', provider: 'OpenAI' },
];

beforeEach(() => {
  useSettingsStore.getState().reset();
});

// ─── Model management ───

describe('settingsStore — model', () => {
  it('starts with the default model (claude-sonnet-4.6)', () => {
    expect(useSettingsStore.getState().activeModel).toBe('claude-sonnet-4.6');
  });

  it('setActiveModel accepts any model when availableModels is null', () => {
    useSettingsStore.getState().setActiveModel('any-model-id');
    expect(useSettingsStore.getState().activeModel).toBe('any-model-id');
  });

  it('setActiveModel validates against availableModels when set', () => {
    useSettingsStore.getState().setAvailableModels(TEST_MODELS);
    useSettingsStore.getState().setActiveModel('unknown-model-xyz');
    expect(useSettingsStore.getState().activeModel).toBe('claude-sonnet-4.6');
  });

  it('setActiveModel accepts a valid model ID from availableModels', () => {
    useSettingsStore.getState().setAvailableModels(TEST_MODELS);
    const id = TEST_MODELS[1].id;
    useSettingsStore.getState().setActiveModel(id);
    expect(useSettingsStore.getState().activeModel).toBe(id);
  });

  it('reset restores the default model', () => {
    useSettingsStore.getState().setActiveModel('gpt-4.1');
    useSettingsStore.getState().reset();
    expect(useSettingsStore.getState().activeModel).toBe('claude-sonnet-4.6');
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
});

// ─── Skill management ───

describe('settingsStore — skills', () => {
  it('starts with no disabled skills', () => {
    expect(useSettingsStore.getState().disabledSkillNames).toEqual([]);
  });

  it('toggleSkill disables an enabled skill', () => {
    useSettingsStore.getState().toggleSkill('excel');
    expect(useSettingsStore.getState().disabledSkillNames).toContain('excel');
  });

  it('toggleSkill re-enables a disabled skill', () => {
    useSettingsStore.getState().toggleSkill('excel');
    useSettingsStore.getState().toggleSkill('excel');
    expect(useSettingsStore.getState().disabledSkillNames).not.toContain('excel');
  });

  it('isSkillEnabled returns true for enabled skills', () => {
    expect(useSettingsStore.getState().isSkillEnabled('excel')).toBe(true);
  });

  it('isSkillEnabled returns false for disabled skills', () => {
    useSettingsStore.getState().toggleSkill('excel');
    expect(useSettingsStore.getState().isSkillEnabled('excel')).toBe(false);
  });

  it('reset clears disabled skills', () => {
    useSettingsStore.getState().toggleSkill('excel');
    useSettingsStore.getState().reset();
    expect(useSettingsStore.getState().disabledSkillNames).toEqual([]);
  });
});

// ─── MCP server management ───

describe('settingsStore — mcp servers', () => {
  it('starts with no disabled MCP servers', () => {
    expect(useSettingsStore.getState().disabledMcpServerNames).toEqual([]);
  });

  it('toggleMcpServer disables an enabled server', () => {
    useSettingsStore.getState().toggleMcpServer('workiq');
    expect(useSettingsStore.getState().disabledMcpServerNames).toContain('workiq');
  });

  it('toggleMcpServer re-enables a disabled server', () => {
    useSettingsStore.getState().toggleMcpServer('workiq');
    useSettingsStore.getState().toggleMcpServer('workiq');
    expect(useSettingsStore.getState().disabledMcpServerNames).not.toContain('workiq');
  });

  it('isMcpServerEnabled returns true for enabled servers', () => {
    expect(useSettingsStore.getState().isMcpServerEnabled('workiq')).toBe(true);
  });

  it('isMcpServerEnabled returns false for disabled servers', () => {
    useSettingsStore.getState().toggleMcpServer('workiq');
    expect(useSettingsStore.getState().isMcpServerEnabled('workiq')).toBe(false);
  });

  it('reset clears disabled MCP servers', () => {
    useSettingsStore.getState().toggleMcpServer('workiq');
    useSettingsStore.getState().reset();
    expect(useSettingsStore.getState().disabledMcpServerNames).toEqual([]);
  });
});

// ─── Imported skills ───

const SAMPLE_SKILL: AgentSkill = {
  metadata: { name: 'test-skill', description: 'A test skill', version: '1.0.0', tags: [], hosts: [] },
  content: 'Do something.',
};

describe('settingsStore — imported skills', () => {
  it('starts with no imported skills', () => {
    expect(useSettingsStore.getState().importedSkills).toEqual([]);
  });

  it('addImportedSkill adds a skill', () => {
    useSettingsStore.getState().addImportedSkill(SAMPLE_SKILL);
    expect(useSettingsStore.getState().importedSkills).toHaveLength(1);
    expect(useSettingsStore.getState().importedSkills[0].metadata.name).toBe('test-skill');
  });

  it('addImportedSkill replaces an existing skill with the same name', () => {
    useSettingsStore.getState().addImportedSkill(SAMPLE_SKILL);
    useSettingsStore.getState().addImportedSkill({ ...SAMPLE_SKILL, content: 'Updated.' });
    expect(useSettingsStore.getState().importedSkills).toHaveLength(1);
    expect(useSettingsStore.getState().importedSkills[0].content).toBe('Updated.');
  });

  it('removeImportedSkill removes a skill by name', () => {
    useSettingsStore.getState().addImportedSkill(SAMPLE_SKILL);
    useSettingsStore.getState().removeImportedSkill('test-skill');
    expect(useSettingsStore.getState().importedSkills).toHaveLength(0);
  });

  it('removeImportedSkill is a no-op for unknown names', () => {
    useSettingsStore.getState().addImportedSkill(SAMPLE_SKILL);
    useSettingsStore.getState().removeImportedSkill('no-such-skill');
    expect(useSettingsStore.getState().importedSkills).toHaveLength(1);
  });

  it('reset clears imported skills', () => {
    useSettingsStore.getState().addImportedSkill(SAMPLE_SKILL);
    useSettingsStore.getState().reset();
    expect(useSettingsStore.getState().importedSkills).toEqual([]);
  });
});

// ─── Imported agents ───

const SAMPLE_AGENT: AgentConfig = {
  metadata: {
    name: 'test-agent',
    description: 'A test agent',
    version: '1.0.0',
    hosts: ['excel'],
    defaultForHosts: [],
  },
  instructions: 'Do something.',
};

describe('settingsStore — imported agents', () => {
  it('starts with no imported agents', () => {
    expect(useSettingsStore.getState().importedAgents).toEqual([]);
  });

  it('addImportedAgent adds an agent', () => {
    useSettingsStore.getState().addImportedAgent(SAMPLE_AGENT);
    expect(useSettingsStore.getState().importedAgents).toHaveLength(1);
    expect(useSettingsStore.getState().importedAgents[0].metadata.name).toBe('test-agent');
  });

  it('addImportedAgent replaces an existing agent with the same name', () => {
    useSettingsStore.getState().addImportedAgent(SAMPLE_AGENT);
    useSettingsStore.getState().addImportedAgent({ ...SAMPLE_AGENT, instructions: 'Updated.' });
    expect(useSettingsStore.getState().importedAgents).toHaveLength(1);
    expect(useSettingsStore.getState().importedAgents[0].instructions).toBe('Updated.');
  });

  it('removeImportedAgent removes an agent by name', () => {
    useSettingsStore.getState().addImportedAgent(SAMPLE_AGENT);
    useSettingsStore.getState().removeImportedAgent('test-agent');
    expect(useSettingsStore.getState().importedAgents).toHaveLength(0);
  });

  it('reset clears imported agents', () => {
    useSettingsStore.getState().addImportedAgent(SAMPLE_AGENT);
    useSettingsStore.getState().reset();
    expect(useSettingsStore.getState().importedAgents).toEqual([]);
  });
});

// ─── Imported MCP servers ───

const SAMPLE_MCP: McpServerConfig = {
  name: 'test-mcp',
  description: 'A test server',
  transport: 'http',
  url: 'https://example.com/mcp',
};

describe('settingsStore — imported MCP servers', () => {
  it('starts with no imported MCP servers', () => {
    expect(useSettingsStore.getState().importedMcpServers).toEqual([]);
  });

  it('addImportedMcpServer adds a server', () => {
    useSettingsStore.getState().addImportedMcpServer(SAMPLE_MCP);
    expect(useSettingsStore.getState().importedMcpServers).toHaveLength(1);
    expect(useSettingsStore.getState().importedMcpServers[0].name).toBe('test-mcp');
  });

  it('addImportedMcpServer replaces an existing server with the same name', () => {
    useSettingsStore.getState().addImportedMcpServer(SAMPLE_MCP);
    useSettingsStore.getState().addImportedMcpServer({ ...SAMPLE_MCP, url: 'https://new.com/mcp' });
    expect(useSettingsStore.getState().importedMcpServers).toHaveLength(1);
    expect(useSettingsStore.getState().importedMcpServers[0].url).toBe('https://new.com/mcp');
  });

  it('removeImportedMcpServer removes a server by name', () => {
    useSettingsStore.getState().addImportedMcpServer(SAMPLE_MCP);
    useSettingsStore.getState().removeImportedMcpServer('test-mcp');
    expect(useSettingsStore.getState().importedMcpServers).toHaveLength(0);
  });

  it('reset clears imported MCP servers', () => {
    useSettingsStore.getState().addImportedMcpServer(SAMPLE_MCP);
    useSettingsStore.getState().reset();
    expect(useSettingsStore.getState().importedMcpServers).toEqual([]);
  });
});
