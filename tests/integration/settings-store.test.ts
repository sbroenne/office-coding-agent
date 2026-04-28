import { describe, it, expect, beforeEach } from 'vitest';
import { useSettingsStore } from '@/stores/settingsStore';
import type { CopilotAgent, CopilotModel } from '@/types';

const TEST_MODELS: CopilotModel[] = [
  { id: 'claude-sonnet-4.6', name: 'Claude Sonnet 4.6', provider: 'Anthropic' },
  { id: 'gpt-4.1', name: 'GPT-4.1', provider: 'OpenAI' },
];

const TEST_AGENTS: CopilotAgent[] = [
  { name: 'office-excel', displayName: 'Office Excel', description: 'Excel agent' },
  { name: 'office-word', displayName: 'Office Word', description: 'Word agent' },
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

// ─── CLI agent management ───

describe('settingsStore — CLI agents', () => {
  it('starts with the default agent selected', () => {
    expect(useSettingsStore.getState().activeAgentName).toBeNull();
  });

  it('setActiveAgent accepts any agent when availableAgents is null', () => {
    useSettingsStore.getState().setActiveAgent('office-excel');
    expect(useSettingsStore.getState().activeAgentName).toBe('office-excel');
  });

  it('setActiveAgent validates against availableAgents when set', () => {
    useSettingsStore.getState().setAvailableAgents(TEST_AGENTS);
    useSettingsStore.getState().setActiveAgent('unknown-agent');
    expect(useSettingsStore.getState().activeAgentName).toBeNull();
  });

  it('setActiveAgent accepts a valid CLI agent name from availableAgents', () => {
    useSettingsStore.getState().setAvailableAgents(TEST_AGENTS);
    useSettingsStore.getState().setActiveAgent('office-word');
    expect(useSettingsStore.getState().activeAgentName).toBe('office-word');
  });

  it('setActiveAgent(null) returns to default agent', () => {
    useSettingsStore.getState().setAvailableAgents(TEST_AGENTS);
    useSettingsStore.getState().setActiveAgent('office-excel');
    useSettingsStore.getState().setActiveAgent(null);
    expect(useSettingsStore.getState().activeAgentName).toBeNull();
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
