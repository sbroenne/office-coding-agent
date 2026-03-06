import { describe, it, expect, beforeEach } from 'vitest';
import { useSettingsStore } from '@/stores/settingsStore';
import type { CopilotModel } from '@/types';

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
