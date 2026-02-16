import { describe, it, expect, beforeEach } from 'vitest';
import { useSettingsStore } from '@/stores/settingsStore';
import type { ModelInfo } from '@/types';

/** Reset store to defaults before each test */
beforeEach(() => {
  useSettingsStore.getState().reset();
});

// ─── Endpoint management ───

describe('settingsStore — endpoints', () => {
  const endpoint = {
    displayName: 'Test',
    resourceUrl: 'https://test.openai.azure.com',
    authMethod: 'apiKey' as const,
    apiKey: 'key-1',
  };

  it('adds an endpoint and returns its id', () => {
    const id = useSettingsStore.getState().addEndpoint(endpoint);
    expect(id).toBeTruthy();
    expect(useSettingsStore.getState().endpoints).toHaveLength(1);
    expect(useSettingsStore.getState().endpoints[0].displayName).toBe('Test');
  });

  it('auto-activates the first endpoint', () => {
    const id = useSettingsStore.getState().addEndpoint(endpoint);
    expect(useSettingsStore.getState().activeEndpointId).toBe(id);
  });

  it('deduplicates by normalized URL (trailing slash ignored)', () => {
    const id1 = useSettingsStore.getState().addEndpoint(endpoint);
    const id2 = useSettingsStore.getState().addEndpoint({
      ...endpoint,
      resourceUrl: 'https://test.openai.azure.com/',
      apiKey: 'key-2',
    });

    expect(id1).toBe(id2);
    expect(useSettingsStore.getState().endpoints).toHaveLength(1);
    // Updated API key
    expect(useSettingsStore.getState().endpoints[0].apiKey).toBe('key-2');
  });

  it('removeEndpoint cascades: clears models and resets active', () => {
    const id = useSettingsStore.getState().addEndpoint(endpoint);
    const model: ModelInfo = {
      id: 'gpt-4o',
      name: 'GPT 4o',
      ownedBy: 'system',
      provider: 'OpenAI',
    };
    useSettingsStore.getState().addModel(id, model);
    useSettingsStore.getState().setActiveModel('gpt-4o');

    useSettingsStore.getState().removeEndpoint(id);

    expect(useSettingsStore.getState().endpoints).toHaveLength(0);
    expect(useSettingsStore.getState().activeEndpointId).toBeNull();
    expect(useSettingsStore.getState().activeModelId).toBeNull();
    expect(useSettingsStore.getState().getModelsForEndpoint(id)).toEqual([]);
  });

  it('removeEndpoint picks next endpoint when active one is removed', () => {
    const id1 = useSettingsStore.getState().addEndpoint(endpoint);
    const id2 = useSettingsStore.getState().addEndpoint({
      ...endpoint,
      displayName: 'Second',
      resourceUrl: 'https://second.openai.azure.com',
    });
    useSettingsStore.getState().setActiveEndpoint(id1);

    useSettingsStore.getState().removeEndpoint(id1);

    expect(useSettingsStore.getState().activeEndpointId).toBe(id2);
  });
});

// ─── Model management ───

describe('settingsStore — models', () => {
  const endpoint = {
    displayName: 'Test',
    resourceUrl: 'https://test.openai.azure.com',
    authMethod: 'apiKey' as const,
    apiKey: 'key',
  };

  const modelA: ModelInfo = {
    id: 'gpt-4o',
    name: 'GPT 4o',
    ownedBy: 'system',
    provider: 'OpenAI',
  };
  const modelB: ModelInfo = {
    id: 'claude-3-sonnet',
    name: 'Claude 3 Sonnet',
    ownedBy: 'system',
    provider: 'Anthropic',
  };

  it('addModel auto-selects the first model on the active endpoint', () => {
    const epId = useSettingsStore.getState().addEndpoint(endpoint);
    useSettingsStore.getState().addModel(epId, modelA);

    expect(useSettingsStore.getState().activeModelId).toBe('gpt-4o');
  });

  it('addModel prevents duplicates', () => {
    const epId = useSettingsStore.getState().addEndpoint(endpoint);
    useSettingsStore.getState().addModel(epId, modelA);
    useSettingsStore.getState().addModel(epId, modelA); // duplicate

    expect(useSettingsStore.getState().getModelsForEndpoint(epId)).toHaveLength(1);
  });

  it('removeModel clears activeModelId if it was the active one', () => {
    const epId = useSettingsStore.getState().addEndpoint(endpoint);
    useSettingsStore.getState().addModel(epId, modelA);
    useSettingsStore.getState().addModel(epId, modelB);
    useSettingsStore.getState().setActiveModel('gpt-4o');

    useSettingsStore.getState().removeModel(epId, 'gpt-4o');

    // Falls back to remaining model
    expect(useSettingsStore.getState().activeModelId).toBe('claude-3-sonnet');
  });

  it('setActiveEndpoint auto-selects default model', () => {
    useSettingsStore.getState().addEndpoint(endpoint);
    const ep2 = useSettingsStore.getState().addEndpoint({
      ...endpoint,
      displayName: 'Second',
      resourceUrl: 'https://second.openai.azure.com',
    });

    // Set models on ep2 with an isDefault
    const defaultModel: ModelInfo = { ...modelB, isDefault: true };
    useSettingsStore.getState().addModel(ep2, modelA);
    useSettingsStore.getState().addModel(ep2, defaultModel);

    // Switch to ep2
    useSettingsStore.getState().setActiveEndpoint(ep2);

    expect(useSettingsStore.getState().activeEndpointId).toBe(ep2);
    // Should pick the isDefault model
    expect(useSettingsStore.getState().activeModelId).toBe('claude-3-sonnet');
  });

  it('setModelsForEndpoint replaces the full model list', () => {
    const epId = useSettingsStore.getState().addEndpoint(endpoint);
    useSettingsStore.getState().addModel(epId, modelA);

    useSettingsStore.getState().setModelsForEndpoint(epId, [modelB]);

    expect(useSettingsStore.getState().getModelsForEndpoint(epId)).toEqual([modelB]);
  });
});

// ─── Getters ───

describe('settingsStore — getters', () => {
  it('getActiveEndpoint returns undefined when none active', () => {
    expect(useSettingsStore.getState().getActiveEndpoint()).toBeUndefined();
  });

  it('getActiveModel returns undefined when nothing selected', () => {
    expect(useSettingsStore.getState().getActiveModel()).toBeUndefined();
  });

  it('getModelsForActiveEndpoint returns [] when no endpoint', () => {
    expect(useSettingsStore.getState().getModelsForActiveEndpoint()).toEqual([]);
  });
});

// ─── Reset ───

describe('settingsStore — reset', () => {
  it('restores default state', () => {
    const epId = useSettingsStore.getState().addEndpoint({
      displayName: 'X',
      resourceUrl: 'https://x.openai.azure.com',
      authMethod: 'apiKey',
      apiKey: 'k',
    });
    useSettingsStore.getState().addModel(epId, {
      id: 'gpt-4o',
      name: 'GPT 4o',
      ownedBy: 'system',
      provider: 'OpenAI',
    });

    useSettingsStore.getState().reset();

    expect(useSettingsStore.getState().endpoints).toEqual([]);
    expect(useSettingsStore.getState().activeEndpointId).toBeNull();
    expect(useSettingsStore.getState().activeModelId).toBeNull();
    expect(useSettingsStore.getState().endpointModels).toEqual({});
  });
});

// ─── Skill management ───

describe('settingsStore — skills', () => {
  it('starts with all skills enabled (null)', () => {
    expect(useSettingsStore.getState().activeSkillNames).toBeNull();
  });

  it('toggleSkill on null materializes list minus toggled skill', () => {
    useSettingsStore.getState().toggleSkill('xa2');
    // Should be an explicit array that does NOT contain xa2
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
    // Start from explicit list
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
});

// ─── Agent management ───

describe('settingsStore — agents', () => {
  it('starts with "Excel" as the default active agent', () => {
    expect(useSettingsStore.getState().activeAgentId).toBe('Excel');
  });

  it('setActiveAgent changes the active agent', () => {
    // Excel is a valid bundled agent
    useSettingsStore.getState().setActiveAgent('Excel');
    expect(useSettingsStore.getState().activeAgentId).toBe('Excel');
  });

  it('setActiveAgent ignores invalid agent names', () => {
    useSettingsStore.getState().setActiveAgent('NonExistentAgent');
    // Should remain unchanged
    expect(useSettingsStore.getState().activeAgentId).toBe('Excel');
  });

  it('getActiveAgent returns the current agent id', () => {
    expect(useSettingsStore.getState().getActiveAgent()).toBe('Excel');
  });

  it('reset restores the default agent', () => {
    // Even if someone managed to set a different agent, reset brings back Excel
    useSettingsStore.getState().reset();
    expect(useSettingsStore.getState().activeAgentId).toBe('Excel');
  });
});
