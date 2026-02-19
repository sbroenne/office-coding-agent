import { describe, it, expect, beforeEach } from 'vitest';
import {
  getAzureProvider,
  getProviderModel,
  invalidateClient,
  clearAllClients,
} from '@/services/ai/aiClientFactory';
import type { FoundryEndpoint } from '@/types';

const validAzureEndpoint: FoundryEndpoint = {
  id: 'test-ep',
  displayName: 'Test',
  resourceUrl: 'https://test.openai.azure.com',
  authMethod: 'apiKey',
  apiKey: 'test-key-123',
  providerType: 'azure',
};

describe('aiClientFactory', () => {
  beforeEach(() => {
    clearAllClients();
  });

  // ── Azure provider ──────────────────────────────────────────

  it('throws when resourceUrl is empty', () => {
    expect(() =>
      getAzureProvider({ ...validAzureEndpoint, resourceUrl: '' })
    ).toThrow('Endpoint URL is required');
  });

  it('throws when resourceUrl is whitespace', () => {
    expect(() =>
      getAzureProvider({ ...validAzureEndpoint, resourceUrl: '   ' })
    ).toThrow('Endpoint URL is required');
  });

  it('throws when apiKey is empty', () => {
    expect(() =>
      getAzureProvider({ ...validAzureEndpoint, apiKey: '' })
    ).toThrow('API Key is required');
  });

  it('throws when apiKey is whitespace', () => {
    expect(() =>
      getAzureProvider({ ...validAzureEndpoint, apiKey: '   ' })
    ).toThrow('API Key is required');
  });

  it('returns a provider for valid config', () => {
    const provider = getAzureProvider(validAzureEndpoint);
    expect(provider).toBeDefined();
  });

  it('caches provider by endpoint ID', () => {
    const first = getAzureProvider(validAzureEndpoint);
    const second = getAzureProvider(validAzureEndpoint);
    expect(first).toBe(second); // same instance
  });

  it('invalidateClient removes cached provider', () => {
    const first = getAzureProvider(validAzureEndpoint);
    invalidateClient(validAzureEndpoint.id);
    const second = getAzureProvider(validAzureEndpoint);
    expect(first).not.toBe(second); // new instance
  });

  it('clearAllClients removes all cached providers', () => {
    const first = getAzureProvider(validAzureEndpoint);
    clearAllClients();
    const second = getAzureProvider(validAzureEndpoint);
    expect(first).not.toBe(second);
  });

  // ── getProviderModel routes to correct provider ─────────────

  it.each([
    ['anthropic', 'sk-ant-key', 'claude-sonnet-4-5'],
    ['openai', 'sk-openai-key', 'gpt-4o'],
    ['mistral', 'mistral-key', 'mistral-large-latest'],
    ['deepseek', 'ds-key', 'deepseek-chat'],
    ['xai', 'xai-key', 'grok-3-beta'],
  ] as const)(
    'getProviderModel returns a LanguageModel for %s',
    (providerType, apiKey, modelId) => {
      const endpoint: FoundryEndpoint = {
        id: `ep-${providerType}`,
        displayName: providerType,
        resourceUrl: '',
        authMethod: 'apiKey',
        apiKey,
        providerType,
      };
      const model = getProviderModel(endpoint, modelId);
      expect(model).toBeDefined();
    }
  );

  it('getProviderModel defaults to azure when providerType is undefined', () => {
    const endpoint: FoundryEndpoint = {
      ...validAzureEndpoint,
      providerType: undefined,
    };
    // Should not throw — falls back to Azure
    const model = getProviderModel(endpoint, 'gpt-4o');
    expect(model).toBeDefined();
  });
});
