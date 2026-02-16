import { describe, it, expect, beforeEach } from 'vitest';
import { getAzureProvider, invalidateClient, clearAllClients } from '@/services/ai/aiClientFactory';
import type { FoundryEndpoint } from '@/types';

const validEndpoint: FoundryEndpoint = {
  id: 'test-ep',
  displayName: 'Test',
  resourceUrl: 'https://test.openai.azure.com',
  authMethod: 'apiKey',
  apiKey: 'test-key-123',
};

describe('aiClientFactory', () => {
  beforeEach(() => {
    clearAllClients();
  });

  it('throws when resourceUrl is empty', () => {
    expect(() =>
      getAzureProvider({ ...validEndpoint, resourceUrl: '' })
    ).toThrow('Endpoint URL is required');
  });

  it('throws when resourceUrl is whitespace', () => {
    expect(() =>
      getAzureProvider({ ...validEndpoint, resourceUrl: '   ' })
    ).toThrow('Endpoint URL is required');
  });

  it('throws when apiKey is empty', () => {
    expect(() =>
      getAzureProvider({ ...validEndpoint, apiKey: '' })
    ).toThrow('API Key is required');
  });

  it('throws when apiKey is whitespace', () => {
    expect(() =>
      getAzureProvider({ ...validEndpoint, apiKey: '   ' })
    ).toThrow('API Key is required');
  });

  it('returns a provider for valid config', () => {
    const provider = getAzureProvider(validEndpoint);
    expect(provider).toBeDefined();
  });

  it('caches provider by endpoint ID', () => {
    const first = getAzureProvider(validEndpoint);
    const second = getAzureProvider(validEndpoint);
    expect(first).toBe(second); // same instance
  });

  it('invalidateClient removes cached provider', () => {
    const first = getAzureProvider(validEndpoint);
    invalidateClient(validEndpoint.id);
    const second = getAzureProvider(validEndpoint);
    expect(first).not.toBe(second); // new instance
  });

  it('clearAllClients removes all cached providers', () => {
    const first = getAzureProvider(validEndpoint);
    clearAllClients();
    const second = getAzureProvider(validEndpoint);
    expect(first).not.toBe(second);
  });
});
