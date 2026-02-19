import { describe, it, expect } from 'vitest';
import { PROVIDER_CONFIGS, PROVIDER_ORDER } from '@/services/ai/providerConfig';
import type { ProviderType } from '@/types';

describe('providerConfig', () => {
  it('PROVIDER_ORDER contains every ProviderType exactly once', () => {
    const allTypes: ProviderType[] = ['azure', 'anthropic', 'openai', 'mistral', 'deepseek', 'xai'];
    expect(PROVIDER_ORDER).toHaveLength(allTypes.length);
    for (const t of allTypes) {
      expect(PROVIDER_ORDER).toContain(t);
    }
  });

  it('PROVIDER_CONFIGS has an entry for every type in PROVIDER_ORDER', () => {
    for (const t of PROVIDER_ORDER) {
      expect(PROVIDER_CONFIGS[t]).toBeDefined();
      expect(PROVIDER_CONFIGS[t].label).toBeTruthy();
    }
  });

  it('azure has no fixed baseUrl (user provides their own)', () => {
    expect(PROVIDER_CONFIGS.azure.baseUrl).toBeUndefined();
  });

  it('non-azure providers have a fixed baseUrl', () => {
    const nonAzure = PROVIDER_ORDER.filter(t => t !== 'azure');
    for (const t of nonAzure) {
      expect(PROVIDER_CONFIGS[t].baseUrl).toBeTruthy();
    }
  });

  it('non-azure providers have at least one default model', () => {
    const nonAzure = PROVIDER_ORDER.filter(t => t !== 'azure');
    for (const t of nonAzure) {
      expect(PROVIDER_CONFIGS[t].defaultModels.length).toBeGreaterThan(0);
    }
  });

  it.each(PROVIDER_ORDER)('%s label is a non-empty string', type => {
    expect(typeof PROVIDER_CONFIGS[type].label).toBe('string');
    expect(PROVIDER_CONFIGS[type].label.length).toBeGreaterThan(0);
  });
});
