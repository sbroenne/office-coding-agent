import { describe, it, expect } from 'vitest';
import { normalizeEndpoint } from '@/services/ai';

describe('normalizeEndpoint', () => {
  it('returns a plain resource URL unchanged', () => {
    expect(normalizeEndpoint('https://my-resource.openai.azure.com')).toBe(
      'https://my-resource.openai.azure.com'
    );
  });

  it('strips trailing slashes', () => {
    expect(normalizeEndpoint('https://my-resource.openai.azure.com/')).toBe(
      'https://my-resource.openai.azure.com'
    );
    expect(normalizeEndpoint('https://my-resource.openai.azure.com///')).toBe(
      'https://my-resource.openai.azure.com'
    );
  });

  it('strips /openai suffix to avoid doubling', () => {
    expect(normalizeEndpoint('https://my-resource.openai.azure.com/openai')).toBe(
      'https://my-resource.openai.azure.com'
    );
    expect(normalizeEndpoint('https://my-resource.openai.azure.com/openai/')).toBe(
      'https://my-resource.openai.azure.com'
    );
  });

  it('strips /openai/v1 suffix', () => {
    expect(normalizeEndpoint('https://my-resource.openai.azure.com/openai/v1')).toBe(
      'https://my-resource.openai.azure.com'
    );
    expect(normalizeEndpoint('https://my-resource.openai.azure.com/openai/v1/')).toBe(
      'https://my-resource.openai.azure.com'
    );
  });

  it('strips Foundry project path', () => {
    expect(
      normalizeEndpoint('https://my-resource.services.ai.azure.com/api/projects/proj-default')
    ).toBe('https://my-resource.services.ai.azure.com');
  });

  it('handles whitespace around the URL', () => {
    expect(normalizeEndpoint('  https://my-resource.openai.azure.com  ')).toBe(
      'https://my-resource.openai.azure.com'
    );
  });

  it('handles combined project path and trailing slash', () => {
    expect(
      normalizeEndpoint('https://my-resource.services.ai.azure.com/api/projects/my-proj/')
    ).toBe('https://my-resource.services.ai.azure.com');
  });
});
