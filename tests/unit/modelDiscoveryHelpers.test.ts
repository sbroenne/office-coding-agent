import { describe, it, expect } from 'vitest';
import { inferProvider, isEmbeddingOrUtilityModel, formatModelName } from '@/services/ai';

// ─── inferProvider ───

describe('inferProvider', () => {
  it.each([
    ['claude-3-sonnet', undefined, 'Anthropic'],
    ['claude-3.5-haiku', undefined, 'Anthropic'],
    ['some-model', 'anthropic', 'Anthropic'],
  ] as const)('maps %s (owner=%s) → %s', (modelId, ownedBy, expected) => {
    expect(inferProvider(modelId, ownedBy)).toBe(expected);
  });

  it.each([
    ['gpt-4o', 'OpenAI'],
    ['gpt-4.1', 'OpenAI'],
    ['o1-preview', 'OpenAI'],
    ['o3-mini', 'OpenAI'],
    ['o4-mini', 'OpenAI'],
  ] as const)('maps %s → %s', (modelId, expected) => {
    expect(inferProvider(modelId)).toBe(expected);
  });

  it.each([
    ['deepseek-r1', 'DeepSeek'],
    ['mai-ds-r1', 'DeepSeek'],
  ] as const)('maps %s → %s', (modelId, expected) => {
    expect(inferProvider(modelId)).toBe(expected);
  });

  it.each([
    ['llama-3.1-70b', undefined, 'Meta'],
    ['some-model', 'meta-llama', 'Meta'],
  ] as const)('maps %s (owner=%s) → %s', (modelId, ownedBy, expected) => {
    expect(inferProvider(modelId, ownedBy)).toBe(expected);
  });

  it.each([
    ['mistral-large', 'Mistral'],
    ['ministral-8b', 'Mistral'],
  ] as const)('maps %s → %s', (modelId, expected) => {
    expect(inferProvider(modelId)).toBe(expected);
  });

  it('maps grok → xAI', () => {
    expect(inferProvider('grok-2')).toBe('xAI');
  });

  it.each([
    ['phi-4', undefined, 'Microsoft'],
    ['some-model', 'microsoft', 'Microsoft'],
  ] as const)('maps %s (owner=%s) → %s', (modelId, ownedBy, expected) => {
    expect(inferProvider(modelId, ownedBy)).toBe(expected);
  });

  it('falls back to Other for unknown models', () => {
    expect(inferProvider('some-unknown-model')).toBe('Other');
  });
});

// ─── isEmbeddingOrUtilityModel ───

describe('isEmbeddingOrUtilityModel', () => {
  it.each([
    'text-embedding-ada-002',
    'text-embedding-3-large',
    'dall-e-3',
    'tts-1',
    'tts-1-hd',
    'whisper-1',
    'transcribe-model',
    'sora-turbo',
    'gpt-image-1',
  ])('returns true for non-chat model: %s', modelId => {
    expect(isEmbeddingOrUtilityModel(modelId)).toBe(true);
  });

  it.each(['gpt-4o', 'claude-3-sonnet', 'deepseek-r1', 'o1-preview', 'phi-4', 'llama-3.1-70b'])(
    'returns false for chat model: %s',
    modelId => {
      expect(isEmbeddingOrUtilityModel(modelId)).toBe(false);
    }
  );
});

// ─── formatModelName ───

describe('formatModelName', () => {
  it('capitalises words and joins with spaces', () => {
    expect(formatModelName('my-custom-model')).toBe('My Custom Model');
  });

  it('uppercases GPT prefix', () => {
    expect(formatModelName('gpt-4o')).toBe('GPT-4o');
    expect(formatModelName('gpt-4.1')).toBe('GPT-4.1');
  });

  it('formats Claude correctly', () => {
    expect(formatModelName('claude-3-sonnet')).toBe('Claude 3 Sonnet');
  });

  it('formats DeepSeek correctly', () => {
    expect(formatModelName('deepseek-r1')).toBe('DeepSeek R1');
  });

  it('formats MAI-DS prefix correctly', () => {
    expect(formatModelName('mai-ds-r1')).toBe('MAI-DS-R1');
  });
});
