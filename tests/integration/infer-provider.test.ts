/**
 * Integration tests for inferProvider().
 *
 * Validates model-ID → provider mapping logic from src/types/settings.ts.
 */
import { describe, it, expect } from 'vitest';
import { inferProvider } from '@/types/settings';

describe('Integration: inferProvider', () => {
  it('maps claude models to Anthropic', () => {
    expect(inferProvider('claude-sonnet-4')).toBe('Anthropic');
    expect(inferProvider('claude-sonnet-4.6')).toBe('Anthropic');
    expect(inferProvider('claude-3-opus')).toBe('Anthropic');
    expect(inferProvider('claude-3.5-haiku')).toBe('Anthropic');
  });

  it('maps gpt models to OpenAI', () => {
    expect(inferProvider('gpt-4')).toBe('OpenAI');
    expect(inferProvider('gpt-4o')).toBe('OpenAI');
    expect(inferProvider('gpt-4o-mini')).toBe('OpenAI');
    expect(inferProvider('gpt-3.5-turbo')).toBe('OpenAI');
  });

  it('maps o1 models to OpenAI', () => {
    expect(inferProvider('o1-preview')).toBe('OpenAI');
    expect(inferProvider('o1-mini')).toBe('OpenAI');
  });

  it('maps o3 models to OpenAI', () => {
    expect(inferProvider('o3')).toBe('OpenAI');
    expect(inferProvider('o3-mini')).toBe('OpenAI');
  });

  it('maps o4 models to OpenAI', () => {
    expect(inferProvider('o4-mini')).toBe('OpenAI');
  });

  it('maps gemini models to Google', () => {
    expect(inferProvider('gemini-pro')).toBe('Google');
    expect(inferProvider('gemini-1.5-flash')).toBe('Google');
    expect(inferProvider('gemini-2.0-flash')).toBe('Google');
  });

  it('maps unknown models to Other', () => {
    expect(inferProvider('llama-3')).toBe('Other');
    expect(inferProvider('mistral-large')).toBe('Other');
    expect(inferProvider('phi-3')).toBe('Other');
    expect(inferProvider('custom-model')).toBe('Other');
  });

  it('is case-sensitive (uppercase prefix returns Other)', () => {
    expect(inferProvider('Claude-3')).toBe('Other');
    expect(inferProvider('GPT-4o')).toBe('Other');
    expect(inferProvider('Gemini-pro')).toBe('Other');
  });

  it('handles empty string as Other', () => {
    expect(inferProvider('')).toBe('Other');
  });
});
