/**
 * Integration tests for system prompt construction.
 *
 * Tests getAppPromptForHost() and buildSystemPrompt() from
 * src/services/ai/systemPrompt.ts.
 */
import { describe, it, expect } from 'vitest';
import { getAppPromptForHost, buildSystemPrompt, BASE_PROMPT } from '@/services/ai/systemPrompt';

describe('Integration: systemPrompt', () => {
  describe('getAppPromptForHost', () => {
    it('returns non-empty prompt for excel', () => {
      const prompt = getAppPromptForHost('excel');
      expect(prompt).toBeTruthy();
      expect(prompt.length).toBeGreaterThan(10);
    });

    it('returns non-empty prompt for powerpoint', () => {
      const prompt = getAppPromptForHost('powerpoint');
      expect(prompt).toBeTruthy();
      expect(prompt.length).toBeGreaterThan(10);
    });

    it('returns non-empty prompt for word', () => {
      const prompt = getAppPromptForHost('word');
      expect(prompt).toBeTruthy();
      expect(prompt.length).toBeGreaterThan(10);
    });

    it('returns non-empty prompt for outlook', () => {
      const prompt = getAppPromptForHost('outlook');
      expect(prompt).toBeTruthy();
      expect(prompt.length).toBeGreaterThan(10);
    });

    it('returns a default fallback for unknown host', () => {
      const prompt = getAppPromptForHost('unknown' as any);
      expect(prompt).toContain('Office');
    });

    it('returns different prompts for different hosts', () => {
      const excelPrompt = getAppPromptForHost('excel');
      const pptPrompt = getAppPromptForHost('powerpoint');
      const wordPrompt = getAppPromptForHost('word');
      const outlookPrompt = getAppPromptForHost('outlook');

      // All four host-specific prompts should be distinct
      const prompts = new Set([excelPrompt, pptPrompt, wordPrompt, outlookPrompt]);
      expect(prompts.size).toBe(4);
    });
  });

  describe('BASE_PROMPT', () => {
    it('is a non-empty string', () => {
      expect(typeof BASE_PROMPT).toBe('string');
      expect(BASE_PROMPT.length).toBeGreaterThan(50);
    });
  });

  describe('buildSystemPrompt', () => {
    it('includes BASE_PROMPT content', () => {
      const prompt = buildSystemPrompt('excel');
      expect(prompt).toContain(BASE_PROMPT);
    });

    it('includes host-specific prompt', () => {
      const hostPrompt = getAppPromptForHost('excel');
      const full = buildSystemPrompt('excel');
      expect(full).toContain(hostPrompt);
    });

    it('joins BASE_PROMPT and host prompt with double newline', () => {
      const full = buildSystemPrompt('powerpoint');
      const hostPrompt = getAppPromptForHost('powerpoint');
      expect(full).toBe(`${BASE_PROMPT}\n\n${hostPrompt}`);
    });

    it('works for every known host', () => {
      const hosts = ['excel', 'powerpoint', 'word', 'outlook'] as const;
      for (const host of hosts) {
        const prompt = buildSystemPrompt(host);
        expect(prompt.length).toBeGreaterThan(BASE_PROMPT.length);
      }
    });
  });
});
