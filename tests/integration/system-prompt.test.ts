/**
 * Integration tests for system prompt construction.
 *
 * Tests getAppPromptForHost() and buildSystemPrompt() from
 * src/services/ai/systemPrompt.ts.
 */
import { describe, it, expect } from 'vitest';
import {
  getAppPromptForHost,
  buildSessionSystemPrompt,
  buildSystemPrompt,
  BASE_PROMPT,
} from '@/services/ai/systemPrompt';

describe('Integration: systemPrompt', () => {
  describe('getAppPromptForHost', () => {
    it('does not add app-owned host behavior for bundled plugin agents', () => {
      const hosts = ['excel', 'powerpoint', 'word', 'outlook', 'unknown'] as const;
      for (const host of hosts) {
        expect(getAppPromptForHost(host as any)).toBe('');
      }
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

    it('uses the same app UI protocol prompt for every known host', () => {
      const hosts = ['excel', 'powerpoint', 'word', 'outlook'] as const;
      for (const host of hosts) {
        const prompt = buildSystemPrompt(host);
        expect(prompt).toBe(BASE_PROMPT);
      }
    });
  });

  describe('buildSessionSystemPrompt', () => {
    it('appends memory context after app UI protocol instructions', () => {
      const prompt = buildSessionSystemPrompt('excel', {
        memoryContext: 'Remember: prefer concise summaries.',
      });

      expect(prompt).not.toContain('## Host Agent Instructions');
      expect(prompt).not.toContain('## Adaptive Agent Instructions');
      expect(prompt).toContain('Remember: prefer concise summaries.');
      expect(prompt.trim().endsWith('Remember: prefer concise summaries.')).toBe(true);
    });
  });
});
