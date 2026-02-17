import { describe, expect, it } from 'vitest';
import { normalizeChatErrorMessage } from '@/services/ai/chatErrorMessage';

describe('normalizeChatErrorMessage', () => {
  it('normalizes rate-limit errors and includes retry hint when present', () => {
    const raw =
      'Failed after 5 attempts. Last error: 429 Too Many Requests. Please retry after 17 seconds.';

    const normalized = normalizeChatErrorMessage(raw);
    expect(normalized).toContain('Rate limit reached');
    expect(normalized).toContain('Retry in about 17s');
  });

  it('normalizes not-found/deployment errors', () => {
    const normalized = normalizeChatErrorMessage('HTTP 404 deployment not found');
    expect(normalized).toContain('Endpoint or deployment not found');
  });

  it('normalizes auth and permission errors', () => {
    expect(normalizeChatErrorMessage('401 Unauthorized')).toContain('Authentication failed');
    expect(normalizeChatErrorMessage('403 Forbidden')).toContain('Access denied');
  });

  it('passes through unknown errors', () => {
    const raw = 'Unexpected transport disconnect';
    expect(normalizeChatErrorMessage(raw)).toBe(raw);
  });
});
