import { describe, expect, it } from 'vitest';
import { isAllowedOrigin, isTrustedRequestOrigin } from '@/serverSecurity.mjs';

describe('serverSecurity origin checks', () => {
  it('allows the deployed GitHub Pages task pane origin', () => {
    expect(isAllowedOrigin('https://sbroenne.github.io')).toBe(true);
    expect(isAllowedOrigin('https://sbroenne.github.io/office-coding-agent/taskpane.html')).toBe(true);
  });

  it('does not allow unrelated GitHub Pages origins', () => {
    expect(isAllowedOrigin('https://example.github.io')).toBe(false);
  });

  it('trusts websocket requests from the deployed GitHub Pages origin', () => {
    expect(isTrustedRequestOrigin('https://sbroenne.github.io', '203.0.113.10')).toBe(true);
  });
});
