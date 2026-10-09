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
    expect(isTrustedRequestOrigin('https://sbroenne.github.io', '127.0.0.1')).toBe(true);
  });

  it('rejects requests from non-loopback addresses even with an allowed origin', () => {
    expect(isTrustedRequestOrigin('https://localhost:3000', '203.0.113.10')).toBe(false);
    expect(isTrustedRequestOrigin('https://sbroenne.github.io', '203.0.113.10')).toBe(false);
    expect(isTrustedRequestOrigin(undefined, '203.0.113.10')).toBe(false);
  });

  it('rejects untrusted origins from loopback addresses', () => {
    expect(isTrustedRequestOrigin('https://example.com', '127.0.0.1')).toBe(false);
  });

  it('allows local clients without an Origin header', () => {
    expect(isTrustedRequestOrigin(undefined, '127.0.0.1')).toBe(true);
    expect(isTrustedRequestOrigin(undefined, '::ffff:127.0.0.1')).toBe(true);
  });
});
