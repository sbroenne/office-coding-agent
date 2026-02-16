/**
 * Unit tests for generateId.
 *
 * Tests the unique ID generation utility.
 */

import { describe, it, expect, vi, afterEach } from 'vitest';
import { generateId } from '@/utils/id';

describe('generateId', () => {
  afterEach(() => {
    vi.restoreAllMocks();
  });

  it('returns a string', () => {
    expect(typeof generateId()).toBe('string');
  });

  it('returns non-empty strings', () => {
    expect(generateId().length).toBeGreaterThan(0);
  });

  it('returns unique values across multiple calls', () => {
    const ids = new Set(Array.from({ length: 100 }, () => generateId()));
    expect(ids.size).toBe(100);
  });

  it('uses crypto.randomUUID when available', () => {
    const mockUUID = '550e8400-e29b-41d4-a716-446655440000';
    vi.spyOn(crypto, 'randomUUID').mockReturnValue(
      mockUUID as `${string}-${string}-${string}-${string}-${string}`
    );

    expect(generateId()).toBe(mockUUID);
  });

  it('falls back to Date.now + random when crypto.randomUUID is undefined', () => {
    const original = crypto.randomUUID;
    // Temporarily remove randomUUID
    Object.defineProperty(crypto, 'randomUUID', { value: undefined, configurable: true });

    const id = generateId();
    expect(id).toMatch(/^\d+-[a-z0-9]+$/); // "timestamp-randomchars"

    // Restore
    Object.defineProperty(crypto, 'randomUUID', { value: original, configurable: true });
  });
});
