/**
 * Unit tests for officeStorage.
 *
 * In jsdom (Vitest), OfficeRuntime is undefined, so all operations
 * fall back to localStorage. These tests verify the fallback path
 * works correctly — the OfficeRuntime path is tested in E2E.
 */

import { describe, it, expect, beforeEach } from 'vitest';
import { officeStorage } from '@/stores/officeStorage';

describe('officeStorage (localStorage fallback)', () => {
  beforeEach(() => {
    localStorage.clear();
  });

  it('getItem returns null for missing keys', async () => {
    expect(await officeStorage.getItem('nonexistent')).toBeNull();
  });

  it('setItem persists a value', async () => {
    await officeStorage.setItem('test-key', 'test-value');
    expect(localStorage.getItem('test-key')).toBe('test-value');
  });

  it('getItem retrieves a previously set value', async () => {
    await officeStorage.setItem('round-trip', '{"data":42}');
    expect(await officeStorage.getItem('round-trip')).toBe('{"data":42}');
  });

  it('removeItem removes a value', async () => {
    await officeStorage.setItem('to-remove', 'value');
    await officeStorage.removeItem('to-remove');
    expect(await officeStorage.getItem('to-remove')).toBeNull();
  });

  it('round-trip: set → get → remove → get returns null', async () => {
    const key = 'lifecycle-key';
    const value = JSON.stringify({ endpoints: [], activeModelId: null });

    await officeStorage.setItem(key, value);
    expect(await officeStorage.getItem(key)).toBe(value);

    await officeStorage.removeItem(key);
    expect(await officeStorage.getItem(key)).toBeNull();
  });

  it('handles empty string values', async () => {
    await officeStorage.setItem('empty', '');
    expect(await officeStorage.getItem('empty')).toBe('');
  });

  it('handles large values', async () => {
    const largeValue = 'x'.repeat(10_000);
    await officeStorage.setItem('large', largeValue);
    expect(await officeStorage.getItem('large')).toBe(largeValue);
  });

  it('overwrites existing values', async () => {
    await officeStorage.setItem('overwrite', 'first');
    await officeStorage.setItem('overwrite', 'second');
    expect(await officeStorage.getItem('overwrite')).toBe('second');
  });
});
