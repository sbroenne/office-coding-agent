/**
 * Unit tests for officeStorage.
 *
 * Two suites:
 *  1. OfficeRuntime.storage — localStorage recovery — tests the scenario where
 *     a previous setItem() failed and wrote to localStorage instead. On the next
 *     getItem() call, OfficeRuntime.storage has nothing for the key; we must also
 *     check localStorage so the data isn't silently lost.
 *     (Happy path + error fallbacks are in officeStorageRuntime.test.ts)
 *  2. localStorage fallback — OfficeRuntime is undefined (jsdom default), so all
 *     operations fall back to localStorage.
 */

import { describe, it, expect, beforeEach, afterEach, vi } from 'vitest';
import { officeStorage } from '@/stores/officeStorage';

// ─── 1. OfficeRuntime path — localStorage recovery scenario ──────────────────

describe('officeStorage (OfficeRuntime.storage — localStorage recovery)', () => {
  const mockStorage = {
    getItem: vi.fn(),
    setItem: vi.fn(),
    removeItem: vi.fn(),
  };

  beforeEach(() => {
    localStorage.clear();
    mockStorage.getItem.mockReset();
    mockStorage.setItem.mockReset();
    mockStorage.removeItem.mockReset();
    (globalThis as Record<string, unknown>).OfficeRuntime = { storage: mockStorage };
  });

  afterEach(() => {
    delete (globalThis as Record<string, unknown>).OfficeRuntime;
  });

  it('getItem checks localStorage when OfficeRuntime.storage returns null (recovery from failed setItem)', async () => {
    // Simulate: a previous setItem() failed and fell back to localStorage.
    // OfficeRuntime.storage has nothing for this key.
    mockStorage.getItem.mockResolvedValue(null);
    localStorage.setItem('settings', '{"state":{"endpoints":[]},"version":1}');

    const result = await officeStorage.getItem('settings');

    expect(result).toBe('{"state":{"endpoints":[]},"version":1}');
  });

  it('getItem checks localStorage when OfficeRuntime.storage returns undefined (recovery)', async () => {
    mockStorage.getItem.mockResolvedValue(undefined);
    localStorage.setItem('settings', '{"state":{"endpoints":[]},"version":1}');

    const result = await officeStorage.getItem('settings');

    expect(result).toBe('{"state":{"endpoints":[]},"version":1}');
  });

  it('getItem returns null when both OfficeRuntime and localStorage have no value', async () => {
    mockStorage.getItem.mockResolvedValue(null);

    expect(await officeStorage.getItem('nonexistent')).toBeNull();
  });
});

// ─── 2. localStorage fallback (OfficeRuntime undefined) ──────────────────────

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
