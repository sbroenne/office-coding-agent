/**
 * Unit tests for officeStorage.
 *
 * officeStorage exclusively uses OfficeRuntime.storage (no localStorage fallback).
 * Tests run against the in-memory mock supplied by tests/setup.ts, which resets
 * the mock store before each test via a global beforeEach hook.
 */

import { describe, it, expect } from 'vitest';
import { officeStorage } from '@/stores/officeStorage';

describe('officeStorage', () => {
  it('getItem returns null for a key that has not been set', async () => {
    expect(await officeStorage.getItem('nonexistent')).toBeNull();
  });

  it('setItem persists a value retrievable by getItem', async () => {
    await officeStorage.setItem('key', 'value');
    expect(await officeStorage.getItem('key')).toBe('value');
  });

  it('round-trip: setItem -> getItem returns the stored value', async () => {
    const payload = JSON.stringify({ endpoints: [{ id: 'abc' }], version: 1 });
    await officeStorage.setItem('settings', payload);
    expect(await officeStorage.getItem('settings')).toBe(payload);
  });

  it('removeItem deletes a previously stored value', async () => {
    await officeStorage.setItem('to-remove', 'value');
    await officeStorage.removeItem('to-remove');
    expect(await officeStorage.getItem('to-remove')).toBeNull();
  });

  it('setItem overwrites an existing value', async () => {
    await officeStorage.setItem('key', 'first');
    await officeStorage.setItem('key', 'second');
    expect(await officeStorage.getItem('key')).toBe('second');
  });

  it('handles large values', async () => {
    const large = 'x'.repeat(10_000);
    await officeStorage.setItem('large', large);
    expect(await officeStorage.getItem('large')).toBe(large);
  });

  it('throws a clear error when OfficeRuntime is not available', async () => {
    const saved = (globalThis as Record<string, unknown>).OfficeRuntime;
    delete (globalThis as Record<string, unknown>).OfficeRuntime;

    await expect(officeStorage.getItem('key')).rejects.toThrow(
      'OfficeRuntime.storage is not available'
    );
    await expect(officeStorage.setItem('key', 'v')).rejects.toThrow(
      'OfficeRuntime.storage is not available'
    );
    await expect(officeStorage.removeItem('key')).rejects.toThrow(
      'OfficeRuntime.storage is not available'
    );

    (globalThis as Record<string, unknown>).OfficeRuntime = saved;
  });
});
