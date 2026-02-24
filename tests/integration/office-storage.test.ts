/**
 * Integration tests for officeStorage.
 *
 * officeStorage prefers OfficeRuntime.storage and falls back to localStorage
 * when OfficeRuntime.storage is unavailable in a host runtime.
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

  it('falls back to localStorage when OfficeRuntime is not available', async () => {
    const saved = (globalThis as Record<string, unknown>).OfficeRuntime;
    delete (globalThis as Record<string, unknown>).OfficeRuntime;

    await officeStorage.setItem('key', 'v');
    await expect(officeStorage.getItem('key')).resolves.toBe('v');
    await officeStorage.removeItem('key');
    await expect(officeStorage.getItem('key')).resolves.toBeNull();

    (globalThis as Record<string, unknown>).OfficeRuntime = saved;
  });
});
