/**
 * Tests for officeStorage when OfficeRuntime IS available.
 *
 * We simulate the OfficeRuntime global so officeStorage uses the
 * OfficeRuntime.storage path instead of the localStorage fallback.
 * Also tests the error-fallback path (OfficeRuntime throws → localStorage).
 */

import { describe, it, expect, beforeEach, afterEach, vi } from 'vitest';

// We need to re-import officeStorage after setting up the global,
// but the module is already cached. Instead, we'll test by directly
// setting up the global before importing.
// Since vitest caches modules, we use vi.resetModules() + dynamic import.

describe('officeStorage (OfficeRuntime path)', () => {
  const mockStorage = {
    getItem: vi.fn(),
    setItem: vi.fn(),
    removeItem: vi.fn(),
  };

  beforeEach(() => {
    vi.resetModules();
    localStorage.clear();
    mockStorage.getItem.mockReset();
    mockStorage.setItem.mockReset();
    mockStorage.removeItem.mockReset();

    // Simulate OfficeRuntime global
    (globalThis as Record<string, unknown>).OfficeRuntime = {
      storage: mockStorage,
    };
  });

  afterEach(() => {
    delete (globalThis as Record<string, unknown>).OfficeRuntime;
  });

  it('getItem uses OfficeRuntime.storage when available', async () => {
    mockStorage.getItem.mockResolvedValue('stored-value');
    const { officeStorage } = await import('@/stores/officeStorage');

    const result = await officeStorage.getItem('key');

    expect(mockStorage.getItem).toHaveBeenCalledWith('key');
    expect(result).toBe('stored-value');
  });

  it('getItem returns null when OfficeRuntime returns undefined', async () => {
    mockStorage.getItem.mockResolvedValue(undefined);
    const { officeStorage } = await import('@/stores/officeStorage');

    const result = await officeStorage.getItem('key');

    expect(result).toBeNull();
  });

  it('setItem uses OfficeRuntime.storage when available', async () => {
    mockStorage.setItem.mockResolvedValue(undefined);
    const { officeStorage } = await import('@/stores/officeStorage');

    await officeStorage.setItem('key', 'value');

    expect(mockStorage.setItem).toHaveBeenCalledWith('key', 'value');
    // Should NOT fall through to localStorage
    expect(localStorage.getItem('key')).toBeNull();
  });

  it('removeItem uses OfficeRuntime.storage when available', async () => {
    mockStorage.removeItem.mockResolvedValue(undefined);
    const { officeStorage } = await import('@/stores/officeStorage');

    await officeStorage.removeItem('key');

    expect(mockStorage.removeItem).toHaveBeenCalledWith('key');
  });

  it('getItem falls back to localStorage when OfficeRuntime throws', async () => {
    mockStorage.getItem.mockRejectedValue(new Error('storage error'));
    localStorage.setItem('fallback-key', 'fallback-value');
    const { officeStorage } = await import('@/stores/officeStorage');

    const result = await officeStorage.getItem('fallback-key');

    expect(mockStorage.getItem).toHaveBeenCalled();
    expect(result).toBe('fallback-value');
  });

  it('setItem falls back to localStorage when OfficeRuntime throws', async () => {
    mockStorage.setItem.mockRejectedValue(new Error('storage error'));
    const { officeStorage } = await import('@/stores/officeStorage');

    await officeStorage.setItem('fallback-key', 'fallback-value');

    expect(localStorage.getItem('fallback-key')).toBe('fallback-value');
  });

  it('removeItem falls back to localStorage when OfficeRuntime throws', async () => {
    mockStorage.removeItem.mockRejectedValue(new Error('storage error'));
    localStorage.setItem('to-remove', 'value');
    const { officeStorage } = await import('@/stores/officeStorage');

    await officeStorage.removeItem('to-remove');

    expect(localStorage.getItem('to-remove')).toBeNull();
  });
});
