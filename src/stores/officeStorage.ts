/**
 * Zustand storage adapter backed by OfficeRuntime.storage.
 *
 * OfficeRuntime.storage is a persistent, async key-value store provided by
 * Office.js for add-ins using the SharedRuntime. Unlike localStorage, it
 * persists reliably across Excel sessions and is not tied to a specific
 * workbook — settings survive when the user opens a different file.
 *
 * Falls back to window.localStorage when OfficeRuntime.storage isn't available
 * for a host/runtime combination, allowing the task pane to boot in hosts
 * with limited SharedRuntime support.
 *
 * @see https://learn.microsoft.com/javascript/api/office-runtime/officeruntime.storage
 */
import type { StateStorage } from 'zustand/middleware';

interface OfficeRuntimeStorage {
  getItem(key: string): Promise<string | null>;
  setItem(key: string, value: string): Promise<void>;
  removeItem(key: string): Promise<void>;
}

let fallbackWarned = false;

function localFallbackStorage(): OfficeRuntimeStorage {
  if (typeof localStorage === 'undefined') {
    throw new Error(
      '[officeStorage] Neither OfficeRuntime.storage nor localStorage is available in this runtime.'
    );
  }

  if (!fallbackWarned) {
    fallbackWarned = true;
    console.warn('[officeStorage] OfficeRuntime.storage unavailable; using localStorage fallback.');
  }

  return {
    getItem: (key: string) => Promise.resolve(localStorage.getItem(key)),
    setItem: (key: string, value: string) => {
      localStorage.setItem(key, value);
      return Promise.resolve();
    },
    removeItem: (key: string) => {
      localStorage.removeItem(key);
      return Promise.resolve();
    },
  };
}

function getStorage(): OfficeRuntimeStorage {
  if (typeof OfficeRuntime === 'undefined' || !OfficeRuntime?.storage) {
    return localFallbackStorage();
  }
  return OfficeRuntime.storage as OfficeRuntimeStorage;
}

export const officeStorage: StateStorage = {
  getItem: async (name: string) =>
    getStorage()
      .getItem(name)
      .then(v => v ?? null),
  setItem: async (name: string, value: string) => getStorage().setItem(name, value),
  removeItem: async (name: string) => getStorage().removeItem(name),
};
