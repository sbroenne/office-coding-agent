/**
 * Zustand storage adapter backed by OfficeRuntime.storage.
 *
 * OfficeRuntime.storage is a persistent, async key-value store provided by
 * Office.js for add-ins using the SharedRuntime. Unlike localStorage (which
 * the WebView may clear), this storage persists reliably across sessions and
 * is NOT tied to a specific workbook — settings survive when the user opens
 * a different file.
 *
 * In tests or environments where OfficeRuntime isn't available, we fall back
 * to localStorage so the store still works outside Excel.
 *
 * @see https://learn.microsoft.com/javascript/api/office-runtime/officeruntime.storage
 */
import type { StateStorage } from 'zustand/middleware';

/**
 * Type for the OfficeRuntime global injected by Office.js shared runtime.
 */
interface OfficeRuntimeStorage {
  getItem(key: string): Promise<string | null>;
  setItem(key: string, value: string): Promise<void>;
  removeItem(key: string): Promise<void>;
}

/**
 * Safely access the OfficeRuntime.storage API.
 *
 * Uses `typeof` to avoid ReferenceError when the global doesn't exist
 * (e.g. in tests, dev server without sideload, or non-Excel environments).
 */
function getOfficeStorage(): OfficeRuntimeStorage | undefined {
  // typeof is safe for undeclared variables — unlike a bare reference,
  // it won't throw ReferenceError when OfficeRuntime doesn't exist.

  if (typeof OfficeRuntime !== 'undefined' && OfficeRuntime?.storage) {
    return OfficeRuntime.storage as OfficeRuntimeStorage;
  }
  return undefined;
}

/** Log which storage backend is actually used on first access */
let _logged = false;
function logBackend(storage: OfficeRuntimeStorage | undefined) {
  if (_logged) return;
  _logged = true;
  const backend = storage ? 'OfficeRuntime.storage' : 'localStorage (fallback)';
  console.log(`[officeStorage] Using ${backend}`);
}

/**
 * Zustand-compatible StateStorage implementation backed by OfficeRuntime.storage,
 * with a graceful fallback to localStorage when running outside the Office runtime
 * (tests, standalone dev server, etc.).
 */
export const officeStorage: StateStorage = {
  getItem: async (name: string): Promise<string | null> => {
    const storage = getOfficeStorage();
    logBackend(storage);
    if (storage) {
      try {
        const value = await storage.getItem(name);
        return value ?? null;
      } catch (err) {
        console.warn('[officeStorage] getItem failed, falling back to localStorage:', err);
      }
    }
    // Fallback
    return localStorage.getItem(name);
  },

  setItem: async (name: string, value: string): Promise<void> => {
    const storage = getOfficeStorage();
    if (storage) {
      try {
        await storage.setItem(name, value);
        return;
      } catch (err) {
        console.warn('[officeStorage] setItem failed, falling back to localStorage:', err);
      }
    }
    // Fallback
    localStorage.setItem(name, value);
  },

  removeItem: async (name: string): Promise<void> => {
    const storage = getOfficeStorage();
    if (storage) {
      try {
        await storage.removeItem(name);
        return;
      } catch (err) {
        console.warn('[officeStorage] removeItem failed, falling back to localStorage:', err);
      }
    }
    // Fallback
    localStorage.removeItem(name);
  },
};
