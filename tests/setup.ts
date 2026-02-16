import '@testing-library/jest-dom/vitest';
import { config } from 'dotenv';

// Load .env for integration test credentials (FOUNDRY_ENDPOINT, FOUNDRY_API_KEY, etc.)
config();

// ─── Polyfills for jsdom ───
// Fluent UI MessageBar uses ResizeObserver which jsdom lacks
if (typeof globalThis.ResizeObserver === 'undefined') {
  globalThis.ResizeObserver = class ResizeObserver {
    observe() {}
    unobserve() {}
    disconnect() {}
  } as unknown as typeof globalThis.ResizeObserver;
}

// Fluent Copilot's useScrollToBottom uses IntersectionObserver which jsdom lacks
if (typeof globalThis.IntersectionObserver === 'undefined') {
  globalThis.IntersectionObserver = class IntersectionObserver {
    readonly root = null;
    readonly rootMargin = '0px';
    readonly thresholds: readonly number[] = [0];
    observe() {}
    unobserve() {}
    disconnect() {}
    takeRecords(): IntersectionObserverEntry[] {
      return [];
    }
  } as unknown as typeof globalThis.IntersectionObserver;
}

// App.tsx uses window.matchMedia for dark mode detection
if (typeof window !== 'undefined' && !window.matchMedia) {
  window.matchMedia = (query: string) =>
    ({
      matches: false,
      media: query,
      onchange: null,
      addListener: () => {},
      removeListener: () => {},
      addEventListener: () => {},
      removeEventListener: () => {},
      dispatchEvent: () => false,
    }) as MediaQueryList;
}

// ─── Clear build-time environment defaults ───
// SetupWizard reads ENV_ENDPOINT / ENV_API_KEY from process.env at import time.
// In tests we want blank defaults so component tests start from a clean state.
process.env.AZURE_OPENAI_ENDPOINT = '';
process.env.AZURE_OPENAI_API_KEY = '';
