import { defineConfig, defaultExclude } from 'vitest/config';
import path from 'path';
import { readFileSync } from 'fs';
import type { Plugin } from 'vite';

/**
 * Vite plugin that imports .md files as raw strings.
 * Matches webpack's `asset/source` behavior for markdown files.
 */
function rawMarkdownPlugin(): Plugin {
  return {
    name: 'raw-markdown',
    transform(_code: string, id: string) {
      if (id.endsWith('.md')) {
        const content = readFileSync(id, 'utf-8');
        return { code: `export default ${JSON.stringify(content)};`, map: null };
      }
    },
  };
}

/**
 * Vitest Configuration for Excel AI Add-in
 *
 * Test configuration for the Office Excel AI Assistant Add-in project.
 */
export default defineConfig({
  plugins: [rawMarkdownPlugin()],
  resolve: {
    alias: {
      '@': path.resolve(__dirname, 'src'),
    },
  },
  define: {
    // Override build-time env vars so SetupWizard starts with blank defaults
    'process.env.AZURE_OPENAI_ENDPOINT': JSON.stringify(''),
    'process.env.AZURE_OPENAI_API_KEY': JSON.stringify(''),
  },
  test: {
    // jsdom for all tests — needed by React component tests;
    // pure .ts unit tests work fine in jsdom as well
    environment: 'jsdom',

    // Match unit and component test files (integration tests use vitest.integration.config.ts)
    include: ['tests/**/*.test.ts', 'tests/**/*.test.tsx'],
    // Exclude API-calling integration tests (they use vitest.integration.config.ts with node env + 60s timeout)
    exclude: [...defaultExclude, 'tests/**/*.integration.test.ts'],

    // Setup file for testing-library matchers
    setupFiles: ['tests/setup.ts'],

    // Test timeout (30 seconds for API calls)
    testTimeout: 30000,

    // Code coverage configuration
    coverage: {
      provider: 'v8',
      include: ['src/**/*.ts', 'src/**/*.tsx'],
      exclude: ['src/**/*.d.ts'],
    },

    // Use globals (describe, it, expect) without imports
    globals: true,
  },
});
