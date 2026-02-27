import { defineConfig, defaultExclude } from 'vitest/config';
import path from 'path';
import { readFileSync } from 'fs';
import type { Plugin } from 'vite';

/**
 * Vite plugin that imports .md files as raw strings.
 * Matches Vite's md-raw plugin behavior for markdown files.
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

const sharedViteConfig = {
  plugins: [rawMarkdownPlugin()],
  resolve: {
    alias: {
      '@': path.resolve(__dirname, 'src'),
    },
  },
};

/**
 * Vitest Configuration for Excel AI Add-in
 *
 * Uses Vitest projects to consolidate unit and integration configs into a single
 * file, preventing the VS Code Vitest extension from exceeding its 5-project limit.
 *
 * Projects:
 *   unit        — jsdom, 30s timeout, all tests except tests/integration/ and *.integration.test.ts
 *   integration — jsdom, 60s timeout, tests/integration/ (including live Copilot WebSocket tests)
 *
 * Run all:         vitest run
 * Run integration: vitest run --project integration
 * Watch:           vitest
 */
export default defineConfig({
  test: {
    coverage: {
      provider: 'v8',
      include: ['src/**/*.ts', 'src/**/*.tsx'],
      exclude: ['src/**/*.d.ts'],
    },
    projects: [
      {
        ...sharedViteConfig,
        define: {
          // Override build-time env vars so tests start with blank defaults
          'process.env.AZURE_OPENAI_ENDPOINT': JSON.stringify(''),
          'process.env.AZURE_OPENAI_API_KEY': JSON.stringify(''),
        },
        test: {
          name: 'unit',
          // jsdom for all tests — needed by React component tests
          environment: 'jsdom',
          // All tests except server-dependent live Copilot WebSocket tests
          include: ['tests/**/*.test.ts', 'tests/**/*.test.tsx'],
          exclude: [
            ...defaultExclude,
            // Skip live-server-dependent integration tests
            'tests/**/*.integration.test.ts',
          ],
          setupFiles: ['tests/setup.ts'],
          testTimeout: 30000,
          globals: true,
        },
      },
      {
        ...sharedViteConfig,
        test: {
          name: 'integration',
          // jsdom — React component integration tests need DOM
          environment: 'jsdom',
          // All tests in tests/integration/ including live Copilot WebSocket tests
          include: ['tests/integration/**/*.test.ts', 'tests/integration/**/*.test.tsx'],
          // 60s — live Copilot calls can be slow
          testTimeout: 60000,
          setupFiles: ['tests/setup.ts'],
          globals: true,
        },
      },
    ],
  },
});
