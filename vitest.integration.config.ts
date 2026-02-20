import { defineConfig } from 'vitest/config';
import path from 'path';
import { readFileSync } from 'fs';
import type { Plugin } from 'vite';

/**
 * Vite plugin that imports .md files as raw strings.
 * Same as the one in vitest.config.ts — needed because
 * chatService → skillService imports SKILL.md.
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
 * Vitest configuration for integration tests.
 *
 * Copilot WebSocket tests require `npm run server` to be running on localhost:3000.
 * They skip automatically when the server is unreachable.
 * Other integration tests are pure component/store wiring with no network calls.
 */
export default defineConfig({
  plugins: [rawMarkdownPlugin()],
  resolve: {
    alias: {
      '@': path.resolve(__dirname, 'src'),
    },
  },
  test: {
    environment: 'jsdom',
    include: ['tests/integration/**/*.test.ts', 'tests/integration/**/*.test.tsx'],
    testTimeout: 60000, // 60s — live Copilot calls can be slow
    setupFiles: ['tests/setup.ts'],
    globals: true,
  },
});
