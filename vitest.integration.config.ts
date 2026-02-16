import { defineConfig } from 'vitest/config';
import path from 'path';
import { config } from 'dotenv';
import { readFileSync } from 'fs';
import type { Plugin } from 'vite';

// Load .env for integration test credentials (FOUNDRY_ENDPOINT, etc.)
config();

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
 * These tests hit live Azure AI Foundry endpoints and require
 * FOUNDRY_ENDPOINT env var (set in .env). Auth is via API key or Entra ID.
 */
export default defineConfig({
  plugins: [rawMarkdownPlugin()],
  resolve: {
    alias: {
      '@': path.resolve(__dirname, 'src'),
    },
  },
  test: {
    environment: 'node',
    include: ['tests/integration/**/*.test.ts'],
    testTimeout: 60000, // 60s — API calls can be slow
  },
});
