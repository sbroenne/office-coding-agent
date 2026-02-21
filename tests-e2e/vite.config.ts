/**
 * Vite configuration for E2E Tests
 *
 * Builds a standalone test taskpane that runs Excel command tests
 * inside a real Excel instance and reports results to a test server.
 * Served on port 3001 to avoid conflicting with the dev add-in on 3000.
 */

import { defineConfig } from 'vite';
import path from 'path';
import { viteStaticCopy } from 'vite-plugin-static-copy';
import fs from 'fs';
import { getHttpsServerOptions } from 'office-addin-dev-certs';

// Custom plugin: import .md files as raw strings
function mdRawPlugin() {
  return {
    name: 'md-raw',
    transform(_code: string, id: string) {
      if (id.endsWith('.md')) {
        const raw = fs.readFileSync(id, 'utf-8');
        return { code: `export default ${JSON.stringify(raw)};`, map: null };
      }
    },
  };
}

export default defineConfig(async () => {
  const httpsOptions = await getHttpsServerOptions();

  return {
    root: __dirname,
    plugins: [
      mdRawPlugin(),
      viteStaticCopy({
        targets: [
          { src: '../assets/*', dest: 'assets' },
          { src: 'src/functions.json', dest: '.' },
        ],
      }),
    ],
    resolve: {
      alias: { '@': path.resolve(__dirname, '../src') },
    },
    define: {
      'process.env.COPILOT_SERVER_URL': JSON.stringify(
        process.env.COPILOT_SERVER_URL || 'wss://localhost:3000/api/copilot'
      ),
    },
    build: {
      outDir: 'dist',
      emptyOutDir: true,
      sourcemap: true,
      rollupOptions: {
        input: {
          taskpane: path.resolve(__dirname, 'test-taskpane.html'),
          functions: path.resolve(__dirname, 'src/functions.ts'),
        },
      },
    },
    server: {
      port: 3001,
      https: httpsOptions,
      headers: { 'Access-Control-Allow-Origin': '*' },
    },
    preview: {
      port: 3001,
      https: httpsOptions,
      headers: { 'Access-Control-Allow-Origin': '*' },
    },
  };
});
