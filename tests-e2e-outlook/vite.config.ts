/**
 * Vite configuration for Outlook E2E Tests
 *
 * Builds a standalone test taskpane that runs Outlook tool tests
 * inside a real Outlook instance and reports results to a test server.
 * Served on port 3004 to avoid conflicts with other test add-ins.
 */

import { defineConfig } from 'vite';
import path from 'path';
import { viteStaticCopy } from 'vite-plugin-static-copy';
import fs from 'fs';
import { getHttpsServerOptions } from 'office-addin-dev-certs';

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
        targets: [{ src: '../assets/*', dest: 'assets' }],
      }),
    ],
    resolve: {
      alias: { '@': path.resolve(__dirname, '../src') },
    },
    build: {
      outDir: 'dist',
      emptyOutDir: true,
      sourcemap: true,
      rollupOptions: {
        input: {
          taskpane: path.resolve(__dirname, 'test-taskpane.html'),
        },
      },
    },
    server: {
      port: 3004,
      https: httpsOptions,
      hmr: false,
    },
    preview: {
      port: 3004,
      https: httpsOptions,
    },
  };
});
