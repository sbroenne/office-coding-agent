import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';
import { viteStaticCopy } from 'vite-plugin-static-copy';
import path from 'node:path';

export default defineConfig({
  root: '.',
  publicDir: false, // we copy assets explicitly

  resolve: {
    alias: {
      '@': path.resolve(__dirname, 'src'),
    },
  },

  plugins: [
    react(),

    // Import .md files as raw strings
    {
      name: 'md-raw',
      transform(_code, id) {
        if (id.endsWith('.md')) {
          const fs = require('node:fs');
          const raw = fs.readFileSync(id, 'utf-8');
          return { code: `export default ${JSON.stringify(raw)};`, map: null };
        }
      },
    },

    // Copy icon assets to dist/assets
    viteStaticCopy({
      targets: [{ src: 'assets/*', dest: 'assets' }],
    }),
  ],

  build: {
    outDir: 'dist',
    emptyOutDir: true,
    sourcemap: false,
    rollupOptions: {
      input: {
        taskpane: path.resolve(__dirname, 'taskpane.html'),
      },
    },
  },

  // Dev server config — only used when running `vite` directly (not middleware mode)
  server: {
    port: 3000,
    strictPort: true,
  },
});
