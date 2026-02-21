import { defineConfig } from '@playwright/test';

export default defineConfig({
  testDir: './tests-ui',
  timeout: 30_000,
  retries: 0,
  use: {
    baseURL: 'https://localhost:3000',
    ignoreHTTPSErrors: true, // dev server uses self-signed cert
    screenshot: 'only-on-failure',
    trace: 'retain-on-failure',
  },
  projects: [
    {
      name: 'chromium',
      use: { browserName: 'chromium' },
    },
  ],
  // Dev server is started separately — don't auto-start it here
  // because Vite dev server with office-addin-dev-certs is complex.
  // Run `npm run dev` in another terminal first (starts proxy + Vite).
  expect: {
    timeout: 5_000,
  },
});
