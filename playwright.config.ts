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
  webServer: {
    command: 'node src/server.mjs',
    url: 'https://localhost:3000/api/ping',
    reuseExistingServer: true,
    ignoreHTTPSErrors: true,
    timeout: 60_000,
  },
  expect: {
    timeout: 5_000,
  },
});
