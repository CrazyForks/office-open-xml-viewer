import { defineConfig, devices } from '@playwright/test';

export default defineConfig({
  testDir: './tests/visual',
  testMatch: '**/*.spec.ts',
  fullyParallel: false,
  reporter: [
    ['list'],
    ['html', { outputFolder: 'tests/visual/report', open: 'never' }],
  ],
  use: {
    baseURL: 'http://localhost:5180',
    actionTimeout: 30_000,
  },
  projects: [
    {
      name: 'chrome',
      use: {
        channel: 'chrome',
        deviceScaleFactor: 1,
        viewport: { width: 1280, height: 960 },
      },
    },
    {
      name: 'firefox',
      testMatch: '**/conformance.spec.ts',
      use: {
        ...devices['Desktop Firefox'],
        deviceScaleFactor: 1,
        viewport: { width: 1280, height: 960 },
      },
    },
    {
      name: 'webkit',
      testMatch: '**/conformance.spec.ts',
      use: {
        ...devices['Desktop Safari'],
        deviceScaleFactor: 1,
        viewport: { width: 1280, height: 960 },
      },
    },
  ],
  // Start the Vite dev server separately before running tests:
  //   pnpm exec vite --port 5180
  webServer: {
    command: 'pnpm exec vite --port 5180 --strictPort',
    url: 'http://localhost:5180/tests/visual/fixture.html',
    reuseExistingServer: true,
    timeout: 120_000,
  },
});
