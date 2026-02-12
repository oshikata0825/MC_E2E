import { defineConfig } from '@playwright/test';

export default defineConfig({
  testDir: './tests',
  workers: 2,
  reporter: 'html',
  use: {
    trace: 'on-first-retry',
    headless: false,
  },
  timeout: 480000, // 8 minutes

  projects: [
    {
      name: 'Google Chrome',
      use: { 
        browserName: 'chromium',
        channel: 'chrome',
        executablePath: 'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe',
        viewport: { width: 1280, height: 720 },
      },
    },
  ],
});