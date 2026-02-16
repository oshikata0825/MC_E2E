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
      name: 'SO-Creation',
      testMatch: /SO-01010100-FullApproval-NotSkipped\.test\.ts/,
      use: { 
        browserName: 'chromium',
        channel: 'chrome',
        executablePath: 'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe',
        viewport: { width: 1280, height: 720 },
      },
    },
    {
      name: 'SO-Approval',
      testMatch: /SO-APP-01010100-FullApproval-NotSkipped\.test\.ts/,
      dependencies: ['SO-Creation'],
      use: { 
        browserName: 'chromium',
        channel: 'chrome',
        executablePath: 'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe',
        viewport: { width: 1280, height: 720 },
      },
    },
    {
      name: 'SO-Creation-Provisional',
      testMatch: /SO-01010100-ProvisionalApproval-NotSkipped\.test\.ts/,
      use: { 
        browserName: 'chromium',
        channel: 'chrome',
        executablePath: 'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe',
        viewport: { width: 1280, height: 720 },
      },
    },
    {
      name: 'SO-Approval-Provisional',
      testMatch: /SO-APP-01010100-ProvisionalApproval-NotSkipped\.test\.ts/,
      dependencies: ['SO-Creation-Provisional'],
      use: { 
        browserName: 'chromium',
        channel: 'chrome',
        executablePath: 'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe',
        viewport: { width: 1280, height: 720 },
      },
    },
    {
      name: 'SO-Creation-Skipped',
      testMatch: /SO-01010100-FullApproval-Skipped\.test\.ts/,
      use: { 
        browserName: 'chromium',
        channel: 'chrome',
        executablePath: 'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe',
        viewport: { width: 1280, height: 720 },
      },
    },
    {
      name: 'SO-Approval-Skipped',
      testMatch: /SO-APP-01010100-FullApproval-Skipped\.test\.ts/,
      dependencies: ['SO-Creation-Skipped'],
      use: { 
        browserName: 'chromium',
        channel: 'chrome',
        executablePath: 'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe',
        viewport: { width: 1280, height: 720 },
      },
    },
  ],
});