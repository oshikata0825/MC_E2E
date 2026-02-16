import { test, expect } from '@playwright/test';
import { BASE_URL } from './config';
import * as fs from 'fs';
import * as path from 'path';

// Load user definitions
const users = JSON.parse(fs.readFileSync(path.join(__dirname, 'account.dat'), 'utf8'));

test.describe('TC_021 Role Based Access Control - Global Tariffs Page Access', () => {
  
  for (const user of users) {
    
    test(`Perform Global Tariffs page access test for ${user.id}`, async ({ page }) => {
      test.setTimeout(210000);

      console.log(`--- Test Start: ${user.id} ---`);

      // Step 1: Login
      await test.step('Login', async () => {
        await page.goto(BASE_URL + '/Logon.aspx?ReturnUrl=%2fgtm%2fhome');
        await page.getByRole('textbox', { name: 'Company' }).fill(user.company);
        await page.getByRole('textbox', { name: 'Username' }).fill(user.id);
        await page.getByRole('textbox', { name: 'Password' }).fill(user.pass);
        await page.getByRole('link', { name: 'Login' }).click();
      });

      // Step 2: Access Global Tariffs and Verify
      await test.step('Access Global Tariffs and Verify', async () => {
        let screenshotTaken = false;
        for (let i = 1; i <= 5; i++) {
            console.log(`\n--- Repetition ${i}/5 for user ${user.id} ---`);
            
            await page.getByRole('button', { name: 'Content' }).click();
            // Use regex for exact match to avoid matching "Global Tariffs (Quick Lookup)" etc.
            await page.locator('a[role="menuitem"]').filter({ hasText: /^Global Tariffs$/ }).click();

            let mainFrameElement = page.locator('iframe[name="legacy-outlet"]');
            await mainFrameElement.waitFor({ state: 'attached', timeout: 30000 });
            let mainFrame = await mainFrameElement.contentFrame();
            if (!mainFrame) throw new Error('Could not access main iframe content.');

            // Verify page name
            const pageNameLabel = mainFrame.locator('#lbxPageName');
            await expect(pageNameLabel).toContainText('Global Tariffs', { timeout: 20000 });
            console.log('SUCCESS: Verified page name contains "Global Tariffs".');

            // Wait for page to fully render
            await page.waitForTimeout(3000);

            // Take screenshot
            const screenshotBuffer = await page.screenshot({ fullPage: true });
            await test.info().attach(`global_tariffs_access_${user.id}.png`, {
                body: screenshotBuffer,
                contentType: 'image/png'
            });

            console.log('SUCCESS: Screenshot attached. Test complete.');
            break;
        }
      });

      // Step 3: Logout
      await test.step('Logout', async () => {
        const userButtonName = `MC Test${user.id.substring(6)}`;
        await page.bringToFront();

        const userButton = page.getByRole('button', { name: new RegExp(`^${userButtonName}`) });
        await userButton.hover();
        await page.waitForTimeout(500); // Wait for menu to start opening
        await userButton.click();

        const logoutButton = page.locator('#logoutButton');
        await expect(logoutButton).toBeVisible({ timeout: 10000 });
        await logoutButton.click({ force: true });

        await expect(page.getByRole('textbox', { name: 'Company' })).toBeVisible({ timeout: 30000 });
      });
    });
  }
});

