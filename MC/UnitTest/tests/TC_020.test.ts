import { test, expect } from '@playwright/test';
import { BASE_URL } from './config';
import * as fs from 'fs';
import * as path from 'path';

// Load user definitions
const users = JSON.parse(fs.readFileSync(path.join(__dirname, 'account.dat'), 'utf8'));

test.describe('TC_020 Role Based Access Control - Regulation List Updates Sort', () => {
  
  for (const user of users) {
    
    test(`Perform regulation list updates sort test for ${user.id}`, async ({ page }) => {
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

      // Step 2: Perform search loop
      await test.step('Perform Search 5 times or until screenshot is taken', async () => {
        let screenshotTaken = false;
        for (let i = 1; i <= 5; i++) {
            console.log(`
--- Repetition ${i}/5 for user ${user.id} ---`);
            
            await page.getByRole('button', { name: 'Content' }).click();
            await page.getByRole('menuitem', { name: 'ECN Regulation List Updates' }).click();

            let mainFrameElement = page.locator('iframe[name="legacy-outlet"]');
            await mainFrameElement.waitFor({ state: 'attached', timeout: 30000 });
            let mainFrame = await mainFrameElement.contentFrame();
            if (!mainFrame) throw new Error('Could not access main iframe content.');

            // Verify page name
            const pageNameLabel = mainFrame.locator('#lbxPageName');
            await expect(pageNameLabel).toContainText('Regulation List Updates', { timeout: 20000 });
            console.log('SUCCESS: Verified page name contains "Regulation List Updates".');

            // Step: Expand the first regulation node
            console.log('INFO: Expanding the first regulation node...');
            const firstExpandButton = mainFrame.locator('input.rgExpand[title="Expand"]').first();
            await expect(firstExpandButton).toBeVisible({ timeout: 10000 });
            await firstExpandButton.click();
            await page.waitForTimeout(3000); // Wait for expansion and grid update

            // Step: Expand the child node (Source Date)
            // After expansion, the first button usually changes to 'Collapse' or stays but reveals new 'Expand' buttons below it.
            // We look for the first visible 'Expand' button again. 
            // If the first one became "Collapse", the next "Expand" button should be the child.
            console.log('INFO: Expanding the child node (Source Date)...');
            const nextExpandButton = mainFrame.locator('input.rgExpand[title="Expand"]').first();
            await expect(nextExpandButton).toBeVisible({ timeout: 10000 });
            await nextExpandButton.click();
            
            // Wait for update
            await page.waitForTimeout(5000);

            // Take screenshot
            const screenshotBuffer = await page.screenshot({ fullPage: true });
            await test.info().attach(`regulation_list_updates_expanded_${user.id}.png`, {
                body: screenshotBuffer,
                contentType: 'image/png'
            });

            console.log('SUCCESS: Expanded nodes and screenshot attached. Test complete.');
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

