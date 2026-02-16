import { test, expect } from '@playwright/test';
import { BASE_URL } from './config';
import * as fs from 'fs';
import * as path from 'path';

// Load user definitions
const users = JSON.parse(fs.readFileSync(path.join(__dirname, 'account.dat'), 'utf8'));

// Define scenarios
const scenarios = [
    {
        regulation: "CN - CHINA - China Export Control List of Dual-Use Items",
        ecnCodes: ['1C902'],
        hsCodes: ['2805', '2846']
    },
    {
        regulation: "CN - CHINA - China Export License Management Catalogue of Dual-use Items and Technologies",
        ecnCodes: ['1.1.2', '1.1.3'],
        hsCodes: ['2801', '2804']
    },
    {
        regulation: "JP - JAPAN - Japan Export Control Order",
        ecnCodes: ['1.10', '1.11', '1.12', '1.13', '1.14'],
        hsCodes: []
    },
    {
        regulation: "JP - JAPAN - Japan Export Control Order: Appended Table 2-3 (Prohibited Items for Export to Russia etc.)",
        ecnCodes: ['1.2'],
        hsCodes: []
    },
    {
        regulation: "JP - JAPAN - Japan Foreign Exchange Order",
        ecnCodes: ['10.1', '10.2', '10.3', '10.4'],
        hsCodes: []
    },
    {
        regulation: "US - UNITED STATES OF AMERICA - EAR/Commerce Control List",
        ecnCodes: ['1', '2', '3', '4', '5', '6', '7', '8'],
        hsCodes: []
    },
    {
        regulation: "US - UNITED STATES OF AMERICA - International Traffic in Arms Regulations (ITAR)",
        ecnCodes: ['4', '5', '6', '7'],
        hsCodes: []
    },
    {
        regulation: "US - UNITED STATES OF AMERICA - Supplement No. 6 to Part 746 - Items That Require a License For Russia and Belarus",
        ecnCodes: ['7', '8', '9'],
        hsCodes: []
    }
];

test.describe('TC_018 Role Based Access Control - Specific Regulation Search', () => {
  
  for (const user of users) {
    
    test(`Perform search with specific regulation for ${user.id}`, async ({ page }) => {
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
            console.log(`\n--- Repetition ${i}/5 for user ${user.id} ---`);
            
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

            // Take screenshot
            const screenshotBuffer = await page.screenshot({ fullPage: true });
            await test.info().attach(`regulation_list_updates_verified_${user.id}.png`, {
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

