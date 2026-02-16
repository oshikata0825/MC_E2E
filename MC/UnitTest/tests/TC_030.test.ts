import { test, expect } from '@playwright/test';
import { BASE_URL } from './config';
import * as fs from 'fs';
import * as path from 'path';

// Load user definitions
const users = JSON.parse(fs.readFileSync(path.join(__dirname, 'account.dat'), 'utf8'));

// Define search criteria for random selection
const searchCriteria = [
    { profile: 'WCO - World Customs Organization - Global HS (IMPORT)', hsCode: '8501' },
    { profile: 'JP - JAPAN - Japan Export Tariff (EXPORT)', hsCode: '0101' },
    { profile: 'JP - JAPAN - Japans Tariff Schedule (IMPORT)', hsCode: '8471' },
    { profile: 'US - UNITED STATES OF AMERICA - US HTS (IMPORT and EXPORT)', hsCode: '8542' },
    { profile: 'US - UNITED STATES OF AMERICA - US Schedule B (EXPORT)', hsCode: '8802' }
];

test.describe('TC_030 Role Based Access Control - Global Tariffs (Classic) Search', () => {
  test.describe.configure({ mode: 'parallel' });
  
  users.forEach((user: any) => {
    
    test(`Perform Global Tariffs (Classic) search test for ${user.id}`, async ({ page }) => {
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

      // Step 2: Access Global Tariffs (Classic) and Perform Search
      await test.step('Access Global Tariffs (Classic) and Search', async () => {
        let screenshotTaken = false;
        for (let i = 1; i <= 5; i++) {
            console.log(`\n--- Repetition ${i}/5 for user ${user.id} ---`);
            
            await page.getByRole('button', { name: 'Content' }).click();
            // Use regex for exact match
            await page.locator('a[role="menuitem"]').filter({ hasText: /^Global Tariffs \(Classic\)$/ }).click();

            let mainFrameElement = page.locator('iframe[name="legacy-outlet"]');
            await mainFrameElement.waitFor({ state: 'attached', timeout: 30000 });
            let mainFrame = await mainFrameElement.contentFrame();
            if (!mainFrame) throw new Error('Could not access main iframe content.');

            // Verify page name
            const pageNameLabel = mainFrame.locator('#lbxPageName');
            await expect(pageNameLabel).toContainText('Global Tariffs (Classic)', { timeout: 20000 });
            console.log('SUCCESS: Verified page name contains "Global Tariffs (Classic)".');

            // Select a random search criteria
            const criteria = searchCriteria[Math.floor(Math.random() * searchCriteria.length)];
            console.log(`Using criteria: Profile='${criteria.profile}', HS Code='${criteria.hsCode}'`);

            // Step 2a: Select Tariff Profile
            const profileInput = mainFrame.locator('#ctl00_MainContent_cmbxTariffSchedule_Input');
            await profileInput.click();
            await profileInput.fill(criteria.profile);
            await page.waitForTimeout(1000); // Wait for input to process
            await page.keyboard.press('Enter');
            
            // Wait for potential postback/reload
            await page.waitForTimeout(12000); 

            // Step 2b: Enter HS Code and Select from Autocomplete
            const hsInput = mainFrame.locator('#txtbxHSNumber');
            
            // Wait for input to be editable
            await expect(hsInput).toBeEditable({ timeout: 10000 });
            await hsInput.click();
            await page.waitForTimeout(2000); // Wait for focus

            // Use pressSequentially with delay
            await hsInput.pressSequentially(criteria.hsCode, { delay: 300 });
            await page.waitForTimeout(3000); // Wait for autocomplete suggestions

            // Select using Keyboard since element location is unstable
            // Press ArrowDown 2 to 4 times (randomly) to skip the first item (header or first result)
            const arrowDownCount = Math.floor(Math.random() * 3) + 2; // 2, 3, or 4
            console.log(`Pressing ArrowDown ${arrowDownCount} times to select autocomplete item.`);
            
            for (let k = 0; k < arrowDownCount; k++) {
                await page.keyboard.press('ArrowDown');
                await page.waitForTimeout(300);
            }
            await page.keyboard.press('Enter');

            // Step 2c: Verify Search Results
            const resultTree = mainFrame.locator('#ctl00_MainContent_rtvHSHierarchy');
            await expect(resultTree).toBeVisible({ timeout: 30000 });
            
            // Verify that the tree content somehow relates to the HS code (basic check)
            // Ideally we check for .rtSelected or specific text, but tree structure is complex.
            // Just verifying visibility and taking screenshot as requested.
            console.log('SUCCESS: Search results tree displayed.');

            // Wait for page to fully render
            await page.waitForTimeout(5000);

            // Take screenshot
            const screenshotBuffer = await page.screenshot({ fullPage: true });
            await test.info().attach(`global_tariffs_classic_result_${user.id}.png`, {
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
  });
});

