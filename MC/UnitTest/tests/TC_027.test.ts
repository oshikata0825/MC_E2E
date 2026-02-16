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

// Keyword list for TC_027
const keywordList = ['system', 'sensor', 'engine', 'laser', 'telecommunications', 'security', 'equipment'];

test.describe('TC_027 Role Based Access Control - Global Tariffs (Quick Lookup) Keyword Search', () => {
  
  for (const user of users) {
    
    test(`Perform Global Tariffs (Quick Lookup) keyword search test for ${user.id}`, async ({ page }) => {
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

      // Step 2: Access Global Tariffs (Quick Lookup) and Verify
      await test.step('Access Global Tariffs (Quick Lookup) and Verify', async () => {
        let screenshotTaken = false;
        for (let i = 1; i <= 5; i++) {
            console.log(`\n--- Repetition ${i}/5 for user ${user.id} ---`);
            
            await page.getByRole('button', { name: 'Content' }).click();
            // Use regex for exact match
            await page.locator('a[role="menuitem"]').filter({ hasText: /^Global Tariffs \(Quick Lookup\)$/ }).click();

            let mainFrameElement = page.locator('iframe[name="legacy-outlet"]');
            await mainFrameElement.waitFor({ state: 'attached', timeout: 30000 });
            let mainFrame = await mainFrameElement.contentFrame();
            if (!mainFrame) throw new Error('Could not access main iframe content.');

            // Verify page name
            const pageNameLabel = mainFrame.locator('#lbxPageName');
            await expect(pageNameLabel).toContainText('Global Tariffs (Quick Lookup)', { timeout: 20000 });
            console.log('SUCCESS: Verified page name contains "Global Tariffs (Quick Lookup)".');

            // Select a random search criteria (only for Profile)
            const criteria = searchCriteria[Math.floor(Math.random() * searchCriteria.length)];
            const keyword = keywordList[Math.floor(Math.random() * keywordList.length)];
            console.log(`Using criteria: Profile='${criteria.profile}', Keyword='${keyword}'`);

            // Step 2a: Select Tariff Profile
            const profileInput = mainFrame.locator('#ctl00_MainContent_cmbxStatusBarTariffSchedule_Input');
            await profileInput.click();
            await profileInput.fill(criteria.profile);
            await page.waitForTimeout(1000); // Wait for input to process
            await page.keyboard.press('Enter');
            
            // WCO profile seems to trigger a heavy reload/postback. Waiting significantly longer.
            await page.waitForTimeout(8000); 

            // Ensure previous autocomplete is gone
            await expect(mainFrame.locator('.rcbSlide').first()).toBeHidden({ timeout: 5000 });

            // Step 2b: Enter Keyword and Select from Autocomplete (Index > 0)
            const hsInput = mainFrame.locator('#ctl00_MainContent_cmbxStatusBarHSNumber_Input');
            
            // Wait for input to be editable
            await expect(hsInput).toBeEditable({ timeout: 10000 });
            await hsInput.click();
            await page.waitForTimeout(2000); // Wait for focus and event handlers

            // Use pressSequentially to ensure events are triggered for autocomplete
            await hsInput.pressSequentially(keyword, { delay: 150 });
            await page.waitForTimeout(3000); // Wait for autocomplete suggestions to appear

            // Select from autocomplete list (MUST NOT select the first one)
            const visibleItems = mainFrame.locator('.rcbSlide .rcbList li:visible');
            await expect(visibleItems.first()).toBeVisible({ timeout: 10000 });
            
            const count = await visibleItems.count();
            console.log(`Autocomplete items found: ${count}`);

            if (count < 2) {
                console.log('Not enough items to skip the first one. Skipping this iteration.');
                // In a real test, we might want to retry or fail. For now, we continue to see if next loop works?
                // Or better, just select the only one and log a warning, but user said "MUST select other than 1st".
                // Let's throw error to fail fast if this happens, or retry logic.
                // Assuming keywords provide ample results.
                throw new Error('Less than 2 autocomplete results found, cannot select non-first item.');
            }

            // Select randomly from index 1 to count-1
            // e.g. count=3, indices 0,1,2. We want 1 or 2.
            // Math.random() * (max - min + 1) + min
            // min=1, max=count-1. 
            // range = (count-1) - 1 + 1 = count - 1.
            const indexToSelect = Math.floor(Math.random() * (count - 1)) + 1;
            
            const autocompleteItem = visibleItems.nth(indexToSelect);
            const selectedText = await autocompleteItem.textContent();
            console.log(`Selecting autocomplete item index ${indexToSelect} (Total: ${count}): ${selectedText?.trim()}`);

            await autocompleteItem.click();

            // Step 2c: Verify HS Code Conversion and Search Results
            await page.waitForTimeout(2000); // Wait for value update/conversion

            // Capture screenshot immediately to check HS code conversion
            const screenshotBufferVerify = await page.screenshot({ fullPage: true });
            await test.info().attach(`global_tariffs_keyword_converted_${user.id}.png`, {
                body: screenshotBufferVerify,
                contentType: 'image/png'
            });

            // Log current value in the input
            const inputValue = await hsInput.inputValue();
            console.log(`Value in HS Code input after selection: '${inputValue}'`);

            // Step 2d: Click Search Button and Verify Results
            const searchButton = mainFrame.locator('#ctl00_MainContent_imgStatusBarSearch');
            await expect(searchButton).toBeVisible({ timeout: 10000 });
            await searchButton.click();
            console.log('Clicked Search button.');

            // Check for the presence of the result tab strip
            const resultTabStrip = mainFrame.locator('#ctl00_MainContent_tabctnData');
            await expect(resultTabStrip).toBeVisible({ timeout: 30000 });
            console.log('SUCCESS: Search results displayed.');

            // Wait for page to fully render
            await page.waitForTimeout(3000);

            // Take final screenshot
            const screenshotBufferResult = await page.screenshot({ fullPage: true });
            await test.info().attach(`global_tariffs_keyword_result_${user.id}.png`, {
                body: screenshotBufferResult,
                contentType: 'image/png'
            });

            console.log('SUCCESS: Test complete.');
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

