import { test, expect } from '@playwright/test';
import { BASE_URL } from './config';
import * as fs from 'fs';
import * as path from 'path';

// Load user definitions
const users = JSON.parse(fs.readFileSync(path.join(__dirname, 'account.dat'), 'utf8'));

// Define profiles for random selection
const profiles = [
    'WCO - World Customs Organization - Global HS (IMPORT)',
    'JP - JAPAN - Japan Export Tariff (EXPORT)',
    'JP - JAPAN - Japans Tariff Schedule (IMPORT)',
    'US - UNITED STATES OF AMERICA - US HTS (IMPORT and EXPORT)',
    'US - UNITED STATES OF AMERICA - US Schedule B (EXPORT)'
];

// Define keywords for random selection
const keywords = ['system', 'sensor', 'engine', 'laser', 'telecommunications', 'security', 'equipment'];

test.describe('TC_023 Role Based Access Control - Global Tariffs Page Access', () => {
  
  for (const user of users) {
    
    test(`Perform Global Tariffs page access test for ${user.id}`, async ({ page }) => {
      test.setTimeout(210000);

      // Select random profile and keyword
      const selectedProfile = profiles[Math.floor(Math.random() * profiles.length)];
      const selectedKeyword = keywords[Math.floor(Math.random() * keywords.length)];
      
      console.log(`--- Test Start: ${user.id} ---`);
      console.log(`INFO: Selected Profile: "${selectedProfile}", Keyword: "${selectedKeyword}"`);

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

            // Wait for potential modal appearance
            await page.waitForTimeout(10000);

            // Handle "Global Tariffs Search Profile" modal if present
            // Check both in iframe and main page as Telerik windows can be appended to body
            console.log('INFO: Checking for "Global Tariffs Search Profile" modal...');
            
            const modalSelector = '.rwCloseButton'; // Common class for Telerik close buttons
            const iframeCloseBtn = mainFrame.locator(modalSelector).first();

            if (await iframeCloseBtn.isVisible({ timeout: 30000 })) {
                console.log('INFO: Modal close button found in IFRAME. Clicking...');
                await iframeCloseBtn.click({ force: true });
                await page.waitForTimeout(2000); // Wait for animation
            } else {
                // Check in main page if not found in iframe
                const pageCloseBtn = page.locator(modalSelector).first();
                if (await pageCloseBtn.isVisible({ timeout: 5000 })) {
                    console.log('INFO: Modal close button found in MAIN PAGE. Clicking...');
                    await pageCloseBtn.click({ force: true });
                    await page.waitForTimeout(2000); // Wait for animation
                } else {
                    console.log('INFO: Modal close button NOT found. Assuming no modal.');
                }
            }

            // Step: Select Tariff Profile
            // ID: ctl00_MainContent_cmbxTariffSchedule_Input
            console.log(`INFO: Selecting Tariff Profile: ${selectedProfile}...`);
            const profileInput = mainFrame.locator('#ctl00_MainContent_cmbxTariffSchedule_Input');
            await profileInput.click();
            await profileInput.fill(selectedProfile);
            await page.keyboard.press('Enter');
            await page.waitForTimeout(1000); // Wait for selection to apply

            // Step: Enter Keyword (using HS Number Filter input as directed)
            // ID: ctl00_MainContent_txtbxHSNumberFilter
            console.log(`INFO: Entering Keyword "${selectedKeyword}"...`);
            const hsInput = mainFrame.locator('#ctl00_MainContent_txtbxHSNumberFilter');
            await hsInput.click();
            await hsInput.fill(selectedKeyword);

            // Wait for autocomplete and select the first option
            console.log('INFO: Waiting for autocomplete suggestions...');
            await page.waitForTimeout(2000); // Wait for suggestion list to appear
            
            // Try to find autocomplete item. 
            // Telerik often uses .rcbList .rcbItem or similar. 
            // We search for an element containing the HS Code text in a list-like structure or just the first visible suggestion.
            // Assuming it might be a standard Telerik RadComboBox dropdown or similar list
            const suggestion = mainFrame.locator('.rcbList .rcbItem, .ui-menu-item, li[role="option"]').first();
            
            if (await suggestion.isVisible({ timeout: 5000 })) {
                 console.log('INFO: Autocomplete suggestion found. Clicking it...');
                 await suggestion.click();
                 await page.waitForTimeout(1000); // Wait for selection
            } else {
                 console.log('WARNING: Autocomplete suggestion not found via mouse. Trying keyboard selection (ArrowDown + Enter)...');
                 await hsInput.focus(); // Ensure focus is on input
                 await page.waitForTimeout(500);
                 await page.keyboard.press('ArrowDown');
                 await page.waitForTimeout(500);
                 await page.keyboard.press('Enter');
                 await page.waitForTimeout(1000); // Wait for selection
            }

            // Step: Click Search
            // ID: ctl00_MainContent_lnxbtnNewSearch
            console.log('INFO: Clicking Search button...');
            const searchButton = mainFrame.locator('#ctl00_MainContent_lnxbtnNewSearch');
            await searchButton.click();

            // Wait for results
            console.log('INFO: Waiting for search results (tab strip) to appear...');
            const resultTabStrip = mainFrame.locator('ul.rtsUL, div.rtsLevel');
            await resultTabStrip.first().waitFor({ state: 'visible', timeout: 30000 });
            
            // Wait extra time for content to settle
            await page.waitForTimeout(3000);

            // Take screenshot
            const screenshotBuffer = await page.screenshot({ fullPage: true });
            await test.info().attach(`global_tariffs_search_result_${user.id}.png`, {
                body: screenshotBuffer,
                contentType: 'image/png'
            });

            console.log('SUCCESS: Search performed and screenshot attached. Test complete.');
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

