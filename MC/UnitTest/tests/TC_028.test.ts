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

test.describe('TC_028 Role Based Access Control - Global Tariffs (Quick Lookup) Page Access and Tab Check', () => {
  test.describe.configure({ mode: 'parallel' });
  
  users.forEach((user: any) => {
    
    test(`Perform Global Tariffs (Quick Lookup) page access test for ${user.id}`, async ({ page }) => {
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

            // Select a random search criteria
            const criteria = searchCriteria[Math.floor(Math.random() * searchCriteria.length)];
            // const criteria = searchCriteria[0]; // Force WCO profile for debugging
            console.log(`Using criteria: Profile='${criteria.profile}', HS Code='${criteria.hsCode}'`);

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

            // Step 2b: Enter HS Code and Select from Autocomplete
            const hsInput = mainFrame.locator('#ctl00_MainContent_cmbxStatusBarHSNumber_Input');
            
            // Wait for input to be editable
            await expect(hsInput).toBeEditable({ timeout: 10000 });
            await hsInput.click();
            await page.waitForTimeout(2000); // Wait for focus and event handlers

            // Use pressSequentially to ensure events are triggered for autocomplete
            await hsInput.pressSequentially(criteria.hsCode, { delay: 300 });
            await page.waitForTimeout(3000); // Wait for autocomplete suggestions to appear

            // Select from autocomplete list with random logic
            const visibleItems = mainFrame.locator('.rcbSlide .rcbList li:visible');
            await expect(visibleItems.first()).toBeVisible({ timeout: 10000 });
            
            const count = await visibleItems.count();
            let indexToSelect = 0;
            if (count >= 3) {
                indexToSelect = Math.floor(Math.random() * 3); // 0, 1, or 2
            } else if (count === 2) {
                indexToSelect = Math.floor(Math.random() * 2); // 0 or 1
            }
            
            const autocompleteItem = visibleItems.nth(indexToSelect);
            const selectedText = await autocompleteItem.textContent();
            console.log(`Selecting autocomplete item index ${indexToSelect} (Total: ${count}): ${selectedText?.trim()}`);

            await autocompleteItem.click();

            // Step 2c: Verify Search Results and Iterate Tabs
            try {
                // Check for the presence of the result tab strip
                const resultTabStrip = mainFrame.locator('#ctl00_MainContent_tabctnData');
                await expect(resultTabStrip).toBeVisible({ timeout: 30000 });
                console.log('SUCCESS: Search results displayed.');

                // Get all tabs
                const tabs = resultTabStrip.locator('.rtsLI a');
                const tabCount = await tabs.count();
                console.log(`Found ${tabCount} tabs.`);

                for (let j = 0; j < tabCount; j++) {
                    const tab = tabs.nth(j);
                    const tabName = await tab.innerText();
                    console.log(`\n--- Processing Tab ${j + 1}/${tabCount}: ${tabName} ---`);

                    // Click tab
                    await tab.click();
                    
                    // Wait for selection and rendering
                    await expect(tab).toHaveClass(/rtsSelected/, { timeout: 10000 });
                    await page.waitForTimeout(5000); // Allow content to load

                    // Take screenshot
                    const screenshotBuffer = await page.screenshot({ fullPage: true });
                    await test.info().attach(`global_tariffs_quick_lookup_tab_${j + 1}_${tabName.replace(/[^a-zA-Z0-9]/g, '_')}_${user.id}.png`, {
                        body: screenshotBuffer,
                        contentType: 'image/png'
                    });
                    console.log(`SUCCESS: Screenshot attached for tab: ${tabName}`);
                }

                console.log('SUCCESS: All tabs processed. Test complete.');
                break; // Exit loop on success
            } catch (e) {
                console.warn(`WARNING: Failed to display search results or process tabs on repetition ${i}. Error: ${e}`);
                
                // Take failure screenshot
                const screenshotBuffer = await page.screenshot({ fullPage: true });
                await test.info().attach(`search_results_failed_rep_${i}_${user.id}.png`, {
                    body: screenshotBuffer,
                    contentType: 'image/png'
                });
                
                // Continue to next repetition
                continue;
            }
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

