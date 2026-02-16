import { test, expect } from '@playwright/test';
import { BASE_URL } from './config';
import * as fs from 'fs';
import * as path from 'path';

// Load user definitions
const users = JSON.parse(fs.readFileSync(path.join(__dirname, 'account.dat'), 'utf8'));

// Define search criteria for random selection (Profile only)
const searchCriteria = [
    { profile: 'WCO - World Customs Organization - Global HS (IMPORT)', hsCode: '8501' },
    { profile: 'JP - JAPAN - Japan Export Tariff (EXPORT)', hsCode: '0101' },
    { profile: 'JP - JAPAN - Japans Tariff Schedule (IMPORT)', hsCode: '8471' },
    { profile: 'US - UNITED STATES OF AMERICA - US HTS (IMPORT and EXPORT)', hsCode: '8542' },
    { profile: 'US - UNITED STATES OF AMERICA - US Schedule B (EXPORT)', hsCode: '8802' }
];

test.describe('TC_034 Role Based Access Control - Global Tariffs Search', () => {
  test.describe.configure({ mode: 'parallel' });
  
  users.forEach((user: any) => {
    
    test(`Perform Global Tariffs search test for ${user.id}`, async ({ page }) => {
      test.setTimeout(180000);

      console.log(`--- Test Start: ${user.id} ---`);

      // Step 1: Login
      await test.step('Login', async () => {
        await page.goto(BASE_URL + '/Logon.aspx?ReturnUrl=%2fgtm%2fhome');
        await page.getByRole('textbox', { name: 'Company' }).fill(user.company);
        await page.getByRole('textbox', { name: 'Username' }).fill(user.id);
        await page.getByRole('textbox', { name: 'Password' }).fill(user.pass);
        await page.getByRole('link', { name: 'Login' }).click();
      });

      // Step 2: Access Global Tariffs and Search
      await test.step('Access Global Tariffs and Search', async () => {
        
        await page.getByRole('button', { name: 'Content' }).click();
        // Use regex for exact match 'Global Tariffs'
        await page.locator('a[role="menuitem"]').filter({ hasText: /^Global Tariffs$/ }).click();

        let mainFrameElement = page.locator('iframe[name="legacy-outlet"]');
        await mainFrameElement.waitFor({ state: 'attached', timeout: 30000 });
        let mainFrame = await mainFrameElement.contentFrame();
        if (!mainFrame) throw new Error('Could not access main iframe content.');

        // Verify page name (Optional but good practice)
        // Assuming #lbxPageName exists
        try {
            const pageNameLabel = mainFrame.locator('#lbxPageName');
            await expect(pageNameLabel).toContainText('Global Tariffs', { timeout: 10000 });
            console.log('SUCCESS: Verified page name.');
        } catch (e) {
            console.warn('Page name verification skipped or failed (non-critical).');
        }

        // Re-acquire frame to be safe
        mainFrameElement = page.locator('iframe[name="legacy-outlet"]');
        mainFrame = await mainFrameElement.contentFrame();
        if (!mainFrame) throw new Error('Could not access main iframe content.');

        // Handle "Global Tariffs Search Profile" modal if present (Initial check)
        console.log('INFO: Checking for "Global Tariffs Search Profile" modal...');
        const modalSelector = '.rwCloseButton'; 
        await page.waitForTimeout(5000);
        const iframeCloseBtn = mainFrame.locator(modalSelector).first();
        if (await iframeCloseBtn.isVisible({ timeout: 5000 })) { 
            console.log('INFO: Modal close button found in IFRAME. Clicking...');
            await iframeCloseBtn.click({ force: true });
            await page.waitForTimeout(2000); 
        } else {
             const pageCloseBtn = page.locator(modalSelector).first();
             if (await pageCloseBtn.isVisible({ timeout: 2000 })) {
                 console.log('INFO: Modal close button found in MAIN PAGE. Clicking...');
                 await pageCloseBtn.click({ force: true });
                 await page.waitForTimeout(2000);
             }
        }

        // Loop to find a record with DUTY tab
        let validDutyFound = false;
        const maxRetries = 5;

        for (let attempt = 1; attempt <= maxRetries; attempt++) {
            console.log(`\n--- Attempt ${attempt}/${maxRetries} to find DUTY tab ---`);
            
            // Select a random search criteria (Profile)
            const criteria = searchCriteria[Math.floor(Math.random() * searchCriteria.length)];
            console.log(`Using criteria: Profile='${criteria.profile}', HSCode='${criteria.hsCode}'`);

            // Step 2a: Select Tariff Profile
            const profileInput = mainFrame.locator('#ctl00_MainContent_cmbxTariffSchedule_Input');
            await profileInput.click();
            await profileInput.fill(criteria.profile);
            await page.waitForTimeout(1000); // Wait for input
            await page.keyboard.press('Enter');
            
            // Wait for potential postback/reload
            await page.waitForTimeout(5000); 

            // Step 2b: Enter HS Code and Select from Autocomplete
            const hsInput = mainFrame.locator('#ctl00_MainContent_txtbxHSNumberFilter');
            await expect(hsInput).toBeVisible();
            await hsInput.click();
            await hsInput.clear();
            console.log(`Entering HS Code: ${criteria.hsCode} slowly...`);
            await hsInput.pressSequentially(criteria.hsCode, { delay: 300 });

            // Wait for autocomplete suggestions
            console.log('Waiting for autocomplete...');
            await page.waitForTimeout(3000);

            // Select using Keyboard: ArrowDown + Enter
            await page.keyboard.press('ArrowDown');
            await page.waitForTimeout(500);
            await page.keyboard.press('Enter');
            console.log('Selected autocomplete item.');
            
            await page.waitForTimeout(2000); 
            let selectedHSValue = await hsInput.inputValue();
            console.log(`INFO: Current HS Value: "${selectedHSValue}"`);

            if (!selectedHSValue || selectedHSValue.trim() === '') {
                console.warn('HS Code input is empty. Refilling...');
                await hsInput.fill(criteria.hsCode);
            }

            // Explicitly click Search button to ensure search executes
            console.log('Clicking Search button...');
            const searchButton = mainFrame.locator('#ctl00_MainContent_lnxbtnNewSearch');
            
            // Wait for any loading overlay to disappear
            await expect(page.locator('.IPModal, .bento-busyloader-inner').first()).not.toBeVisible({ timeout: 10000 }).catch(() => {});
            
            await searchButton.click({ force: true });

            // Wait for loading to start (short timeout as it might be quick)
            try {
                await expect(page.locator('.IPModal, .bento-busyloader-inner').first()).toBeVisible({ timeout: 3000 });
            } catch (e) {
                console.log('Loading overlay did not appear or appeared too quickly.');
            }
            
            // Wait for loading to finish
            await expect(page.locator('.IPModal, .bento-busyloader-inner').first()).not.toBeVisible({ timeout: 30000 });

            // Wait for search results
            console.log('Waiting for search results to load...');
            const resultsGrid = mainFrame.locator('#ctl00_MainContent_RadGrid1');
            try {
                await expect(resultsGrid).toBeVisible({ timeout: 30000 });
                console.log('SUCCESS: Search results grid displayed.');
            } catch (e) {
                console.warn('Search results grid not found. Retrying...');
                continue;
            }

            // Wait for rows to populate
            const rows = resultsGrid.locator('tr.rgRow, tr.rgAltRow');
            try {
                await expect(async () => {
                    const count = await rows.count();
                    expect(count).toBeGreaterThan(0);
                }).toPass({ timeout: 10000 });
            } catch (e) {
                console.warn('Rows did not appear in time. Retrying...');
                continue;
            }

            // Find the first actual link in the grid (some rows might be headers/text only)
            const targetLink = resultsGrid.locator('a.IPGridLink').first();
            
            // Ensure visible and log text
            if (await targetLink.isVisible()) {
                 const text = await targetLink.innerText();
                 console.log(`Target link text: ${text}`);
            } else {
                 console.log('Link reported not visible. Attempting JS scroll...');
            }

            // Force scroll using JS on the element
            // This usually handles the parent container scroll automatically
            await targetLink.evaluate((el: any) => {
                el.scrollIntoView({ behavior: 'auto', block: 'center', inline: 'center' });
            });
            await page.waitForTimeout(1000); // Wait for scroll to finish

            console.log('Clicking result link...');
            
            // Wait for loading overlay again before clicking link
            await expect(page.locator('.IPModal, .bento-busyloader-inner').first()).not.toBeVisible({ timeout: 10000 }).catch(() => {});
            
            await targetLink.click({ force: true });
            await page.waitForTimeout(3000); // Wait for popup

            // DEBUG SCREENSHOT: After click
            const postClickScreenshot = await page.screenshot({ fullPage: true });
            await test.info().attach(`debug_post_click_attempt_${attempt}_${user.id}.png`, {
                body: postClickScreenshot,
                contentType: 'image/png'
            });

            // Wait for Detail Modal (Nested Iframe)
            // Look for iframe with src containing 'fugGlobalTariffsDetail.aspx'
            const detailFrameElement = mainFrame.locator('iframe[src*="fugGlobalTariffsDetail.aspx"]');
            try {
                await detailFrameElement.waitFor({ state: 'attached', timeout: 30000 });
            } catch (e) {
                console.warn('Detail frame not found or timed out. Retrying...');
                continue;
            }
            
            const detailFrame = await detailFrameElement.contentFrame();
            if (!detailFrame) {
                console.warn('Could not access detail frame content. Retrying...');
                continue;
            }

            // Check Tabs for 'DUTY'
            // Wait for the main tabstrip that contains 'Header'
            const mainTabStrip = detailFrame.locator('ul.rtsUL').filter({ hasText: 'Header' }).first();
            await mainTabStrip.waitFor({ state: 'visible', timeout: 20000 });
            
            const tabs = detailFrame.locator('.rtsTxt');
            const tabTexts = await tabs.allInnerTexts();
            console.log('Tabs found:', tabTexts);

            // Case insensitive check
            const dutyTabIndex = tabTexts.findIndex(t => t.toUpperCase() === 'DUTY');
            if (dutyTabIndex !== -1) {
                console.log('SUCCESS: "DUTY" tab found. Clicking it...');
                
                const dutyTab = tabs.nth(dutyTabIndex);
                await dutyTab.click();
                
                // Wait for content load (tab selection)
                await page.waitForTimeout(5000); 
                
                validDutyFound = true;
                
                // Take Screenshot (of the detail view with DUTY tab selected)
                const screenshotBufferResult = await page.screenshot({ fullPage: true });
                await test.info().attach(`global_tariffs_duty_tab_${user.id}.png`, {
                    body: screenshotBufferResult,
                    contentType: 'image/png'
                });
                console.log('SUCCESS: Screenshot attached. Test complete.');

                break;
            } else {
                console.log('WARNING: "DUTY" tab NOT found. Closing modal and retrying...');
                
                // Close the detail modal
                // The close button is likely on the RadWindow wrapper in the mainFrame (parent of detailFrame)
                // We look for the visible close button in mainFrame that corresponds to the topmost window
                const closeButtons = mainFrame.locator('.rwCloseButton');
                // We click the last one as it's likely the newest/topmost
                if (await closeButtons.count() > 0) {
                     await closeButtons.last().click();
                     await page.waitForTimeout(5000); // Wait for close
                } else {
                    console.warn('Could not find close button for detail modal. Reloading page...');
                    await page.reload();
                    // Need to navigate back to Global Tariffs... complicated. 
                    // Let's hope close button works.
                }

                // Close Search Result Tab(s) to reset state
                const closeTabBtn = mainFrame.locator('input[src*="x-red-16x16.png"]');
                const tabCount = await closeTabBtn.count();
                if (tabCount > 0) {
                    console.log(`Closing ${tabCount} search result tab(s) to reset...`);
                    // Click them one by one. Re-querying or clicking first repeatedly.
                    for (let k = 0; k < tabCount; k++) {
                        // Use evaluate to avoid visibility issues or just click
                        try {
                            await closeTabBtn.first().click();
                            await page.waitForTimeout(1000);
                        } catch (e) {
                            console.warn('Error closing tab:', e);
                        }
                    }
                }
            }
        }

        if (!validDutyFound) throw new Error('Could not find a record with DUTY tab after max retries.');
      });

      // Step 3: Logout
      await test.step('Logout', async () => {
        const userButtonName = `MC Test${user.id.substring(6)}`;
        await page.bringToFront();

        const userButton = page.getByRole('button', { name: new RegExp(`^${userButtonName}`)
        });
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

