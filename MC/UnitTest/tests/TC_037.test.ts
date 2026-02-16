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

test.describe('TC_037 Role Based Access Control - Global Tariffs Search - Country level Controls', () => {
  test.describe.configure({ mode: 'parallel' });
  
  users.forEach((user: any) => {
    
    test(`Perform Global Tariffs search (Country level Controls) test for ${user.id}`, async ({ page }) => {
      test.setTimeout(300000);

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
        await page.locator('a[role="menuitem"]').filter({ hasText: /^Global Tariffs$/ }).click();

        let mainFrameElement = page.locator('iframe[name="legacy-outlet"]');
        await mainFrameElement.waitFor({ state: 'attached', timeout: 30000 });
        let mainFrame = await mainFrameElement.contentFrame();
        if (!mainFrame) throw new Error('Could not access main iframe content.');

        try {
            const pageNameLabel = mainFrame.locator('#lbxPageName');
            await expect(pageNameLabel).toContainText('Global Tariffs', { timeout: 10000 });
            console.log('SUCCESS: Verified page name.');
        } catch (e) {
            console.warn('Page name verification skipped or failed (non-critical).');
        }

        // Handle "Global Tariffs Search Profile" modal
        console.log('INFO: Checking for "Global Tariffs Search Profile" modal...');
        const modalSelector = '.rwCloseButton'; 
        await page.waitForTimeout(5000);
        const iframeCloseBtn = mainFrame.locator(modalSelector).first();
        if (await iframeCloseBtn.isVisible({ timeout: 5000 })) { 
            console.log('INFO: Modal close button found in IFRAME. Waiting 3s before clicking...');
            await page.waitForTimeout(3000);
            await iframeCloseBtn.click({ force: true });
            await page.waitForTimeout(2000); 
        } else { 
             const pageCloseBtn = page.locator(modalSelector).first();
             if (await pageCloseBtn.isVisible({ timeout: 2000 })) {
                 console.log('INFO: Modal close button found in MAIN PAGE. Waiting 3s before clicking...');
                 await page.waitForTimeout(3000);
                 await pageCloseBtn.click({ force: true });
                 await page.waitForTimeout(2000);
             }
        }

        // Loop to find a record with REGULATORY CONTROLS tab
        let validTargetFound = false;
        const maxRetries = 5;

        for (let attempt = 1; attempt <= maxRetries; attempt++) {
            console.log(`
--- Attempt ${attempt}/${maxRetries} to find CONTROLS / Country level Controls tab ---`);
            
            // Refresh frame reference to avoid stale element errors
            mainFrameElement = page.locator('iframe[name="legacy-outlet"]');
            await mainFrameElement.waitFor({ state: 'attached', timeout: 30000 });
            mainFrame = await mainFrameElement.contentFrame();
            if (!mainFrame) throw new Error('Could not access main iframe content.');

            // Ensure no modals are blocking
            const blockingModal = mainFrame.locator('.rwWindowContent, .IPModal').first();
            if (await blockingModal.isVisible()) {
                console.log('Blocking modal detected at start of loop. Attempting to close...');
                const closeBtns = mainFrame.locator('.rwCloseButton');
                if (await closeBtns.count() > 0) {
                    await closeBtns.last().click({ force: true });
                    await page.waitForTimeout(2000);
                } else {
                    await page.keyboard.press('Escape');
                    await page.waitForTimeout(2000);
                }
            }

            const criteria = searchCriteria[Math.floor(Math.random() * searchCriteria.length)];
            console.log(`Using criteria: Profile='${criteria.profile}', HSCode='${criteria.hsCode}'`);

            const profileInput = mainFrame.locator('#ctl00_MainContent_cmbxTariffSchedule_Input');
            await expect(profileInput).toBeVisible({ timeout: 10000 });
            await profileInput.click({ force: true }); // Use force to ensure click
            await profileInput.fill(criteria.profile);
            await page.waitForTimeout(1000); 
            await page.keyboard.press('Enter');
            await page.waitForTimeout(5000); 

            const hsInput = mainFrame.locator('#ctl00_MainContent_txtbxHSNumberFilter');
            await expect(hsInput).toBeVisible();
            await hsInput.click();
            await hsInput.clear();
            console.log(`Entering HS Code: ${criteria.hsCode} slowly...`);
            await hsInput.pressSequentially(criteria.hsCode, { delay: 300 });

            console.log('Waiting for autocomplete...');
            await page.waitForTimeout(3000);

            await page.keyboard.press('ArrowDown');
            await page.waitForTimeout(500);
            await page.keyboard.press('Enter');
            console.log('Selected autocomplete item.');
            
            await page.waitForTimeout(2000); 
            let selectedHSValue = await hsInput.inputValue();
            
            if (!selectedHSValue || selectedHSValue.trim() === '') {
                console.warn('HS Code input is empty. Refilling...');
                await hsInput.fill(criteria.hsCode);
            }

            console.log('Clicking Search button...');
            const searchButton = mainFrame.locator('#ctl00_MainContent_lnxbtnNewSearch');
            
            try {
                await expect(page.locator('.IPModal, .bento-busyloader-inner').first()).toBeVisible({ timeout: 3000 });
            } catch (e) { } 
            await expect(page.locator('.IPModal, .bento-busyloader-inner').first()).not.toBeVisible({ timeout: 30000 });
            
            await searchButton.click({ force: true });

            console.log('Waiting for search results to load...');
            const resultsGrid = mainFrame.locator('#ctl00_MainContent_RadGrid1');
            try {
                await expect(resultsGrid).toBeVisible({ timeout: 30000 });
                console.log('SUCCESS: Search results grid displayed.');
            } catch (e) {
                console.warn('Search results grid not found. Retrying...');
                continue;
            }

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

            const targetLink = resultsGrid.locator('a.IPGridLink').first();
            
            if (await targetLink.isVisible()) {
                 const text = await targetLink.innerText();
                 console.log(`Target link text: ${text}`);
            }

            await targetLink.evaluate((el: any) => {
                el.scrollIntoView({ behavior: 'auto', block: 'center', inline: 'center' });
            });
            await page.waitForTimeout(1000); 

            console.log('Clicking result link...');
            // JS click to avoid interception
            await targetLink.evaluate((el: any) => el.click());
            await page.waitForTimeout(3000); 

            const detailFrameElement = mainFrame.locator('iframe[src*="fugGlobalTariffsDetail.aspx"]');
            try {
                await detailFrameElement.waitFor({ state: 'attached', timeout: 30000 });
            } catch (e) {
                console.warn('Detail frame not found or timed out. Retrying...');
                // Recovery attempt
                console.log('Attempting to close detail modal (if any) and retry...');
                // ... (Close logic below)
                continue;
            }
            
            const detailFrame = await detailFrameElement.contentFrame();
            if (!detailFrame) {
                console.warn('Could not access detail frame content. Retrying...');
                continue;
            }

            const mainTabStrip = detailFrame.locator('ul.rtsUL').filter({ hasText: 'Header' }).first();
            await mainTabStrip.waitFor({ state: 'visible', timeout: 20000 });
            
            const tabs = detailFrame.locator('.rtsTxt');
            const tabTexts = await tabs.allInnerTexts();
            console.log('Tabs found:', tabTexts);

            // CHANGED TARGET HERE for TC_037
            // Only proceed if CONTROLS tab is present (Standard structure)
            const targetTabIndex = tabTexts.findIndex(t => t.toUpperCase() === 'CONTROLS');
            
            let countryControlsFound = false;

            if (targetTabIndex !== -1) {
                console.log('SUCCESS: "CONTROLS" tab found. Clicking it...');
                const targetTab = tabs.nth(targetTabIndex);
                await targetTab.click();
                console.log('Waiting 5s for sub-tabs to load...');
                await page.waitForTimeout(5000); 
                
                // Dynamically find all sub-tabs under CONTROLS
                const allTabStrips = detailFrame.locator('.RadTabStrip');
                let countStrips = await allTabStrips.count();
                console.log(`Found ${countStrips} TabStrips initially.`);
                
                // Retry looking for strips if only 1 found (might be slow loading)
                if (countStrips <= 1) {
                    console.log('Only 1 TabStrip found. Waiting another 3s and retrying...');
                    await page.waitForTimeout(3000);
                    countStrips = await allTabStrips.count();
                }

                let subTabStrip;
                if (countStrips > 1) {
                     subTabStrip = allTabStrips.nth(1);
                     if (!await subTabStrip.isVisible()) {
                         // fallback logic
                         for(let i=1; i<countStrips; i++) {
                             if(await allTabStrips.nth(i).isVisible()) {
                                 subTabStrip = allTabStrips.nth(i);
                                 break;
                             }
                         }
                     }
                } else {
                     subTabStrip = allTabStrips.last();
                }

                if (await subTabStrip.isVisible()) {
                    // Find "Country Level Controls" sub-tab
                    const subTab = subTabStrip.locator('.rtsLink').filter({ hasText: /Country Level Controls/i }).first();
                    if (await subTab.isVisible()) {
                        console.log('Found "Country Level Controls" sub-tab. Clicking...');
                        await subTab.click();
                        countryControlsFound = true;
                    } else {
                        console.warn('"Country Level Controls" sub-tab NOT found in sub-strip.');
                    }
                } else {
                    console.warn('Secondary TabStrip for sub-tabs not found or not visible.');
                }
            } else {
                 console.warn('CONTROLS tab not found. Skipping this record (WCO or unsupported profile).');
            }

            if (countryControlsFound) {
                console.log('Waiting for grid to load...');
                // Wait for loading indicator to disappear
                await expect(page.locator('.IPModal, .bento-busyloader-inner').first()).not.toBeVisible({ timeout: 10000 });
                await page.waitForTimeout(3000);

                // Pagination Logic
                console.log('Starting pagination check...');
                
                // Try finding grid with partial ID
                const grid = detailFrame.locator('[id*="rgdCountryLevelControls"]');
                let isGridVisible = false;
                
                try {
                    await grid.first().waitFor({ state: 'visible', timeout: 5000 });
                    isGridVisible = true;
                } catch (e) {
                    console.warn('Grid with ID *rgdCountryLevelControls* not visible.');
                    // Check for "No records" message
                    const bodyText = await detailFrame.locator('body').innerText();
                    if (bodyText.includes('No records to display') || bodyText.includes('No data available')) {
                        console.log('INFO: "No records to display" message found. Treating as success (empty data).');
                        validTargetFound = true; 
                        
                        const screenshotName = `global_tariffs_controls_country_empty_${user.id}.png`;
                        const screenshot = await page.screenshot({ fullPage: true });
                        await test.info().attach(screenshotName, { body: screenshot, contentType: 'image/png' });
                    }
                }

                if (isGridVisible) {
                    let hasNextPage = true;
                    let pageCount = 1;
                    const maxPages = 10; 

                    // Try to get total pages from info part
                    const infoPart = detailFrame.locator('.rgInfoPart').first();
                    let totalPages = 1;
                    try {
                        if (await infoPart.isVisible()) {
                            const infoText = await infoPart.innerText();
                            console.log(`Pagination Info Found: "${infoText}"`);
                            // Example text: "569 items in 23 pages" or "items in 2 pages"
                            const match = infoText.match(/in\s+(\d+)\s+pages/);
                            if (match) {
                                totalPages = parseInt(match[1], 10);
                                console.log(`Detected total pages: ${totalPages}`);
                            } else {
                                console.log('Could not parse total pages from info text.');
                            }
                        } else {
                            console.log('Pagination info part (.rgInfoPart) not visible.');
                        }
                    } catch(e) {
                        console.log('Error reading pagination info:', e);
                    }

                    while (hasNextPage && pageCount <= maxPages) {
                        console.log(`Processing page ${pageCount} of ${totalPages > 1 ? totalPages : '?' }...`);
                        
                        // Re-locate grid to avoid stale element
                        const currentGrid = detailFrame.locator('[id*="rgdCountryLevelControls"]').first();
                        await currentGrid.scrollIntoViewIfNeeded();

                        const gridRows = detailFrame.locator('tr.rgRow, tr.rgAltRow');
                        const rowCount = await gridRows.count();
                        if (rowCount > 0) {
                            console.log(`Page ${pageCount}: Found ${rowCount} rows.`);
                        } else {
                            console.warn(`Page ${pageCount}: No rows found.`);
                        }

                        const screenshotName = `global_tariffs_controls_country_${pageCount}_${user.id}.png`;
                        const screenshot = await page.screenshot({ fullPage: true });
                        await test.info().attach(screenshotName, {
                            body: screenshot,
                            contentType: 'image/png'
                        });
                        console.log(`Attached screenshot: ${screenshotName}`);

                        // Check if we reached the detected total pages
                        if (totalPages > 1 && pageCount >= totalPages) {
                            console.log('Reached detected total pages.');
                            hasNextPage = false;
                            break;
                        }

                        const nextPageBtn = detailFrame.locator('input.rgPageNext[name*="rgdCountryLevelControls"]').first();
                        
                        // Get current page number before click
                        const currentPageEl = detailFrame.locator('.rgCurrentPage span').first();
                        const currentPageBefore = await currentPageEl.isVisible() ? await currentPageEl.innerText() : String(pageCount);

                        if (await nextPageBtn.isVisible()) {
                            // Check if disabled class exists
                            const btnClass = await nextPageBtn.getAttribute('class');
                            if (btnClass && btnClass.includes('rgPageNextDisabled')) {
                                console.log('Next page button is disabled. End of pagination.');
                                hasNextPage = false;
                                break;
                            }

                            console.log('Attempting to go to next page...');
                            try {
                                await nextPageBtn.click({ timeout: 5000 });
                            } catch (e) {
                                console.log('Next page button click failed/timeout. End of pages.');
                                hasNextPage = false;
                                break;
                            }

                            // Wait for loading
                            try {
                                await expect(page.locator('.IPModal, .bento-busyloader-inner').first()).toBeVisible({ timeout: 2000 });
                            } catch(e) {} 
                            await expect(page.locator('.IPModal, .bento-busyloader-inner').first()).not.toBeVisible({ timeout: 10000 });
                            await page.waitForTimeout(2000); 

                            // Verify page changed by checking current page number
                            const currentPageAfter = await currentPageEl.isVisible() ? await currentPageEl.innerText() : '';
                            
                            if (currentPageAfter && currentPageBefore !== currentPageAfter) {
                                console.log(`Page changed from ${currentPageBefore} to ${currentPageAfter}.`);
                                pageCount++;
                            } else if (!currentPageAfter) {
                                console.warn('Could not retrieve new page number. Assuming end of pagination or error.');
                                hasNextPage = false;
                            } else {
                                // Page number same
                                if (totalPages > 1 && pageCount < totalPages) {
                                     console.warn(`Page number did not change (${pageCount}). Stopping to avoid infinite loop.`);
                                     hasNextPage = false; 
                                } else {
                                     console.log('Page number unchanged. End of pagination.');
                                     hasNextPage = false;
                                }
                            }
                        } else {
                            console.log('Next page button not found. End of pagination.');
                            hasNextPage = false;
                        }
                    }

                    validTargetFound = true;
                    console.log('SUCCESS: Country level Controls test complete.');
                }
            } else {
                console.warn('"Country level Controls" tab NOT found. Closing modal and retrying...');
            }

            // Cleanup: Close Modal Logic (Robust)
            console.log('Attempting to close detail modal...');
            let modalClosed = false;
            
            // Strategy 1: Click Close Button in Main Frame (RadWindow)
            const closeBtns = mainFrame.locator('.rwCloseButton');
            if (await closeBtns.count() > 0) {
                const count = await closeBtns.count();
                for (let i = count - 1; i >= 0; i--) {
                    const btn = closeBtns.nth(i);
                    if (await btn.isVisible()) {
                        await btn.click();
                        await page.waitForTimeout(1000);
                    }
                }
            }

            // Strategy 2: Press Escape
            if (await detailFrameElement.isVisible()) {
                await page.keyboard.press('Escape');
                await page.waitForTimeout(1000);
            }

            // Verify closure
            try {
                await expect(detailFrameElement).not.toBeVisible({ timeout: 5000 });
                console.log('SUCCESS: Detail modal closed.');
                modalClosed = true;
            } catch (e) {
                console.warn('WARNING: Detail modal still visible after close attempts.');
            }

            // Strategy 4: Force Reload if still open (Last Resort)
            if (!modalClosed) {
                console.warn('Strategy 4: Reloading page to reset state...');
                await page.reload();
                await mainFrameElement.waitFor({ state: 'attached', timeout: 30000 });
                mainFrame = await mainFrameElement.contentFrame(); // Refresh frame reference
                
                // Navigate back to Global Tariffs Search if needed
                const isSearchVisible = await mainFrame.locator('#ctl00_MainContent_lnxbtnNewSearch').isVisible();
                if (!isSearchVisible) {
                     await page.getByRole('button', { name: 'Content' }).click();
                     await page.locator('a[role="menuitem"]').filter({ hasText: /^Global Tariffs$/ }).click();
                     await mainFrameElement.waitFor({ state: 'attached', timeout: 30000 });
                     mainFrame = await mainFrameElement.contentFrame();
                }
            }
            
            // Close any remaining search tabs
            const closeTabBtn = mainFrame.locator('input[src*="x-red-16x16.png"]:visible'); // Only target visible buttons
            let closeTabAttempts = 0;
            while (await closeTabBtn.count() > 0 && closeTabAttempts < 5) {
                 try {
                     await closeTabBtn.first().click({ timeout: 2000, force: true });
                     await page.waitForTimeout(1000);
                 } catch (e) {
                     console.warn('Failed to close search tab:', e);
                     break; // Stop if we can't click
                 }
                 closeTabAttempts++;
            }

            if (validTargetFound) break;
        }

        if (!validTargetFound) throw new Error('Could not find a record with Country Level Controls after max retries.');
      });

      // Step 3: Logout
      await test.step('Logout', async () => {
        const userButtonName = `MC Test${user.id.substring(6)}`;
        await page.bringToFront();

        const userButton = page.getByRole('button', { name: new RegExp(`^${userButtonName}`)
        });
        await userButton.hover();
        await page.waitForTimeout(500); 
        await userButton.click();

        const logoutButton = page.locator('#logoutButton');
        await expect(logoutButton).toBeVisible({ timeout: 10000 });
        await logoutButton.click({ force: true });

        await expect(page.getByRole('textbox', { name: 'Company' })).toBeVisible({ timeout: 30000 });
      });
    });
  });
});

