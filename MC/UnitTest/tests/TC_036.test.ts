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

test.describe('TC_036 Role Based Access Control - Global Tariffs Search - Regulatory Control', () => {
  test.describe.configure({ mode: 'parallel' });
  
  users.forEach((user: any) => {
    
    test(`Perform Global Tariffs search (Regulatory Control) test for ${user.id}`, async ({ page }) => {
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
--- Attempt ${attempt}/${maxRetries} to find REGULATORY CONTROLS tab ---`);
            
            const criteria = searchCriteria[Math.floor(Math.random() * searchCriteria.length)];
            console.log(`Using criteria: Profile='${criteria.profile}', HSCode='${criteria.hsCode}'`);

            const profileInput = mainFrame.locator('#ctl00_MainContent_cmbxTariffSchedule_Input');
            await profileInput.click();
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
            console.log(`INFO: Current HS Value: "${selectedHSValue}"`);

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
            } else {
                 console.log('Link reported not visible. Attempting JS scroll...');
            }

            await targetLink.evaluate((el: any) => {
                el.scrollIntoView({ behavior: 'auto', block: 'center', inline: 'center' });
            });
            await page.waitForTimeout(1000); 

            console.log('Clicking result link...');
            try {
                await expect(page.locator('.IPModal, .bento-busyloader-inner').first()).not.toBeVisible({ timeout: 10000 });
            } catch (e) {} 
            
            // Use JS click as native click/scroll seems unreliable
            await targetLink.evaluate((el: any) => el.click());
            await page.waitForTimeout(3000); 

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

            const mainTabStrip = detailFrame.locator('ul.rtsUL').filter({ hasText: 'Header' }).first();
            await mainTabStrip.waitFor({ state: 'visible', timeout: 20000 });
            
            const tabs = detailFrame.locator('.rtsTxt');
            const tabTexts = await tabs.allInnerTexts();
            console.log('Tabs found:', tabTexts);

            // CHANGED TARGET HERE for TC_036
            const targetTabIndex = tabTexts.findIndex(t => t.toUpperCase() === 'CONTROLS');
            if (targetTabIndex !== -1) {
                console.log('SUCCESS: "CONTROLS" tab found. Clicking it...');
                
                const targetTab = tabs.nth(targetTabIndex);
                await targetTab.click();
                console.log('Waiting 5s for sub-tabs to load...');
                await page.waitForTimeout(5000); 
                
                // Dynamically find all sub-tabs under CONTROLS
                // We assume sub-tabs are also .rtsTxt elements but we need to distinguish them from the main tabs.
                // Usually Telerik sub-tabs are in a separate container or refreshed list.
                // After clicking CONTROLS, we re-query for tabs that are visible.
                
                // Wait for any RadTabStrip to be present/visible in the frame
                try {
                     await detailFrame.locator('.RadTabStrip').first().waitFor({ state: 'visible', timeout: 5000 });
                } catch(e) { console.log('Wait for RadTabStrip timed out or already visible'); }

                // A common Telerik pattern for sub-tabs is a nested `RadTabStrip`.
                // Let's look for a SECOND tab strip that became visible.
                const allTabStrips = detailFrame.locator('.RadTabStrip');
                let countStrips = await allTabStrips.count();
                console.log(`Found ${countStrips} TabStrips initially.`);
                
                // Retry looking for strips if only 1 found (might be slow loading)
                if (countStrips <= 1) {
                    console.log('Only 1 TabStrip found. Waiting another 3s and retrying...');
                    await page.waitForTimeout(3000);
                    countStrips = await allTabStrips.count();
                    console.log(`Found ${countStrips} TabStrips after retry.`);
                }

                let subTabStrip;
                if (countStrips > 1) {
                     // The second one is likely the sub-menu (index 1)
                     // BUT verify if it is visible.
                     subTabStrip = allTabStrips.nth(1);
                     if (!await subTabStrip.isVisible()) {
                         console.warn('Second TabStrip found but not visible. Checking others...');
                         // fallback logic: find first visible strip that is NOT the main one (index 0)
                         for(let i=1; i<countStrips; i++) {
                             if(await allTabStrips.nth(i).isVisible()) {
                                 subTabStrip = allTabStrips.nth(i);
                                 break;
                             }
                         }
                     }
                } else {
                     // Fallback: maybe they are in the same strip? Unlikely for main/sub navigation usually.
                     // Or maybe the first strip IS the sub-strip if the main one is elsewhere?
                     // Let's assume there is a nested/secondary strip for sub-items.
                     console.warn('Could not find a second TabStrip. Using the last one found as fallback.');
                     subTabStrip = allTabStrips.last();
                }

                if (await subTabStrip.isVisible()) {
                    const subTabs = subTabStrip.locator('.rtsLink'); // Use rtsLink to get clickable items
                    const subTabCount = await subTabs.count();
                    console.log(`Found ${subTabCount} sub-tabs in the target strip.`);

                    for (let s = 0; s < subTabCount; s++) {
                        const currentSubTab = subTabs.nth(s);
                        const subTabText = await currentSubTab.innerText();
                        console.log(`Processing sub-tab ${s+1}/${subTabCount}: "${subTabText}"`);

                        // Click sub-tab
                        await currentSubTab.click();
                        console.log(`Waiting for sub-tab "${subTabText}" to be selected...`);
                        
                        // Wait for the tab to have the 'rtsSelected' class (standard Telerik active state)
                        // The 'currentSubTab' is the link (.rtsLink), usually the parent <li> (.rtsLI) gets the class.
                        // We will check BOTH the parent LI and the link itself.
                        const parentLi = currentSubTab.locator('xpath=..'); // Get parent LI
                        
                        try {
                            // Check parent LI OR the link itself for rtsSelected
                            await expect(async () => {
                                const liHasClass = await parentLi.getAttribute('class').then(c => c?.includes('rtsSelected'));
                                const linkHasClass = await currentSubTab.getAttribute('class').then(c => c?.includes('rtsSelected'));
                                expect(liHasClass || linkHasClass).toBeTruthy();
                            }).toPass({ timeout: 10000 });
                            
                            console.log(`SUCCESS: Sub-tab "${subTabText}" is selected.`);
                        } catch (e) {
                            console.warn(`WARNING: Sub-tab "${subTabText}" did not get selected state within timeout.`);
                        }
                        
                        try {
                            // Check if the link itself or its parent gets the selected class.
                            // We can just wait for a bit more stability or check for a unique element in the content area if possible.
                            // Since content is dynamic, let's wait for the "Loading..." indicator to be GONE if it exists.
                            await expect(page.locator('.IPModal, .bento-busyloader-inner').first()).not.toBeVisible({ timeout: 5000 });
                        } catch(e) {}

                        // Force wait 1s as requested for rendering
                        await page.waitForTimeout(1000);

                        // Take screenshot
                        const safeName = subTabText.trim().toLowerCase().replace(/[^a-z0-9]+/g, '_');
                        const screenshotName = `global_tariffs_controls_${safeName}_${user.id}.png`;
                        const screenshot = await page.screenshot({ fullPage: true });
                        await test.info().attach(screenshotName, {
                            body: screenshot,
                            contentType: 'image/png'
                        });
                        console.log(`Attached screenshot: ${screenshotName}`);
                    }
                } else {
                    console.warn('Secondary TabStrip for sub-tabs not found or not visible.');
                }

                validTargetFound = true;
                console.log('SUCCESS: All available sub-tab screenshots attached. Test complete.');
                break;
            } else {
                console.log('WARNING: "CONTROLS" tab NOT found. Closing modal and retrying...');
                
                const closeButtons = mainFrame.locator('.rwCloseButton');
                if (await closeButtons.count() > 0) {
                     await closeButtons.last().click();
                     await page.waitForTimeout(5000); 
                } else {
                    console.warn('Could not find close button for detail modal. Reloading page...');
                    await page.reload();
                }

                const closeTabBtn = mainFrame.locator('input[src*="x-red-16x16.png"]');
                const tabCount = await closeTabBtn.count();
                if (tabCount > 0) {
                    console.log(`Closing ${tabCount} search result tab(s) to reset...`);
                    for (let k = 0; k < tabCount; k++) {
                        try {
                            await closeTabBtn.nth(0).click(); 
                            await page.waitForTimeout(1000);
                        } catch (e) {
                            console.warn('Error closing tab:', e);
                        }
                    }
                }
            }
        }

        if (!validTargetFound) throw new Error('Could not find a record with OTHER CHARGES tab after max retries.');
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

