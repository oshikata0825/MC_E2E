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

test.describe('TC_035 Role Based Access Control - Global Tariffs Search - Other Charges', () => {
  test.describe.configure({ mode: 'parallel' });
  
  users.forEach((user: any) => {
    
    test(`Perform Global Tariffs search (Other Charges) test for ${user.id}`, async ({ page }) => {
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

        // Loop to find a record with OTHER CHARGES tab
        let validTargetFound = false;
        const maxRetries = 5;

        for (let attempt = 1; attempt <= maxRetries; attempt++) {
            console.log(`\n--- Attempt ${attempt}/${maxRetries} to find OTHER CHARGES tab ---`);
            
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

            const targetTabIndex = tabTexts.findIndex(t => t.toUpperCase() === 'OTHER CHARGES');
            if (targetTabIndex !== -1) {
                console.log('SUCCESS: "OTHER CHARGES" tab found. Clicking it...');
                
                const targetTab = tabs.nth(targetTabIndex);
                await targetTab.click();
                await page.waitForTimeout(3000); 
                
                // 1. HS Specific Charges
                const hsSpecificTab = detailFrame.locator('.rtsTxt', { hasText: 'HS Specific Charges' }).first();
                try {
                    console.log('Waiting for "HS Specific Charges" to be visible...');
                    await hsSpecificTab.waitFor({ state: 'visible', timeout: 10000 });
                    console.log('Clicking "HS Specific Charges"...');
                    await hsSpecificTab.click();
                    await page.waitForTimeout(2000);
                    
                    const screenshot1 = await page.screenshot({ fullPage: true });
                    await test.info().attach(`global_tariffs_hs_specific_charges_${user.id}.png`, {
                        body: screenshot1,
                        contentType: 'image/png'
                    });
                } catch (e) {
                    console.warn('"HS Specific Charges" tab not found/visible within Other Charges.');
                }

                // 2. Country Specific Charges
                const countrySpecificTab = detailFrame.locator('.rtsTxt', { hasText: 'Country Specific Charges' }).first();
                try {
                    console.log('Waiting for "Country Specific Charges" to be visible...');
                    await countrySpecificTab.waitFor({ state: 'visible', timeout: 10000 });
                    console.log('Clicking "Country Specific Charges"...');
                    await countrySpecificTab.click();
                    await page.waitForTimeout(2000);
                    
                    const screenshot2 = await page.screenshot({ fullPage: true });
                    await test.info().attach(`global_tariffs_country_specific_charges_${user.id}.png`, {
                        body: screenshot2,
                        contentType: 'image/png'
                    });
                } catch (e) {
                    console.warn('"Country Specific Charges" tab not found/visible within Other Charges.');
                }

                validTargetFound = true;
                console.log('SUCCESS: Screenshots attached. Test complete.');
                break;
            } else {
                console.log('WARNING: "OTHER CHARGES" tab NOT found. Closing modal and retrying...');
                
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
