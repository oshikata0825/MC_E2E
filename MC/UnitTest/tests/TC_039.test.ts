import { test, expect } from '@playwright/test';
import { BASE_URL } from './config';
import * as fs from 'fs';
import * as path from 'path';

// Load user definitions
const users = JSON.parse(fs.readFileSync(path.join(__dirname, 'account.dat'), 'utf8'));

// Load valid profile/HS code combinations from file
const profileComboPath = path.join(__dirname, 'profile_combination');
let searchCriteria: { profile: string, hsCode: string }[] = [];

try {
    const rawCombos = fs.readFileSync(profileComboPath, 'utf8').split('\n').filter(line => line.trim() !== '');
    searchCriteria = rawCombos.map(line => {
        const parts = line.split(':');
        if (parts.length >= 2) {
            return { profile: parts[0].trim(), hsCode: parts[1].trim() };
        }
        return null;
    }).filter(item => item !== null) as { profile: string, hsCode: string }[];
} catch (e) {
    console.error(`Error reading profile_combination file: ${e}`);
    // Fallback or fail if critical
    searchCriteria = [
        { profile: 'JP - JAPAN - Japans Tariff Schedule (IMPORT)', hsCode: '870310000' }
    ];
}

test.describe('TC_039 Role Based Access Control - Global Tariffs Search - DUTY Tab Verification', () => {
  test.describe.configure({ mode: 'parallel' });
  
  users.forEach((user: any) => {
    
    test(`Perform Global Tariffs search (DUTY Tab) test for ${user.id}`, async ({ page }) => {
      test.setTimeout(300000); // 5 minutes

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

        // Select a criteria from the loaded list
        const criteria = searchCriteria[Math.floor(Math.random() * searchCriteria.length)];
        console.log(`Using criteria: Profile='${criteria.profile}', HSCode='${criteria.hsCode}'`);

        const profileInput = mainFrame.locator('#ctl00_MainContent_cmbxTariffSchedule_Input');
        await expect(profileInput).toBeVisible({ timeout: 10000 });
        await profileInput.click({ force: true });
        await profileInput.fill(criteria.profile);
        await page.waitForTimeout(1000); 
        await page.keyboard.press('Enter');
        await page.waitForTimeout(5000); 

                    const hsInput = mainFrame.locator('#ctl00_MainContent_txtbxHSNumberFilter');
                    await expect(hsInput).toBeVisible({ timeout: 30000 }); // Increased timeout to 30s
                    await hsInput.click();
                    await hsInput.clear();
                    console.log(`Entering HS Code: ${criteria.hsCode} slowly...`);        await hsInput.pressSequentially(criteria.hsCode, { delay: 300 });

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
        await expect(resultsGrid).toBeVisible({ timeout: 30000 });
        console.log('SUCCESS: Search results grid displayed.');

        const rows = resultsGrid.locator('tr.rgRow, tr.rgAltRow');
        await expect(async () => {
            const count = await rows.count();
            expect(count).toBeGreaterThan(0);
        }).toPass({ timeout: 10000 });

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
        await detailFrameElement.waitFor({ state: 'attached', timeout: 30000 });
        
        const detailFrame = await detailFrameElement.contentFrame();
        if (!detailFrame) throw new Error('Could not access detail frame content.');

        const mainTabStrip = detailFrame.locator('ul.rtsUL').filter({ hasText: 'Header' }).first();
        await mainTabStrip.waitFor({ state: 'visible', timeout: 20000 });
        
        const tabs = detailFrame.locator('.rtsTxt');
        const tabTexts = await tabs.allInnerTexts();
        console.log('Tabs found:', tabTexts);

        // Verify DUTY tab exists
        const targetTabIndex = tabTexts.findIndex(t => t.toUpperCase() === 'DUTY');
        
                    if (targetTabIndex !== -1) {
                        console.log('SUCCESS: "DUTY" tab found. Clicking it...');
                        const targetTab = tabs.nth(targetTabIndex);
                        await targetTab.click();
                        
                        // Wait for loading
                        await expect(page.locator('.IPModal, .bento-busyloader-inner').first()).not.toBeVisible({ timeout: 10000 });
                        await page.waitForTimeout(3000); 
                        
                        // Screenshot of DUTY tab content (Initial)
                        console.log('Taking screenshot of DUTY tab (Initial)...');
                        const screenshotName = `global_tariffs_duty_${criteria.hsCode}_${user.id}.png`;
                        const screenshot = await page.screenshot({ fullPage: true });
                        await test.info().attach(screenshotName, {
                            body: screenshot,
                            contentType: 'image/png'
                        });
                        console.log(`Attached screenshot: ${screenshotName}`);
        
                                        // --- Country Filter Logic ---
                                        console.log('Starting Country Filter tests...');
                                        
                                        // Define a pool of countries to choose from (excluding JP which is mandatory)
                                        const countryPool = ['US', 'CN', 'DE', 'FR', 'GB', 'CA', 'AU', 'IN', 'BR', 'KR'];
                                        
                                        // Select 2 random distinct countries from the pool
                                        const randomCountries: string[] = [];
                                        while (randomCountries.length < 2) {
                                            const randomPick = countryPool[Math.floor(Math.random() * countryPool.length)];
                                            if (!randomCountries.includes(randomPick)) {
                                                randomCountries.push(randomPick);
                                            }
                                        }
                        
                                        // Construct final test list: JP + 2 Random
                                        const countriesToTest = ['JP', ...randomCountries];
                                        console.log(`Selected countries for filter test: ${countriesToTest.join(', ')}`);
                        
                                        // The filter input is inside the DETAIL frame (iframe src*="fugGlobalTariffsDetail.aspx")                        // Double check if the filter is in the detail frame or main frame.
                        // Usually detail view filters are inside the detail frame.
                        // The user provided element: #ctl00_MainContent_cmbxCountryFilter_Input
                        // Based on previous context, we are inside `detailFrame`.
                        
                        const countryFilterInput = detailFrame.locator('#ctl00_MainContent_cmbxCountryFilter_Input');
        
                        for (const countryCode of countriesToTest) {
                            console.log(`Filtering by Country: ${countryCode}`);
        
                            // Ensure input is interactable
                            await expect(countryFilterInput).toBeVisible({ timeout: 5000 });
                            await countryFilterInput.click({ force: true });
                            await countryFilterInput.clear();
                            await page.waitForTimeout(500);
                            await countryFilterInput.fill(countryCode);
                            await page.waitForTimeout(1000);
                            await page.keyboard.press('Enter');
        
                                                                    // Wait for reload/update
        
                                                                    console.log('Waiting for filter update...');
        
                                                                    
        
                                                                    // 1. Wait for loader to appear (short timeout) - Both in Page and Detail Frame
        
                                                                    try {
        
                                                                        await Promise.any([
        
                                                                            page.locator('.IPModal, .bento-busyloader-inner, .radLoading').first().waitFor({ state: 'visible', timeout: 3000 }),
        
                                                                            detailFrame.locator('.IPModal, .bento-busyloader-inner, .radLoading').first().waitFor({ state: 'visible', timeout: 3000 })
        
                                                                        ]);
        
                                                                    } catch (e) { }
        
                                                
        
                                                                    // 2. Wait for ALL loaders to DISAPPEAR
        
                                                                    await expect(page.locator('.IPModal, .bento-busyloader-inner, .radLoading').first()).not.toBeVisible({ timeout: 40000 });
        
                                                                    await expect(detailFrame.locator('.IPModal, .bento-busyloader-inner, .radLoading').first()).not.toBeVisible({ timeout: 40000 });
        
                                                
        
                                                                    // 3. Wait for the Country Filter Input to be ENABLED/INTERACTABLE again
        
                                                                    // Telerik inputs often get disabled or read-only during postbacks
        
                                                                    await expect(countryFilterInput).toBeEnabled({ timeout: 20000 });
        
                                                
        
                                                                    // 4. Wait for Content Stabilization
        
                                                                    const dutyContent = detailFrame.locator('.RadGrid, .rgMasterTable').first(); 
        
                                                                    await dutyContent.waitFor({ state: 'visible', timeout: 20000 }).catch(() => console.log('Grid not found, waiting for generic stabilization...'));
        
                                                                    
        
                                                                    console.log('Loaders gone. Waiting 10s for visual stability...');
        
                                                                    await page.waitForTimeout(10000); // 10s fixed wait
        
                                                
        
                                                                    // Take screenshot
        
                                                                    const filterScreenshotName = `global_tariffs_duty_filter_${countryCode}_${criteria.hsCode}_${user.id}.png`;                            const filterScreenshot = await page.screenshot({ fullPage: true });
                            await test.info().attach(filterScreenshotName, {
                                body: filterScreenshot,
                                contentType: 'image/png'
                            });
                            console.log(`Attached screenshot: ${filterScreenshotName}`);
                        }
                        
                                        console.log('SUCCESS: DUTY tab verification and Country Filtering complete.');
                        
                                        
                        
                                        // Append to custom report
                        
                                        const reportLine = `${new Date().toISOString()} - User: ${user.id} - Profile: ${criteria.profile} - HS Code: ${criteria.hsCode} - Result: SUCCESS (DUTY Tab Found & Filtered)\n`;
                        
                                        fs.appendFileSync(path.join(__dirname, '../execution_report.txt'), reportLine);
                        
                        
                        
                                    } else {
                        
                                        // Fail the test if DUTY tab is not found (since we expect it for these inputs)
                        
                                        const errorMsg = `"DUTY" tab NOT found for Profile: ${criteria.profile}, HS Code: ${criteria.hsCode}. Tabs visible: ${tabTexts.join(', ')}`;
                        
                                        const reportLine = `${new Date().toISOString()} - User: ${user.id} - Profile: ${criteria.profile} - HS Code: ${criteria.hsCode} - Result: FAILED (${errorMsg})\n`;
                        
                                        fs.appendFileSync(path.join(__dirname, '../execution_report.txt'), reportLine);
                        
                                        throw new Error(errorMsg);
                        
                                    }

        // Cleanup: Close Modal Logic
        console.log('Attempting to close detail modal...');
        
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
        } else {
             // Strategy 2: Press Escape
             await page.keyboard.press('Escape');
             await page.waitForTimeout(1000);
        }
        
        // Verify closure
        try {
            await expect(detailFrameElement).not.toBeVisible({ timeout: 5000 });
        } catch (e) {
            console.warn('WARNING: Detail modal still visible after close attempts.');
        }

        // Close any remaining search tabs
        const closeTabBtn = mainFrame.locator('input[src*="x-red-16x16.png"]:visible');
        if (await closeTabBtn.count() > 0) {
             try {
                 await closeTabBtn.first().click({ timeout: 2000, force: true });
             } catch (e) {}
        }

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
