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

test.describe('TC_033 Role Based Access Control - Global Tariffs (Classic) Deep Tree Expansion', () => {
  test.describe.configure({ mode: 'parallel' });
  
  users.forEach((user: any) => {
    
    test(`Perform Global Tariffs (Classic) deep tree expansion test for ${user.id}`, async ({ page }) => {
      test.setTimeout(300000); // Increased timeout for multiple expansions

      console.log(`--- Test Start: ${user.id} ---`);

      // Step 1: Login
      await test.step('Login', async () => {
        await page.goto(BASE_URL + '/Logon.aspx?ReturnUrl=%2fgtm%2fhome');
        await page.getByRole('textbox', { name: 'Company' }).fill(user.company);
        await page.getByRole('textbox', { name: 'Username' }).fill(user.id);
        await page.getByRole('textbox', { name: 'Password' }).fill(user.pass);
        await page.getByRole('link', { name: 'Login' }).click();
      });

      // Step 2: Access Global Tariffs (Classic) and Expand Tree Node Deeply
      await test.step('Access Global Tariffs (Classic) and Expand Tree Node Deeply', async () => {
        
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

        // Select a random search criteria (Profile) to load the tree
        const criteria = searchCriteria[Math.floor(Math.random() * searchCriteria.length)];
        console.log(`Using criteria: Profile='${criteria.profile}'`);

        // Step 2a: Select Tariff Profile
        const profileInput = mainFrame.locator('#ctl00_MainContent_cmbxTariffSchedule_Input');
        await profileInput.click();
        await profileInput.fill(criteria.profile);
        await page.waitForTimeout(1000); // Wait for input to process
        await page.keyboard.press('Enter');
        
        // Wait for potential postback/reload and Tree visibility
        await page.waitForTimeout(12000); 
        const resultTree = mainFrame.locator('#ctl00_MainContent_rtvHSHierarchy');
        await expect(resultTree).toBeVisible({ timeout: 30000 });
        console.log('SUCCESS: Tree structure displayed.');

        // Step 2b: Deeply Expand Tree Nodes until "View Details" is found
        let foundDetails = false;
        const maxDepth = 15; // Allow deeper traversal

        for (let i = 0; i < maxDepth; i++) {
             console.log(`--- Depth Iteration ${i+1}/${maxDepth} ---`);

             // Check for "View Details" link
             // We look for a link with exact text 'View Details' or containing it.
             const viewDetailsLink = mainFrame.locator('a').filter({ hasText: 'View Details' }).first();
             
             if (await viewDetailsLink.isVisible()) {
                 console.log('Found "View Details" link. Clicking...');
                 
                 // Take Pre-Click Screenshot
                 await viewDetailsLink.scrollIntoViewIfNeeded();
                 await page.waitForTimeout(500);
                 const preClickScreenshot = await page.screenshot({ fullPage: true });
                 await test.info().attach(`global_tariffs_classic_pre_click_${user.id}.png`, {
                    body: preClickScreenshot,
                    contentType: 'image/png'
                 });

                 await viewDetailsLink.click();
                 foundDetails = true;
                 break;
             }

             // Find expand buttons (.rtPlus)
             const expandButtons = mainFrame.locator('.rtPlus');
             const count = await expandButtons.count();

             if (count > 0) {
                 console.log(`Found ${count} expand buttons. Expanding the first one...`);
                 
                 const targetButton = expandButtons.first();
                 
                 // Scroll to it
                 try {
                     await targetButton.scrollIntoViewIfNeeded();
                 } catch (e) {
                     // Fallback scroll if needed, though scrollIntoViewIfNeeded is robust
                     await mainFrame.locator('body').evaluate((body) => {
                        body.ownerDocument.defaultView?.scrollBy(0, 100);
                     });
                 }
                 
                 await targetButton.click();
                 await page.waitForTimeout(3000); // Wait for node content to load
             } else {
                 console.warn('No more expand buttons found, and "View Details" not seen.');
                 break;
             }
        }

        if (foundDetails) {
            // Wait for Modal to appear
            console.log('Waiting for modal window...');
            await page.waitForTimeout(5000); // Wait for animation/load

            // Verify some modal presence (RadWindow usually creates a wrapper)
            // We simply take a screenshot to verify visual appearance as requested.
            
            // Take Screenshot
            const screenshotBufferResult = await page.screenshot({ fullPage: true });
            await test.info().attach(`global_tariffs_classic_deep_expansion_${user.id}.png`, {
                body: screenshotBufferResult,
                contentType: 'image/png'
            });
            console.log('SUCCESS: Screenshot attached. Test complete.');

        } else {
            // Take debug screenshot
            const debugScreenshot = await page.screenshot({ fullPage: true });
            await test.info().attach(`debug_failed_expansion_${user.id}.png`, { body: debugScreenshot, contentType: 'image/png' });
            throw new Error('Failed to find "View Details" link after max depth expansions.');
        }

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
