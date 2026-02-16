import { test, expect } from '@playwright/test';
import { BASE_URL } from './config';
import * as fs from 'fs';
import * as path from 'path';

// Load user definitions
const users = JSON.parse(fs.readFileSync(path.join(__dirname, 'account.dat'), 'utf8'));

// Define scenarios
const scenarios = [
    {
        regulation: "CN - CHINA - China Export Control List of Dual-Use Items",
        ecnCodes: ['1C902'],
        hsCodes: ['2805', '2846']
    },
    {
        regulation: "CN - CHINA - China Export License Management Catalogue of Dual-use Items and Technologies",
        ecnCodes: ['1.1.2', '1.1.3'],
        hsCodes: ['2801', '2804']
    },
    {
        regulation: "JP - JAPAN - Japan Export Control Order",
        ecnCodes: ['1.10', '1.11', '1.12', '1.13', '1.14'],
        hsCodes: []
    },
    {
        regulation: "JP - JAPAN - Japan Export Control Order: Appended Table 2-3 (Prohibited Items for Export to Russia etc.)",
        ecnCodes: ['1.2'],
        hsCodes: []
    },
    {
        regulation: "JP - JAPAN - Japan Foreign Exchange Order",
        ecnCodes: ['10.1', '10.2', '10.3', '10.4'],
        hsCodes: []
    },
    {
        regulation: "US - UNITED STATES OF AMERICA - EAR/Commerce Control List",
        ecnCodes: ['2', '3', '4'],
        hsCodes: []
    },
    {
        regulation: "US - UNITED STATES OF AMERICA - International Traffic in Arms Regulations (ITAR)",
        ecnCodes: ['4', '5', '6', '7'],
        hsCodes: []
    },
    {
        regulation: "US - UNITED STATES OF AMERICA - Supplement No. 6 to Part 746 - Items That Require a License For Russia and Belarus",
        ecnCodes: ['7', '8', '9'],
        hsCodes: []
    }
];

test.describe('Role Based Access Control - Specific Regulation Search', () => {
  
  for (const user of users) {
    
    test(`Perform search with specific regulation for ${user.id}`, async ({ page }) => {
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

      // Step 2: Perform search loop
      await test.step('Perform Search 5 times or until screenshot is taken', async () => {
        let screenshotTaken = false;
        for (let i = 1; i <= 5; i++) {
            console.log(`\n--- Repetition ${i}/5 for user ${user.id} ---`);
            
            await page.getByRole('button', { name: 'Content' }).click();
            await page.getByRole('menuitem', { name: 'ECN/Dual Use List Query' }).click();

            let mainFrameElement = page.locator('iframe[name="legacy-outlet"]');
            await mainFrameElement.waitFor({ state: 'attached', timeout: 30000 });
            let mainFrame = await mainFrameElement.contentFrame();
            if (!mainFrame) throw new Error('Could not access main iframe content.');

            const scenario = scenarios[Math.floor(Math.random() * scenarios.length)];
            console.log(`INFO: Testing with Regulation: "${scenario.regulation}"`);

            const regulationInput = mainFrame.locator('#ctl00_MainContent_cmbxRegulationList_Input');
            await regulationInput.click();
            await mainFrame.locator('#ctl00_MainContent_cmbxRegulationList_DropDown').getByText(scenario.regulation, { exact: true }).click();
            await expect(regulationInput).toHaveValue(scenario.regulation);
            await page.waitForTimeout(3000); // Wait for UI to settle

            mainFrame = await page.locator('iframe[name="legacy-outlet"]').contentFrame();
            if (!mainFrame) throw new Error('Could not access reloaded iframe content.');

            const canSearchEcn = scenario.ecnCodes.length > 0;
            const canSearchHs = scenario.hsCodes.length > 0;
            let searchType = '';
            if (canSearchEcn && canSearchHs) { searchType = Math.random() < 0.5 ? 'ECN' : 'HS'; } 
            else if (canSearchEcn) { searchType = 'ECN'; } 
            else if (canSearchHs) { searchType = 'HS'; } 
            else { throw new Error(`No codes defined for regulation: ${scenario.regulation}`); } 
            
            if (searchType === 'ECN') {
                const randomEcn = scenario.ecnCodes[Math.floor(Math.random() * scenario.ecnCodes.length)];
                console.log(`INFO: Searching with ECN code: ${randomEcn}`);
                const ecnInput = mainFrame.locator('#ctl00_MainContent_cmbxStatusBarECN_Input');
                await ecnInput.click();
                await ecnInput.type(randomEcn, { delay: 200 });
                await mainFrame.locator('#ctl00_MainContent_cmbxStatusBarECN_DropDown').locator('li.rcbItem').first().click();
            } else {
                const randomHs = scenario.hsCodes[Math.floor(Math.random() * scenario.hsCodes.length)];
                console.log(`INFO: Searching with HS code: ${randomHs}`);
                const hsInput = mainFrame.locator('#ctl00_MainContent_cmbxStatusBarHSNumber_Input');
                await hsInput.click();
                await hsInput.type(randomHs, { delay: 200 });
                await mainFrame.locator('#ctl00_MainContent_cmbxStatusBarHSNumber_DropDown').locator('li.rcbItem').first().click();
            }

            await mainFrame.locator('#ctl00_MainContent_lnxbtnSearch').click();
            
            const resultsGrid = mainFrame.locator('#ctl00_MainContent_rgdMain');
            await expect(resultsGrid).toBeVisible({ timeout: 30000 });
            await expect(resultsGrid.locator('table.rgMasterTable > tbody > tr').first()).toBeVisible({ timeout: 20000 });
            await page.waitForTimeout(3000);

            console.log('INFO: Expanding the first result node...');
            await mainFrame.locator('[id$="_GECBtnExpandColumn"]').first().click();

            const tabStrip = mainFrame.locator('.RadTabStrip');
            await expect(tabStrip).toBeVisible({ timeout: 15000 });

            const tabNames = await mainFrame.locator('div.rtsLevel.rtsLevel1').getByRole('link').evaluateAll(links => 
                links.map(link => (link.textContent || '').trim()).filter(name => name)
            );
            console.log(`INFO: Found tabs: ${tabNames.join(', ')}`);

            const hsRelatedControlsTabName = tabNames.find(name => name.toUpperCase().startsWith('HS RELATED CONTROLS'));

            if (!hsRelatedControlsTabName) {
                console.log('INFO: "HS RELATED CONTROLS" tab not found, skipping...');
            } else {
                console.log(`INFO: Found and clicking "${hsRelatedControlsTabName}" tab...`);
                await mainFrame.locator('div.rtsLevel.rtsLevel1').getByRole('link', { name: hsRelatedControlsTabName }).click();

                const hsGrid = mainFrame.locator('div[id$="rgdHSRelatedControls"]');
                await expect(hsGrid).toBeVisible({ timeout: 15000 });
                await expect(hsGrid.locator('tr.rgRow, tr.rgAltRow').first()).toBeVisible({ timeout: 20000 });
                
                const controlLinks = hsGrid.locator('a[href*="javascript:OpenInPopup"]');
                const linkCount = await controlLinks.count();

                if (linkCount > 0) {
                    console.log(`INFO: Found ${linkCount} control links. Clicking the first one.`);
                    await controlLinks.first().click();

                    // --- Final Modal Interaction Logic ---
                    // The RadWindow and its overlay are created in the context of the iframe.
                    const radWindow = mainFrame.locator('div.RadWindow');
                    await expect(radWindow).toBeVisible({ timeout: 20000 });
                    console.log('SUCCESS: RadWindow modal is visible.');

                    console.log('INFO: Waiting 5 seconds for modal content to load before screenshot...');
                    await page.waitForTimeout(5000);
                    
                    console.log('INFO: Taking screenshot of modal...');
                    const screenshotBuffer = await radWindow.screenshot();
                    await test.info().attach(`${user.id}-${scenario.regulation.substring(0, 10)}-${i}.png`, { body: screenshotBuffer, contentType: 'image/png' });
                    screenshotTaken = true;
                    console.log('SUCCESS: Screenshot taken.');

                    const closeButton = mainFrame.locator('span.rwCloseButton[title="Close"]');
                    await expect(closeButton).toBeVisible();
                    console.log('INFO: Clicking close button...');
                    await closeButton.click({ force: true });
                    
                    const overlay = mainFrame.locator('div.TelerikModalOverlay');
                    await expect(radWindow).toBeHidden({ timeout: 15000 });
                    await expect(overlay).toBeHidden({ timeout: 15000 });
                    console.log("SUCCESS: Modal closed successfully.");

                } else {
                    console.log('INFO: No control links found to click.');
                }
            }

            if (screenshotTaken) {
                console.log('SUCCESS: Screenshot captured, ending test for this user.');
                break;
            }

            console.log('INFO: Clicking "New Search" to reset for next loop.');
            await mainFrame.locator('#ctl00_MainContent_hyxlnkNewSearch').click();
            await expect(mainFrame.locator('#ctl00_MainContent_rgdMain')).toBeHidden();
        }
      });

      // Step 3: Logout
      await test.step('Logout', async () => {
        const userButtonName = `MC Test${user.id.substring(6)}`;
        await page.bringToFront();
        await page.getByRole('button', { name: new RegExp(`^${userButtonName}`) }).click();
        await page.getByRole('button', { name: ' Sign out' }).click({ force: true }); 
        await expect(page.getByRole('textbox', { name: 'Company' })).toBeVisible({ timeout: 30000 });
      });
    });
  }
});

