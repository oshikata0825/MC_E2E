import { test, expect } from '@playwright/test';
import { BASE_URL } from './config';
import * as fs from 'fs';
import * as path from 'path';

// Load user definitions
const users = JSON.parse(fs.readFileSync(path.join(__dirname, 'account.dat'), 'utf8'));

test.describe('TC_038 Role Based Access Control - Rulings Page Access', () => {
  test.describe.configure({ mode: 'parallel' });
  
  users.forEach((user: any) => {
    
    test(`Perform Rulings page access test for ${user.id}`, async ({ page }) => {
      test.setTimeout(120000);

      console.log(`--- Test Start: ${user.id} ---`);

      // Step 1: Login
      await test.step('Login', async () => {
        await page.goto(BASE_URL + '/Logon.aspx?ReturnUrl=%2fgtm%2fhome');
        await page.getByRole('textbox', { name: 'Company' }).fill(user.company);
        await page.getByRole('textbox', { name: 'Username' }).fill(user.id);
        await page.getByRole('textbox', { name: 'Password' }).fill(user.pass);
        await page.getByRole('link', { name: 'Login' }).click();
      });

      // Step 2: Access Rulings Page
      await test.step('Access Rulings Page', async () => {
        
        await page.getByRole('button', { name: 'Content' }).click();
        
        // Use regex for exact match 'Rulings'
        await page.locator('a[role="menuitem"]').filter({ hasText: /^Rulings$/ }).click();

        let mainFrameElement = page.locator('iframe[name="legacy-outlet"]');
        await mainFrameElement.waitFor({ state: 'attached', timeout: 30000 });
        let mainFrame = await mainFrameElement.contentFrame();
        if (!mainFrame) throw new Error('Could not access main iframe content.');

        // Verify page name
        // Assuming #lbxPageName is common across these legacy pages
        const pageNameLabel = mainFrame.locator('#lbxPageName');
        await expect(pageNameLabel).toContainText('Rulings', { timeout: 20000 });
        console.log('SUCCESS: Verified page name contains "Rulings".');

        // Step 2a: Ensure 'Classification' is selected in Rulings Type
        const rulingsTypeInput = mainFrame.locator('#ctl00_MainContent_cmbxBindingRulingsType_Input');
        await expect(rulingsTypeInput).toBeVisible({ timeout: 10000 });
        
        const currentValue = await rulingsTypeInput.inputValue();
        console.log(`Current Rulings Type: '${currentValue}'`);

        if (currentValue !== 'Classification') {
            console.log("Selecting 'Classification'...");
            await rulingsTypeInput.click();
            await rulingsTypeInput.fill('Classification');
            await page.waitForTimeout(500);
            await page.keyboard.press('Enter');
            await page.waitForTimeout(2000); // Wait for potential reload/update
        } else {
             console.log("'Classification' is already selected.");
        }

        // Step 2b: Enter HS Code and Search
        const hsCodes = ['0101', '8501', '8471', '8542', '8802', '3926', '8413'];
        const hsInput = mainFrame.locator('#ctl00_MainContent_cmbxStatusBarHSNumber_Input');
        await hsInput.click();
        const randomHS = hsCodes[Math.floor(Math.random() * hsCodes.length)];
        console.log(`Entering HS Code: ${randomHS}`);
        await hsInput.fill(randomHS);
        
        // Wait for autocomplete suggestions
        console.log('Waiting for autocomplete...');
        await page.waitForTimeout(3000);

        // Select using Keyboard: ArrowDown 1 to 3 times (randomly)
        const arrowDownCount = Math.floor(Math.random() * 3) + 1; 
        console.log(`Pressing ArrowDown ${arrowDownCount} times to select autocomplete item.`);
        
        for (let k = 0; k < arrowDownCount; k++) {
            await page.keyboard.press('ArrowDown');
            await page.waitForTimeout(500);
        }
        await page.keyboard.press('Enter');

        // Step 2c: Verify Search Results
        console.log('Waiting for search results...');
        const resultsGrid = mainFrame.locator('#ctl00_MainContent_RadGrid2_ctl00');
        await expect(resultsGrid).toBeVisible({ timeout: 30000 });
        
        // Wait a bit for rows to populate
        await page.waitForTimeout(2000);

        // Verify Reference Code links in rows
        const rows = resultsGrid.locator('tr.rgRow, tr.rgAltRow');
        const rowCount = await rows.count();
        console.log(`Found ${rowCount} rows in search results.`);

        if (rowCount > 0) {
            // Check up to first 5 rows to save time
            const rowsToCheck = Math.min(rowCount, 5);
            for (let i = 0; i < rowsToCheck; i++) {
                const row = rows.nth(i);
                // First cell should contain an anchor
                const refLink = row.locator('td').first().locator('a');
                await expect(refLink).toBeVisible();
                const refText = await refLink.innerText();
                console.log(`Row ${i+1}: Found Reference Code link '${refText}'`);
            }
        } else {
            console.warn('No search results found for this HS Code.');
        }

        // Take Screenshot
        const screenshotBufferResult = await page.screenshot({ fullPage: true });
        await test.info().attach(`rulings_page_access_${user.id}.png`, {
            body: screenshotBufferResult,
            contentType: 'image/png'
        });

        console.log('SUCCESS: Screenshot attached. Test complete.');
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

