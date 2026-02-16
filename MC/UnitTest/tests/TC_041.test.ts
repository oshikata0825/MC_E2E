import { test, expect } from '@playwright/test';
import { BASE_URL } from './config';
import * as fs from 'fs';
import * as path from 'path';

// Load user definitions
const users = JSON.parse(fs.readFileSync(path.join(__dirname, 'account.dat'), 'utf8'));

test.describe('TC_041 Role Based Access Control - Search History Page Access', () => {
  test.describe.configure({ mode: 'parallel' });
  
  users.forEach((user: any) => {
    
    test(`Perform Search History page access test for ${user.id}`, async ({ page }) => {
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

      // Step 2: Access Search History Page
      await test.step('Access Search History Page', async () => {
        
        await page.getByRole('button', { name: 'Content' }).click();
        
        // Use regex for exact match 'Search History Detail'
        await page.locator('a[role="menuitem"]').filter({ hasText: /^Search History Detail$/ }).click();

        let mainFrameElement = page.locator('iframe[name="legacy-outlet"]');
        await mainFrameElement.waitFor({ state: 'attached', timeout: 30000 });
        let mainFrame = await mainFrameElement.contentFrame();
        if (!mainFrame) throw new Error('Could not access main iframe content.');

        // Verify page name
        const pageNameLabel = mainFrame.locator('#lbxPageName');
        await expect(pageNameLabel).toContainText('Search History Detail', { timeout: 20000 });
        console.log('SUCCESS: Verified page name contains "Search History Detail".');

        // Step 2.1: Select 'Recent Searches Type'
        await test.step('Select Recent Searches Type', async () => {
            // Target the arrow button to open the dropdown. Derived from Telerik ID naming convention.
            const comboBoxArrow = mainFrame.locator('#ctl00_MainContent_cmbxRecentSearchesType_Arrow');
            const dropDown = mainFrame.locator('#ctl00_MainContent_cmbxRecentSearchesType_DropDown');

            await expect(comboBoxArrow).toBeVisible({ timeout: 10000 });
            await comboBoxArrow.click();

            // Wait for dropdown to be visible
            await expect(dropDown).toBeVisible({ timeout: 5000 });

            // Get all items in the dropdown
            const items = dropDown.locator('li.rcbItem');
            const count = await items.count();
            
            if (count === 0) {
                throw new Error('No items found in Recent Searches Type dropdown.');
            }

            // Select random item
            const randomIndex = Math.floor(Math.random() * count);
            const itemToSelect = items.nth(randomIndex);
            const itemText = await itemToSelect.innerText();
            
            console.log(`Selecting Recent Search Type: ${itemText}`);
            await itemToSelect.click();

            // Wait for the grid to update/loading indicator to vanish. 
            // Telerik grids usually show a loading overlay with class 'rgLoader' or similar, 
            // or we can wait for the grid to be visible and stable.
            await page.waitForTimeout(3000); // Wait for postback/ajax to complete
        });

        // Step 2.2: Verify Search History Grid
        await test.step('Verify Search History Grid', async () => {
            const historyGrid = mainFrame.locator('#ctl00_MainContent_rgdSearchHistory');
            await expect(historyGrid).toBeVisible({ timeout: 20000 });

            // Check if there are any rows in the Master Table or Detail Table
            // Based on elements.txt, there is a master table '#ctl00_MainContent_rgdSearchHistory_ctl00'
            // and potentially detail tables like '#ctl00_MainContent_rgdSearchHistory_ctl00_ctl06_Detail12'
            // We'll look for standard row classes .rgRow or .rgAltRow inside the grid
            const rows = historyGrid.locator('tr.rgRow, tr.rgAltRow');
            
            // Wait for at least one row to be present to confirm data retrieval
            // Note: If the user has NO history, this might fail. 
            // We assume for this test that "Global Tariffs" history exists (based on previous tests).
            await expect(rows.first()).toBeVisible({ timeout: 30000 });
            
            const rowCount = await rows.count();
            console.log(`SUCCESS: Found ${rowCount} history items in the grid.`);
        });

        // Take Screenshot
        const screenshotBufferResult = await page.screenshot({ fullPage: true });
        await test.info().attach(`search_history_grid_${user.id}.png`, {
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
