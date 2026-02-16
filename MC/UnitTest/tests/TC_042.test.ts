import { test, expect } from '@playwright/test';
import { BASE_URL } from './config';
import * as fs from 'fs';
import * as path from 'path';

// Load user definitions
const users = JSON.parse(fs.readFileSync(path.join(__dirname, 'account.dat'), 'utf8'));

test.describe('TC_042 Role Based Access Control - Search History Verification with Retry', () => {
  test.describe.configure({ mode: 'parallel' });
  
  users.forEach((user: any) => {
    
    test(`Perform Search History verification with retry for ${user.id}`, async ({ page }) => {
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

        // Step 2.1: Select 'Recent Searches Type' with Retry
        let historyFound = false;
        const maxRetries = 5;
        const triedTypes = new Set<string>();

        for (let attempt = 1; attempt <= maxRetries; attempt++) {
            await test.step(`Attempt ${attempt}: Select Recent Searches Type`, async () => {
                console.log(`--- Attempt ${attempt} of ${maxRetries} ---`);

                // Target the arrow button to open the dropdown.
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

                // Find items that haven't been tried yet
                const availableIndices: number[] = [];
                for (let i = 0; i < count; i++) {
                    const text = await items.nth(i).innerText();
                    if (!triedTypes.has(text)) {
                        availableIndices.push(i);
                    }
                }

                if (availableIndices.length === 0) {
                    console.log('All available types have been tried. resetting tried list for randomness (or could stop).');
                    // Optional: clear triedTypes if we want to allow retrying same types, 
                    // but logic suggests we should just pick one randomly if we ran out of unique ones.
                    // For now, let's just pick any random one if we exhausted unique ones (though unlikely with 5 retries and many types).
                     for (let i = 0; i < count; i++) availableIndices.push(i);
                }

                // Select random item from available
                const randomIndex = availableIndices[Math.floor(Math.random() * availableIndices.length)];
                const itemToSelect = items.nth(randomIndex);
                const itemText = await itemToSelect.innerText();
                
                console.log(`Selecting Recent Search Type: ${itemText}`);
                triedTypes.add(itemText);
                await itemToSelect.click();

                // Wait for the grid to update
                await page.waitForTimeout(3000); 

                // Verify "View" links
                // Based on elements.txt: <a ... role="button">View</a>
                const viewLinks = mainFrame.locator('a[role="button"]').filter({ hasText: 'View' });
                const linkCount = await viewLinks.count();

                if (linkCount > 0) {
                    console.log(`SUCCESS: Found ${linkCount} history items with 'View' link.`);
                    historyFound = true;
                } else {
                    console.log('No history found for this type.');
                }
            });

            if (historyFound) break;
        }

        if (!historyFound) {
            throw new Error(`Failed to find search history after ${maxRetries} attempts.`);
        }

        // Step 3: Click random 'View' link and verify detail page
        await test.step('Click View Link and Verify Detail Page', async () => {
            // Re-locate links to ensure freshness (though we just found them)
            const viewLinks = mainFrame.locator('a[role="button"]').filter({ hasText: 'View' });
            const count = await viewLinks.count();
            
            // Should be guaranteed > 0 by previous logic, but safe check
            if (count === 0) throw new Error('Unexpected error: View links lost.');

            const randomIndex = Math.floor(Math.random() * count);
            const linkToClick = viewLinks.nth(randomIndex);
            
            console.log(`Clicking View link at index ${randomIndex}`);
            
            // Click and wait for navigation. 
            // The href indicates a page navigation within the frame or a new window.
            // Assuming same frame navigation based on standard web app behavior for such links unless target=_blank
            await linkToClick.click();
            
            // Wait for load
            // Since navigation happens inside the iframe, the old frame execution context might be destroyed.
            // We need to wait for the iframe to reload/navigate and then get the new frame reference.
            
            // Wait for the iframe to potentially detach/attach or just wait a bit and re-grab
            await page.waitForTimeout(5000); 

            // Re-acquire the frame
            mainFrameElement = page.locator('iframe[name="legacy-outlet"]');
            mainFrame = await mainFrameElement.contentFrame();
            if (!mainFrame) throw new Error('Could not access main iframe content after navigation.');
            
            // Wait for the page name label to be visible in the new frame context
            // This confirms the DOM inside the iframe is accessible
            await mainFrame.locator('#lbxPageName').waitFor({ state: 'visible', timeout: 30000 });
            
            console.log('Detail page loaded (verified via page name visibility).');

            // Take Screenshot of the detail view
            const screenshotBufferResult = await page.screenshot({ fullPage: true });
            await test.info().attach(`search_history_detail_view_${user.id}.png`, {
                body: screenshotBufferResult,
                contentType: 'image/png'
            });
            console.log('SUCCESS: Detail view screenshot attached.');
        });

        console.log('SUCCESS: Test complete.');
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
