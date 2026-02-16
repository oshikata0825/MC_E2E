import { test, expect, Page } from '@playwright/test';
import { BASE_URL } from './config';
import * as fs from 'fs';
import * as path from 'path';

// Load user definitions
const allUsers = JSON.parse(fs.readFileSync(path.join(__dirname, 'account.dat'), 'utf8'));
const users = allUsers.filter((user: any) => user.id === 'MCTest9');

test.describe('MC_DPS_UI_005 Denied Party Screening - MC Company Lookup UI Verification', () => {
  test.describe.configure({ mode: 'parallel' });
  
  users.forEach((user: any) => {
    
    test(`Perform UI verification for ${user.id}`, async ({ page }, testInfo) => {
      test.setTimeout(600000); 

      console.log(`--- Test Start: ${user.id} ---`);

      // Step 1: Login
      await test.step('Login', async () => {
        await page.goto(BASE_URL + '/Logon.aspx?ReturnUrl=%2fgtm%2fhome');
        await page.getByRole('textbox', { name: 'Company' }).fill(user.company);
        await page.getByRole('textbox', { name: 'Username' }).fill(user.id);
        await page.getByRole('textbox', { name: 'Password' }).fill(user.pass);
        await page.getByRole('link', { name: 'Login' }).click();
      });

      // Step 2: Access MC Company Lookup Page
      await test.step('Access MC Company Lookup Page', async () => {
        await page.getByRole('button', { name: 'DPS' }).click();
        await page.locator('a[role="menuitem"]').filter({ hasText: /^DPS Search\/Reporting Lookup$/i }).click();

        const mainFrameElement = page.locator('iframe[name="legacy-outlet"]');
        await mainFrameElement.waitFor({ state: 'attached', timeout: 40000 });
        const frame = await mainFrameElement.contentFrame();
        if (!frame) throw new Error('Could not access main iframe content.');

        // Using selector from elements.txt: <span class="mm-text">MC Company Lookup</span>
        const lookupLink = frame.locator('span.mm-text:has-text("MC Company Lookup")');
        await expect(lookupLink).toBeVisible({ timeout: 20000 });
        await lookupLink.click();
        await page.waitForTimeout(3000);
      });

      // Step 3: Filter by CreatedBy = MDM and open Edit form
      let popup: Page | null = null;
      await test.step('Filter by CreatedBy=MDM and open Edit form', async () => {
        const mainFrameElement = page.locator('iframe[name="legacy-outlet"]');
        const frame = await mainFrameElement.contentFrame();
        if (!frame) throw new Error('Could not access main iframe content.');

        console.log('Filtering for CreatedBy=MDM...');

        const targetHeader = frame.locator('th.rgHeader').filter({ hasText: /^CreatedBy$/i });
        await targetHeader.locator('button.rgActionButton.rgOptions').first().click();
        await page.waitForTimeout(2000);

        const condInput = frame.locator('input[id*="HCFMRCMBFirstCond_Input"]');
        await condInput.click();
        await page.waitForTimeout(1000);
        
        await frame.locator('li.rcbItem', { hasText: /^EqualTo$/i }).or(page.locator('li.rcbItem', { hasText: /^EqualTo$/i })).first().click();
        await page.waitForTimeout(1000);

        const valInput = frame.locator('input[id*="HCFMRTBFirstCond"]').first();
        await valInput.click();
        await valInput.fill('');
        await valInput.pressSequentially('MDM', { delay: 100 });
        await valInput.press('Tab');
        await page.waitForTimeout(1000);

        const filterBtn = frame.locator('input[type="submit"][value="Filter"], button:has-text("Filter"), .rgFilterButton, span.rgButtonText:has-text("Filter")').first();
        await filterBtn.click({ force: true });
        
        console.log('Waiting for CreatedBy filter to be applied (5s)...');
        await page.waitForTimeout(5000);

        // --- Add SanFlag filter ---
        console.log('Filtering for SanFlag=Y-はい...');
        const sanFlagHeader = frame.locator('th.rgHeader').filter({ hasText: /^SanFlag$/i });
        await sanFlagHeader.locator('button.rgActionButton.rgOptions').first().click();
        await page.waitForTimeout(2000);

        const sfCondInput = frame.locator('input[id*="HCFMRCMBFirstCond_Input"]');
        await sfCondInput.click();
        await page.waitForTimeout(1000);
        
        await frame.locator('li.rcbItem', { hasText: /^EqualTo$/i }).or(page.locator('li.rcbItem', { hasText: /^EqualTo$/i })).first().click();
        await page.waitForTimeout(1000);

        const sfValInput = frame.locator('input[id*="HCFMRTBFirstCond"]').first();
        await sfValInput.click();
        await sfValInput.fill('');
        await sfValInput.pressSequentially('Y-はい', { delay: 100 });
        await sfValInput.press('Tab');
        await page.waitForTimeout(1000);

        const sfFilterBtn = frame.locator('input[type="submit"][value="Filter"], button:has-text("Filter"), .rgFilterButton, span.rgButtonText:has-text("Filter")').first();
        await sfFilterBtn.click({ force: true });

        console.log('Waiting for SanFlag filter to be applied (5s)...');
        await page.waitForTimeout(5000);

        // Attach screenshot to verify table content after filtering
        await testInfo.attach('after_filters_applied.png', {
            body: await page.screenshot({ fullPage: true }),
            contentType: 'image/png'
        });

        // フィルターメニューを閉じる
        await page.keyboard.press('Escape');
        await page.waitForTimeout(1000);

        // フィルタリング後の MDM 行を特定して Edit リンクを取得
        const mdmRows = frame.locator('tr.rgRow, tr.rgAltRow').filter({ hasText: 'MDM' });
        const mdmCount = await mdmRows.count();
        console.log(`After filtering: Found ${mdmCount} rows containing "MDM".`);

        if (mdmCount > 0) {
          const firstMdmRow = mdmRows.first();
          const link = firstMdmRow.locator('a:has-text("Edit")');
          
          console.log('Targeting MDM row Edit link...');
          await link.scrollIntoViewIfNeeded();
          await page.waitForTimeout(1000);

          const popupPromise = page.context().waitForEvent('page', { timeout: 30000 });
          // Enterキーまたはクリックでポップアップを開く
          await link.focus();
          await page.keyboard.press('Enter');
          popup = await popupPromise;

          await popup.waitForLoadState('load');
          await popup.waitForTimeout(5000); 

          console.log('MDM record form opened.');
          await testInfo.attach('mdm_record_form.png', {
            body: await popup.screenshot({ fullPage: true }),
            contentType: 'image/png'
          });
        } else {
          const errorMsg = 'FAILED: No test data found matching CreatedBy=MDM and SanFlag=Y-はい.';
          console.error(errorMsg);
          throw new Error(errorMsg);
        }
      });

      // Step 4: Logout
      await test.step('Logout', async () => {
        if (popup) {
            await popup.close();
        }
        
        const userButtonName = `MC Test${user.id.substring(6)}`;
        await page.bringToFront();

        const userButton = page.getByRole('button', { name: new RegExp(`^${userButtonName}`) });
        await userButton.hover();
        await page.waitForTimeout(500);
        await userButton.click();

        const logoutButton = page.locator('#logoutButton');
        await expect(logoutButton).toBeVisible({ timeout: 10000 });
        await logoutButton.click({ force: true });
      });
    });
  });
});
