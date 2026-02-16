import { test, expect } from '@playwright/test';
import { BASE_URL } from './config';
import * as fs from 'fs';
import * as path from 'path';

// Load user definitions
const users = JSON.parse(fs.readFileSync(path.join(__dirname, 'account.dat'), 'utf8'));

test.describe('TC_043 Role Based Access Control - Country Information Access', () => {
  test.describe.configure({ mode: 'parallel' });
  
  users.forEach((user: any) => {
    
    test(`Verify Country Information Access for ${user.id}`, async ({ page }) => {
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

      // Step 2: Access Country Information Page
      await test.step('Access Country Information Page', async () => {
        
        await page.getByRole('button', { name: 'Content' }).click();
        
        // Use regex for exact match 'Country Information'
        await page.locator('a[role="menuitem"]').filter({ hasText: /^Country Information$/ }).click();

                let mainFrameElement = page.locator('iframe[name="legacy-outlet"]');
                await mainFrameElement.waitFor({ state: 'attached', timeout: 40000 });
                let mainFrame = await mainFrameElement.contentFrame();
                if (!mainFrame) throw new Error('Could not access main iframe content.');
        
                const pageNameLabel = mainFrame.locator('#lbxPageName');
        
                try {
                  // ページのタイトルが正しく表示されるか検証 (タイムアウトを30秒に延長)
                  await expect(pageNameLabel).toHaveText('Country Information', { timeout: 30000 });
                  console.log('SUCCESS: Verified page name is "Country Information".');
                } catch (error) {
                  console.error('ASSERTION FAILED: Page name was not "Country Information".');
                  // 検証失敗時にデバッグ用のスクリーンショットを撮影
                  await page.screenshot({ path: `tests/failure_screenshot_${user.id}.png`, fullPage: true });
                  console.log(`Debug screenshot saved to tests/failure_screenshot_${user.id}.png`);
                  // エラーを再スローしてテストを失敗させる
                  throw error;
                }
        
                        // elements.txtの3行目の要素が表示されるまで待機
                        await mainFrame.locator('#ctl00_MainContent_cmbxCountryFilter_Input').waitFor({ state: 'visible', timeout: 20000 });
                        console.log('SUCCESS: Country filter input is visible.');
                
                        // Take Screenshot
                        const screenshotBufferResult = await page.screenshot({ fullPage: true });                await test.info().attach(`country_information_page_${user.id}.png`, {
                    body: screenshotBufferResult,
                    contentType: 'image/png'
                });
                console.log('SUCCESS: Page screenshot attached.');      });

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
