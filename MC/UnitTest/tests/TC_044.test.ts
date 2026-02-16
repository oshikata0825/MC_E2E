import { test, expect } from '@playwright/test';
import { BASE_URL } from './config';
import * as fs from 'fs';
import * as path from 'path';

// Load user definitions
const users = JSON.parse(fs.readFileSync(path.join(__dirname, 'account.dat'), 'utf8'));

test.describe('TC_044 Role Based Access Control - Country Information - Filter Verification', () => {
  test.describe.configure({ mode: 'parallel' });
  
  users.forEach((user: any) => {
    
    test(`Perform Country Filter verification for ${user.id}`, async ({ page }) => {
      test.setTimeout(180000); // Timeoutを延長

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
      const mainFrame = await test.step('Access Country Information Page', async () => {
        await page.getByRole('button', { name: 'Content' }).click();
        await page.locator('a[role="menuitem"]').filter({ hasText: /^Country Information$/ }).click();

        const mainFrameElement = page.locator('iframe[name="legacy-outlet"]');
        await mainFrameElement.waitFor({ state: 'attached', timeout: 40000 });
        const frame = await mainFrameElement.contentFrame();
        if (!frame) throw new Error('Could not access main iframe content.');

        const pageNameLabel = frame.locator('#lbxPageName');
        await expect(pageNameLabel).toHaveText('Country Information', { timeout: 30000 });
        
        return frame;
      });

      // Step 3: Select Country and Verify Tabs
      await test.step('Select Country and Verify Tabs', async () => {
        // 国フィルターの矢印をクリックしてドロップダウンを開く
        const countryFilterArrow = mainFrame.locator('#ctl00_MainContent_cmbxCountryFilter_Arrow');
        await countryFilterArrow.click();

        // ドロップダウンが表示されるのを待つ
        const dropdown = mainFrame.locator('#ctl00_MainContent_cmbxCountryFilter_DropDown');
        await dropdown.waitFor({ state: 'visible', timeout: 10000 });

        // 国のリストを取得
        const countryItems = dropdown.locator('li.rcbItem');
        const countryCount = await countryItems.count();
        if (countryCount <= 1) { // プレースホルダーのみ、または1つしか選択肢がない場合
          throw new Error('Not enough countries found in the filter to perform random selection.');
        }
        
        // ランダムな国を選択 (最初の"Select..."項目は除く)
        const randomIndex = Math.floor(Math.random() * (countryCount - 1)) + 1;
        const selectedCountryItem = countryItems.nth(randomIndex);
        const selectedCountryName = await selectedCountryItem.innerText();
        console.log(`Selecting random country: ${selectedCountryName}`);
        await selectedCountryItem.click();

        // タブが表示されるのを待つ
        const tabContainer = mainFrame.locator('div.rtsLevel.rtsLevel1 ul.rtsUL');
        await tabContainer.waitFor({ state: 'visible', timeout: 30000 });
        console.log('Tab container is visible.');

        // タブを動的に読み取ってスクリーンショットを撮る
        const tabs = tabContainer.locator('li.rtsLI');
        const tabCount = await tabs.count();
        console.log(`Found ${tabCount} tabs.`);

        if (tabCount === 0) {
            console.log('No tabs found for the selected country. Test will pass.');
            return;
        }

        for (let i = 0; i < tabCount; i++) {
            const tab = tabs.nth(i);
            const tabText = await tab.locator('span.rtsTxt').innerText();
            console.log(`Processing tab #${i + 1}: ${tabText}`);

            await tab.click();
            // コンテンツの描画を少し待つ
            await page.waitForTimeout(1500);

            const screenshot = await page.screenshot();
            await test.info().attach(`country_tab_${tabText.replace(/[\/\s]/g, '_')}_${user.id}.png`, {
                body: screenshot,
                contentType: 'image/png'
            });
        }
      });
      
      // Step 4: Logout
      await test.step('Logout', async () => {
        const userButtonName = `MC Test${user.id.substring(6)}`;
        await page.bringToFront();

        const userButton = page.getByRole('button', { name: new RegExp(`^${userButtonName}`) });
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