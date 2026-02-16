import { test, expect } from '@playwright/test';
import { BASE_URL } from './config';
import * as fs from 'fs';
import * as path from 'path';

// ベースのユーザー定義はそのまま流用
const users = JSON.parse(fs.readFileSync(path.join(__dirname, 'account.dat'), 'utf8'));

// elements.txtから読み込んだ要素IDのリスト
const elementIds = [
  'ctl00_MainContent_cmbxCultureCode_Input',
  'ctl00_MainContent_rdpEffectiveDatePicker_dateInput',
  'ctl00_MainContent_cmbxRegulationList_Input',
  'txtbxECN',
  'lnxbtnSearch'
];

test.describe('Role Based Access Control - ECN/Dual Use List UI Verification', () => {
  
  for (const user of users) {
    
    test(`Verify UI elements on ECN/Dual Use List page for ${user.id}`, async ({ page }) => {
      test.setTimeout(180000); // タイムアウトを180秒(3分)に延長
      
      console.log(`--- UI Element Verification Test Start: ${user.id} ---`);

      // ---------------------------------------------------
      // Step 1: ログイン処理 (変更なし)
      // ---------------------------------------------------
      await test.step('Login', async () => {
        await page.goto(BASE_URL + '/Logon.aspx?ReturnUrl=%2fgtm%2fhome');
        await page.getByRole('textbox', { name: 'Company' }).click();
        await page.getByRole('textbox', { name: 'Company' }).fill(user.company);
        await page.getByRole('textbox', { name: 'Company' }).press('Tab');
        await page.getByRole('textbox', { name: 'Username' }).click();
        await page.getByRole('textbox', { name: 'Username' }).fill(user.id);
        await page.getByRole('textbox', { name: 'Password' }).click();
        await page.getByRole('textbox', { name: 'Password' }).fill(user.pass);
        await page.getByRole('link', { name: 'Login' }).click();
      });

      // ---------------------------------------------------
      // Step 2: ECN/Dual Use List画面に遷移し、UI要素を確認する
      // ---------------------------------------------------
      await test.step('Navigate and Verify UI Elements', async () => {
        // 1. Contentメニューをクリック
        await page.getByRole('button', { name: 'Content' }).click();
        
        // 2. 'ECN/Dual Use List' メニュー項目をクリック
        await page.getByText('ECN/Dual Use List', { exact: true }).click();

        // 3. iframeを待機し、そのコンテンツフレームを取得
        let mainFrameElement = page.locator('iframe[name="legacy-outlet"]');
        await mainFrameElement.waitFor({ state: 'attached', timeout: 30000 });
        let mainFrame = await mainFrameElement.contentFrame();
        if (!mainFrame) throw new Error('メインフレームのコンテンツにアクセスできませんでした。');

        // 4. リストにある各要素が表示されていることを確認
        console.log('ℹ️ 要素の可視性を確認中...');
        for (const elementId of elementIds) {
          console.log(`    -> 確認中: #${elementId}`);
          const elementLocator = mainFrame.locator(`#${elementId}`);
          await expect(elementLocator).toBeVisible({ timeout: 20000 });
          console.log(`    ✅ 可視: #${elementId}`);
        }
        console.log('✅ 指定されたすべての要素が表示されています。');

        // 5. 証拠としてスクリーンショットを撮影
        console.log('ℹ️ 確認された要素のスクリーンショットを撮影中。');
        const screenshotBuffer = await page.screenshot({ fullPage: true });
        await test.info().attach(`${user.id}-ECN-UI-Verification.png`, {
          body: screenshotBuffer,
          contentType: 'image/png',
        });
      });


      // ---------------------------------------------------
      // Step 3: ログアウト処理 (変更なし)
      // ---------------------------------------------------
      await test.step('Logout', async () => {
        const userButtonName = `MC Test${user.id.substring(6)}`;
        // iframeからメインページにフォーカスを戻してからトップメニューを操作する
        await page.bringToFront();
        await page.getByRole('button', { name: new RegExp(`^${userButtonName}`) }).click();
        await page.getByRole('button', { name: ' Sign out' }).click();
        await expect(page.getByRole('textbox', { name: 'Company' })).toBeVisible({ timeout: 30000 });
      });
      
    });
  }
});
