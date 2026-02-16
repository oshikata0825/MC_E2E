import { test, expect } from '@playwright/test';
import { BASE_URL } from './config';
import * as fs from 'fs';
import * as path from 'path';

// ベースのユーザー定義はそのまま流用
const users = JSON.parse(fs.readFileSync(path.join(__dirname, 'account.dat'), 'utf8'));

test.describe('Role Based Access Control - ECN Query Page Verification', () => {
  
  for (const user of users) {
    
    test(`Verify ECN Query page title for ${user.id}`, async ({ page }) => {
      test.setTimeout(180000); // タイムアウトを180秒に設定

      console.log(`--- ECN Query Page Test Start: ${user.id} ---`);

      // ---------------------------------------------------
      // Step 1: ログイン処理
      // ---------------------------------------------------
      await test.step('Login', async () => {
        await page.goto(BASE_URL + '/Logon.aspx?ReturnUrl=%2fgtm%2fhome');
        await page.getByRole('textbox', { name: 'Company' }).fill(user.company);
        await page.getByRole('textbox', { name: 'Username' }).fill(user.id);
        await page.getByRole('textbox', { name: 'Password' }).fill(user.pass);
        await page.getByRole('link', { name: 'Login' }).click();
      });

      // ---------------------------------------------------
      // Step 2: ECN Queryページに移動し、タイトルを検証する
      // ---------------------------------------------------
      await test.step('Navigate to ECN Query and Verify Title', async () => {
        await page.getByRole('button', { name: 'Content' }).click();
        // 'ECN/Dual Use List Query' メニュー項目をクリック
        await page.getByRole('menuitem', { name: 'ECN/Dual Use List Query' }).click();

        // メインのiframeを取得
        const mainFrameElement = page.locator('iframe[name="legacy-outlet"]');
        await mainFrameElement.waitFor({ state: 'attached', timeout: 30000 });
        const mainFrame = await mainFrameElement.contentFrame();
        if (!mainFrame) throw new Error('メインのiframeコンテンツにアクセスできませんでした。');

        // ページタイトル要素を検証
        console.log('ℹ️ ページタイトルを検証します...');
        const pageTitle = mainFrame.locator('#lbxPageName');
        await expect(pageTitle).toBeVisible({ timeout: 20000 });
        await expect(pageTitle).toHaveText('ECN Query');
        console.log('✅ ページタイトルが正しいことを確認しました: "ECN Query"');

        // スクリーンショットを撮影
        console.log('ℹ️ ECN Queryページのスクリーンショットを撮影します。');
        const screenshotBuffer = await page.screenshot({ fullPage: true });
        await test.info().attach(`${user.id}-ECN-Query-Page.png`, {
          body: screenshotBuffer,
          contentType: 'image/png',
        });
      });

      // ---------------------------------------------------
      // Step 3: ログアウト処理
      // ---------------------------------------------------
      await test.step('Logout', async () => {
        const userButtonName = `MC Test${user.id.substring(6)}`;
        await page.bringToFront();

        const userMenuButton = page.getByRole('button', { name: new RegExp(`^${userButtonName}`) });
        await userMenuButton.waitFor({ state: 'visible', timeout: 10000 });
        await userMenuButton.click();

        const signOutButton = page.getByRole('button', { name: ' Sign out' });
        await signOutButton.waitFor({ state: 'visible', timeout: 5000 });
        
        // 他要素に隠されている場合でもクリックを続行
        await signOutButton.click({ force: true }); 
        
        await expect(page.getByRole('textbox', { name: 'Company' })).toBeVisible({ timeout: 30000 });
      });
      
    });
  }
});
