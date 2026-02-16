import { test, expect } from '@playwright/test';
import { BASE_URL } from './config';
import * as fs from 'fs';
import * as path from 'path';

// Read account information from account.dat
const users = JSON.parse(fs.readFileSync(path.join(__dirname, 'account.dat'), 'utf8'));

test.describe('Role Based Access Control - Search Button Visibility', () => {
  
  // ユーザーの数だけテストケースを自動生成します
  for (const user of users) {
    
    test(`Verify Search Button for ${user.id}`, async ({ page }) => {
      
      console.log(`--- テスト開始: ${user.id} ---`);

      // ---------------------------------------------------
      // Step 1: ログイン処理 (TC_001_org.test.tsベースに修正)
      // ---------------------------------------------------
      await test.step('Login', async () => {
        // 正しいログインURLに変更
        await page.goto(BASE_URL + '/Logon.aspx?ReturnUrl=%2fgtm%2fhome');
        
        // クリックやタブ移動を含む詳細な操作に修正
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
      // Step 2: メインコンテンツへ移動 (TC_001_org.test.tsベースに修正)
      // ---------------------------------------------------
      await test.step('Navigate to Main Content', async () => {
        // ログイン後の画面遷移を追加
        await page.getByRole('button', { name: 'Content' }).click();
        await page.getByRole('menuitem', { name: 'ECN/Dual Use List (Quick' }).click();
        // 最終的に目的のページへ
        await page.goto(BASE_URL + '/gtm/aspx?href=%2FContent%2FfugECCNDetail.aspx');
      });

      // ---------------------------------------------------
      // Step 3: "View As" 等のオーバーレイ対策
      // ---------------------------------------------------
      await test.step('Handle Overlays', async () => {
        const viewAsOverlay = page.locator('text=Select group(s) to view as');
        // isVisible()だとタイムアウトする可能性があるため、低めのタイムアウトでcount()を呼び出す
        if (await viewAsOverlay.count() > 0 && await viewAsOverlay.isVisible()) {
            console.log('⚠️ "View As" パネルを検出。必要に応じて閉じる処理を行ってください。');
        }
      });

      // ---------------------------------------------------
      // Step 4: メイン画面(legacy-outlet)の特定と検証 (TC_001_org.test.tsベースに修正)
      // ---------------------------------------------------
      await test.step('Verify Search Button', async () => {
        const mainFrameElement = page.locator('iframe[name="legacy-outlet"]');
        // 安定性のため、waitForとawait contentFrame()を両方使用
        await mainFrameElement.waitFor({ state: 'attached', timeout: 30000 });
        const mainFrame = await mainFrameElement.contentFrame();
        if (!mainFrame) throw new Error('メインフレームの中身にアクセスできませんでした。');

        const statusBar = mainFrame.locator('#ctl00_MainContent_pnlStatusBarItems');
        const searchBtn = statusBar.locator('a[href*="lnxbtnStatusBarSearch"]');

        // 検証
        await expect(statusBar).toBeVisible({ timeout: 60000 });
        await expect(searchBtn).toBeVisible();
        console.log(`✅ ${user.id}: Searchボタンを確認しました`);

        // 証拠としてページ全体のスクリーンショットを撮影し、レポートに添付
        const screenshotBuffer = await page.screenshot();
        await test.info().attach('page-screenshot', {
          body: screenshotBuffer,
          contentType: 'image/png',
        });
      });

      // ---------------------------------------------------
      // Step 5: ログアウト処理
      // ---------------------------------------------------
      await test.step('Logout', async () => {
        // test-2.spec.tsの記録から取得した正しいログアウト処理
        // ユーザーメニューのボタン名は "MC Test4 " のように動的に変わるため、user.idから生成
        const userButtonName = `MC Test${user.id.substring(6)}`;
        
        // 正規表現を使い、ボタン名の後ろに特殊文字があってもマッチさせる
        await page.getByRole('button', { name: new RegExp(`^${userButtonName}`) }).click();
        await page.getByRole('button', { name: ' Sign out' }).click();
        
        await page.waitForURL(/Logon|default/i);
      });
      
    });
  }
});

