import { test, expect } from '@playwright/test';
import { BASE_URL } from './config';
import * as fs from 'fs';
import * as path from 'path';

// Read account information from account.dat
const users = JSON.parse(fs.readFileSync(path.join(__dirname, 'account.dat'), 'utf8'));

// テスト用のECCNリスト
const eccnList = ['5A002', '3E001', '9A001', '9E003', '0A979', 'EAR99'];

test.describe('Role Based Access Control - Search Functionality', () => {
  
  for (const user of users) {
    
    test(`Verify Search Functionality for ${user.id}`, async ({ page }) => {
      
      console.log(`--- テスト開始: ${user.id} ---`);

      // ---------------------------------------------------
      // Step 1: ログイン処理
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
      // Step 2: メインコンテンツへ移動
      // ---------------------------------------------------
      await test.step('Navigate to Main Content', async () => {
        await page.getByRole('button', { name: 'Content' }).click();
        await page.getByRole('menuitem', { name: 'ECN/Dual Use List (Quick' }).click();
        await page.goto(BASE_URL + '/gtm/aspx?href=%2FContent%2FfugECCNDetail.aspx');
      });

      // ---------------------------------------------------
      // Step 3: "View As" 等のオーバーレイ対策
      // ---------------------------------------------------
      await test.step('Handle Overlays', async () => {
        const viewAsOverlay = page.locator('text=Select group(s) to view as');
        if (await viewAsOverlay.count() > 0 && await viewAsOverlay.isVisible()) {
            console.log('⚠️ "View As" パネルを検出。必要に応じて閉じる処理を行ってください。');
        }
      });

      // ---------------------------------------------------
      // Step 4: 検索ボタンの確認と検索実行
      // ---------------------------------------------------
      await test.step('Verify and Perform Search', async () => {
        const mainFrameElement = page.locator('iframe[name="legacy-outlet"]');
        await mainFrameElement.waitFor({ state: 'attached', timeout: 30000 });
        const mainFrame = await mainFrameElement.contentFrame();
        if (!mainFrame) throw new Error('メインフレームの中身にアクセスできませんでした。');

        const statusBar = mainFrame.locator('#ctl00_MainContent_pnlStatusBarItems');
        const searchBtn = statusBar.locator('a[href*="lnxbtnStatusBarSearch"]');

        await expect(statusBar).toBeVisible({ timeout: 60000 });

        // Searchボタンが表示されているか確認
        if (await searchBtn.isVisible()) {
          console.log(`✅ ${user.id}: Searchボタンを検出。検索機能をテストします。`);
          
          // --- 検索操作 (test-1.spec.tsベース) ---
          
          // 1. 国リストを選択
          await mainFrame.locator('#ctl00_MainContent_cmbxRegulationList_Arrow').click();
          await mainFrame.getByText('US - UNITED STATES OF AMERICA - EAR/Commerce Control List').click();
          
          // 2. ランダムなECCNを選択して入力
          const randomECCN = eccnList[Math.floor(Math.random() * eccnList.length)];
          console.log(`ℹ️ 使用するECCN: ${randomECCN}`);
          const eccnTextbox = mainFrame.getByRole('textbox', { name: 'Enter an ECN Number or' });
          await eccnTextbox.click();
          await eccnTextbox.fill(randomECCN);
          
          // 3. オートコンプリートの候補をクリックし、それが消えるのを待つ
          const suggestionLocator = mainFrame.getByText(new RegExp(`^${randomECCN}\\s-`));
          try {
            await suggestionLocator.first().click({ timeout: 5000 });
            console.log('ℹ️ オートコンプリート候補をクリックしました。');

            // 候補リストが消えるのを待つことで、検索が開始されたことを確認
            await expect(suggestionLocator.first()).not.toBeVisible({ timeout: 10000 });
            console.log('ℹ️ オートコンプリート候補が消え、検索が開始されました。');

          } catch (e) {
            console.log(`ℹ️ "${randomECCN}" のオートコンプリート候補が見つかりませんでした。`);
          }

          // 4. 検索結果のヘッダー（同じテキストだが、別の要素）が表示されるまで待つ
          console.log('ℹ️ 検索結果が表示されるのを待っています...');
          const resultHeaderLocator = mainFrame.getByText(new RegExp(`^${randomECCN}\\s-`));
          await expect(resultHeaderLocator.first()).toBeVisible({ timeout: 30000 });
          const headerText = await resultHeaderLocator.first().textContent();
          console.log(`ℹ️ 確認した結果ヘッダー: ${headerText?.trim()}`);
          
          // 5. 成功の証拠としてスクリーンショットを撮影
          const screenshotBuffer = await page.screenshot();
          await test.info().attach(`${user.id}-search-result.png`, {
            body: screenshotBuffer,
            contentType: 'image/png',
          });

          // 6. "Exit"ボタンをクリックして元の画面に戻る
          const exitButton = mainFrame.getByRole('link', { name: 'Exit' });
          await exitButton.click(); 
          await page.waitForLoadState('domcontentloaded');

        } else {
          console.log(`⏭️ ${user.id}: Searchボタンが見つかりません。このロールのテストはスキップします。`);
        }
      });

      // ---------------------------------------------------
      // Step 5: ログアウト処理
      // ---------------------------------------------------
      await test.step('Logout', async () => {
        const userButtonName = `MC Test${user.id.substring(6)}`;
        await page.getByRole('button', { name: new RegExp(`^${userButtonName}`) }).click();
        await page.getByRole('button', { name: ' Sign out' }).click();
                
                // ログアウト成功の確認を、URL遷移ではなく、ログイン画面の要素が表示されるかで判断する
                await expect(page.getByRole('textbox', { name: 'Company' })).toBeVisible({ timeout: 30000 });
              });
      
    });
  }
});
