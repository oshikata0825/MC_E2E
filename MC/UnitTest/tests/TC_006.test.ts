import { test, expect } from '@playwright/test';
import { BASE_URL } from './config';
import * as fs from 'fs';
import * as path from 'path';

// ベースのユーザー定義はそのまま流用
const users = JSON.parse(fs.readFileSync(path.join(__dirname, 'account.dat'), 'utf8'));

// ランダム選択用のECN番号リスト
const ecnNumberList = ['0A501', '1C350', '3A001', '9E515', '2B350'];

test.describe('Role Based Access Control - ECN Search and Capture', () => {
  
  for (const user of users) {
    
    test(`Perform ECN Search and capture results for ${user.id}`, async ({ page }) => {
      test.setTimeout(180000); // タイムアウトを180秒に設定

      console.log(`--- ECN Search Test Start: ${user.id} ---`);

      // ---------------------------------------------------
      // Step 1: ログイン処理 (変更なし)
      // ---------------------------------------------------
      await test.step('Login', async () => {
        await page.goto(BASE_URL + '/Logon.aspx?ReturnUrl=%2fgtm%2fhome');
        await page.getByRole('textbox', { name: 'Company' }).fill(user.company);
        await page.getByRole('textbox', { name: 'Username' }).fill(user.id);
        await page.getByRole('textbox', { name: 'Password' }).fill(user.pass);
        await page.getByRole('link', { name: 'Login' }).click();
      });

      // ---------------------------------------------------
      // Step 2: ECN検索を実行し、スクリーンショットをキャプチャする
      // ---------------------------------------------------
      await test.step('Perform ECN Search and Capture', async () => {
        // ECN/Dual Use List ページに移動
        await page.getByRole('button', { name: 'Content' }).click();
        await page.getByText('ECN/Dual Use List', { exact: true }).click();

        // iframeを取得
        let mainFrameElement = page.locator('iframe[name="legacy-outlet"]');
        await mainFrameElement.waitFor({ state: 'attached', timeout: 30000 });
        let mainFrame = await mainFrameElement.contentFrame();
        if (!mainFrame) throw new Error('メインのiframeコンテンツにアクセスできませんでした。');

        // Regulation Listを選択
        await mainFrame.locator('#ctl00_MainContent_cmbxRegulationList_Input').click();
        await mainFrame.locator('#ctl00_MainContent_cmbxRegulationList_DropDown').getByText('US - UNITED STATES OF AMERICA - EAR/Commerce Control List').click();

        // 選択後のページリロードを待機
        console.log('ℹ️ Regulation List選択後のページリロードを待機中...');
        mainFrameElement = page.locator('iframe[name="legacy-outlet"]');
        await mainFrameElement.waitFor({ state: 'attached', timeout: 30000 });
        mainFrame = await mainFrameElement.contentFrame();
        if (!mainFrame) throw new Error('リロード後のiframeコンテンツにアクセスできませんでした。');
        await expect(mainFrame.locator('#ctl00_MainContent_cmbxRegulationList_Input')).toBeVisible({ timeout: 30000 });
        console.log('✅ ページがリロードされました。');

        // ECN番号を入力し、オートコンプリートから選択して検索する
        const ecnToType = ecnNumberList[0]; // デバッグのため、リストの最初のECNを確定で使用
        console.log(`ℹ️ ECN番号を入力: ${ecnToType}`);
        const ecnInput = mainFrame.locator('#txtbxECN');
        await expect(ecnInput).toBeVisible({ timeout: 10000 });
        await ecnInput.click();
        await ecnInput.pressSequentially(ecnToType, { delay: 100 });

        // 入力したテキストを含む候補(li要素)が表示されるのを直接待つ
        console.log(`ℹ️ テキスト「${ecnToType}」を含む候補が表示されるのを待機中...`);
        const suggestionItem = mainFrame.locator('li', { hasText: ecnToType });
        await expect(suggestionItem.first()).toBeVisible({ timeout: 15000 });
        
        console.log('ℹ️ 見つかった最初の候補をクリックします。');
        await suggestionItem.first().click();

        // 検索結果の表示を待機
        console.log('ℹ️ オートコンプリート選択後の画面遷移を待機中...');
        const resultsHierarchy = mainFrame.locator('#ctl00_MainContent_rtvECNHierarchy');
        await expect(resultsHierarchy).toBeVisible({ timeout: 30000 });


        // ユーザー指定のIDを持つ最初の結果ノードの要素が表示されるのを待つ
        console.log('ℹ️ 検索結果の最初のノードが表示されるのを待機中...');
        await expect(mainFrame.locator('#ctl00_MainContent_rtvECNHierarchy_i0_lblECNNumber')).toBeVisible({ timeout: 10000 });
        await expect(mainFrame.locator('#ctl00_MainContent_rtvECNHierarchy_i0_lblECNDescription')).toBeVisible({ timeout: 10000 });

        // 描画が安定するまで固定で待機
        console.log('ℹ️ 描画の安定を待つため3秒間待機します...');
        await page.waitForTimeout(3000);

        console.log('✅ 検索結果が完全に読み込まれました。');

        // スクリーンショットを撮影し、レポートに添付
        console.log('ℹ️ 検索結果のスクリーンショットを撮影中...');
        const screenshotBuffer = await page.screenshot({ fullPage: true });
        await test.info().attach(`${user.id}-ECN-Search-Result.png`, {
          body: screenshotBuffer,
          contentType: 'image/png',
        });
      });

      // ---------------------------------------------------
      // Step 3: ログアウト処理 (変更なし)
      // ---------------------------------------------------
      await test.step('Logout', async () => {
        const userButtonName = `MC Test${user.id.substring(6)}`;
        await page.bringToFront();

        // ユーザーメニューボタンを待機してクリック
        const userMenuButton = page.getByRole('button', { name: new RegExp(`^${userButtonName}`) });
        await userMenuButton.waitFor({ state: 'visible', timeout: 10000 });
        await userMenuButton.click();

        // サインアウトボタンが表示されるのを待機
        const signOutButton = page.getByRole('button', { name: ' Sign out' });
        await signOutButton.waitFor({ state: 'visible', timeout: 5000 });

        // 強制クリックで、他の要素に隠されている場合でも安定してクリックする
        await signOutButton.click({ force: true }); 
        
        await expect(page.getByRole('textbox', { name: 'Company' })).toBeVisible({ timeout: 30000 });
      });
      
    });
  }
});

