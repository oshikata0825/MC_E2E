import { test, expect } from '@playwright/test';
import { BASE_URL } from './config';
import * as fs from 'fs';
import * as path from 'path';

// ベースのユーザー定義はそのまま流用
const users = JSON.parse(fs.readFileSync(path.join(__dirname, 'account.dat'), 'utf8'));

// 検索用のキーワードリスト
const keywordList = ['system', 'sensor', 'engine', 'laser', 'telecommunications', 'security', 'equipment'];

test.describe('Role Based Access Control - Keyword Search Detail Modal Verification', () => {
  
  for (const user of users) {
    
    test(`Verify search result detail modal can be opened for ${user.id}`, async ({ page }) => {
      test.setTimeout(180000); // タイムアウトを180秒に設定

      console.log(`--- Keyword Search Detail Modal Test Start: ${user.id} ---`);

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
      // Step 2: キーワード検索から詳細表示とクリーンアップ
      // ---------------------------------------------------
      await test.step('Perform Search, Open Modal, and Cleanup', async () => {
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
        mainFrameElement = page.locator('iframe[name="legacy-outlet"]');
        await mainFrameElement.waitFor({ state: 'attached', timeout: 30000 });
        mainFrame = await mainFrameElement.contentFrame();
        if (!mainFrame) throw new Error('リロード後のiframeコンテンツにアクセスできませんでした。');
        await expect(mainFrame.locator('#ctl00_MainContent_cmbxRegulationList_Input')).toBeVisible({ timeout: 30000 });
        console.log('✅ ページがリロードされました。');

        // [ROBUSTNESS] 前回の検索結果が表示されている場合、"New Search" をクリックしてリセットする
        const resultsHierarchy = mainFrame.locator('#ctl00_MainContent_rtvECNHierarchy');
        if (await resultsHierarchy.isVisible()) {
            console.log('ℹ️ 前回の検索結果が残っているため、"New Search" をクリックします。');
            await mainFrame.locator('#hyxlnkNewSearch').click();
            // ページがリセットされ、結果ツリーが非表示になるのを待つ
            await expect(resultsHierarchy).toBeHidden();
            console.log('✅ ページがリセットされました。');
        }

        // キーワードを入力
        const randomKeyword = keywordList[Math.floor(Math.random() * keywordList.length)];
        console.log(`ℹ️ キーワードを入力: ${randomKeyword}`);
        const searchInput = mainFrame.locator('#txtbxECN');
        await expect(searchInput).toBeVisible();
        await searchInput.click();
        await searchInput.pressSequentially(randomKeyword, { delay: 100 });

        // オートコンプリートから候補を選択
        const dropdown = mainFrame.locator('ul#ECN_completionListElem');
        await expect(dropdown).toBeVisible({ timeout: 15000 });
        const suggestions = dropdown.locator('li');
        await expect(suggestions.first()).toBeVisible({ timeout: 10000 });
        const count = await suggestions.count();
        expect(count).toBeGreaterThan(0);

        const firstSuggestionText = await suggestions.first().innerText();
        if (firstSuggestionText.includes('ECN Quick Search Keywords')) {
            expect(count, '最初の候補が"Quick Search"ですが、他に選択肢がありません。').toBeGreaterThan(1);
            await suggestions.nth(1).click();
        } else {
            await suggestions.first().click();
        }

        // 検索結果ツリーの表示を待機
        await expect(resultsHierarchy).toBeVisible({ timeout: 30000 });
        await page.waitForTimeout(3000); // 描画が安定するまで待機
        console.log('✅ 検索結果が読み込まれました。');

        // --- "View Details"リンクをクリック ---
        console.log('ℹ️ 検索結果内の"View Details"リンクを探してクリックします...');
        const viewDetailsLinks = mainFrame.locator('#ctl00_MainContent_rtvECNHierarchy').getByRole('link', { name: 'View Details' });
        const linkCount = await viewDetailsLinks.count();
        expect(linkCount).toBeGreaterThan(0);
        
        const randomIndex = Math.floor(Math.random() * linkCount);
        const targetLink = viewDetailsLinks.nth(randomIndex);

        console.log(`-> ${linkCount}個中の${randomIndex + 1}番目のリンクをスクロールしてクリックします。`);
        await targetLink.scrollIntoViewIfNeeded();
        await page.waitForTimeout(500); // スクロールが完了するのを待つ
        await targetLink.click(); // 通常のクリックを実行

        // --- モーダル表示とコンテンツの読み込みを待機 ---
        console.log('ℹ️ モーダル内のコンテンツ(iframe)が読み込まれるのを待機中...');
        const modalContentFrame = mainFrame.frameLocator('iframe[name="rwmECNDetail"]');
        const modalContentTitle = modalContentFrame.locator('#ctl00_MainContent_lblTitle');
        await expect(modalContentTitle).toBeVisible({ timeout: 40000 });
        console.log('✅ モーダル内のコンテンツが表示されました。');

        console.log('ℹ️ 最終的な描画の安定を待つため3秒間待機します...');
        await page.waitForTimeout(3000);

        // --- スクリーンショット ---
        console.log('ℹ️ 詳細表示モーダルのスクリーンショットを撮影中...');
        const screenshotBuffer = await page.screenshot({ fullPage: true });
        await test.info().attach(`${user.id}-Detail-Modal.png`, {
          body: screenshotBuffer,
          contentType: 'image/png',
        });

        // --- クリーンアップ処理 ---
        console.log('ℹ️ クリーンアップ処理を開始します...');

        // 1. モーダルを閉じる
        console.log('-> 詳細モーダルを閉じます...');
        await mainFrame.locator('#rwmECNDetailCloseModalButton').click();
        
        // モーダルが閉じたことを確認
        await expect(page.locator('#ctl00_MainContent_rwmECNDetail_title')).toBeHidden();
        console.log('✅ モーダルを閉じました。');

        // 2. 検索結果ページから抜ける
        console.log('-> Exitリンクをクリックして検索ページから抜けます...');
        await mainFrame.locator('#hyxlnkExit').click();

        // 検索結果ページから抜けたことを確認
        await expect(resultsHierarchy).toBeHidden();
        console.log('✅ 検索ページから抜けました。');
        
        console.log('✅ クリーンアップが完了しました。');
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
