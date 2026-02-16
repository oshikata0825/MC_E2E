import { test, expect } from '@playwright/test';
import { BASE_URL } from './config';
import * as fs from 'fs';
import * as path from 'path';

// ベースのユーザー定義はそのまま流用
const users = JSON.parse(fs.readFileSync(path.join(__dirname, 'account.dat'), 'utf8'));

// 検索用のキーワードリスト
const keywordList = ['system', 'sensor', 'engine', 'laser', 'telecommunications', 'security', 'equipment'];

test.describe('Role Based Access Control - Keyword Search Result Expansion', () => {
  
  for (const user of users) {
    
    test(`Verify search result nodes are expandable for ${user.id}`, async ({ page }) => {
      test.setTimeout(180000); // タイムアウトを180秒に設定

      console.log(`--- Keyword Search Expansion Test Start: ${user.id} ---`);

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
      // Step 2: キーワード検索を実行し、ノードを展開する
      // ---------------------------------------------------
      await test.step('Perform Keyword Search and Expand Nodes', async () => {
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
            console.log('ℹ️ 最初の候補は "Quick Search Keywords" です。2番目の候補をクリックします。');
            expect(count, '最初の候補が"Quick Search"ですが、他に選択肢がありません。').toBeGreaterThan(1);
            await suggestions.nth(1).click();
        } else {
            console.log('ℹ️ 最初の候補が有効なため、クリックします。');
            await suggestions.first().click();
        }

        // 検索結果ツリーの表示を待機
        const resultsHierarchy = mainFrame.locator('#ctl00_MainContent_rtvECNHierarchy');
        await expect(resultsHierarchy).toBeVisible({ timeout: 30000 });
        await page.waitForTimeout(3000); // 描画が安定するまで待機
        console.log('✅ 検索結果が読み込まれました。');

        // --- ノードを展開する ---
        console.log('ℹ️ 検索結果のノードを展開します...');
        const plusIcons = mainFrame.locator('span.rtPlus');
        const initialIconCount = await plusIcons.count();
        const clickCount = Math.min(initialIconCount, 3);
        
        if (clickCount > 0) {
            console.log(`-> ${clickCount}個のノードを展開します。`);
            for (let i = 0; i < clickCount; i++) {
                // DOMが変更されるため、ループごとに最初の'+'アイコンを再取得してクリック
                await plusIcons.first().click();
                console.log(`  -> ${i + 1}個目のノードを展開しました。`);
                await page.waitForTimeout(1000); // アニメーション待機
            }
        } else {
            console.log('ℹ️ 展開可能なノードが見つかりませんでした。');
        }

        // 最後の展開処理の描画が完了するのを待つ
        console.log('ℹ️ 最終的な描画の安定を待つため3秒間待機します...');
        await page.waitForTimeout(3000);

        // --- Screenshot ---
        console.log('ℹ️ 展開後のスクリーンショットを撮影中...');
        const screenshotBuffer = await page.screenshot({ fullPage: true });
        await test.info().attach(`${user.id}-Expanded-Keyword-Search-Result.png`, {
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
