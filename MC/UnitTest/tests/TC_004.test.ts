import { test, expect } from '@playwright/test';
import { BASE_URL } from './config';
import * as fs from 'fs';
import * as path from 'path';

// ベースのユーザー定義はそのまま流用
const users = JSON.parse(fs.readFileSync(path.join(__dirname, 'account.dat'), 'utf8'));

// テスト用のキーワードリスト
const keywordList = ['security', 'systems', 'equipment', 'components', 'telecommunications', 'laser', 'engine', 'aircraft'];

test.describe('Role Based Access Control - Keyword Search and Tab Verification', () => {
  
  for (const user of users) {
    
    test(`Verify Keyword Search and Tabs for ${user.id}`, async ({ page }) => {
      test.setTimeout(60000); // タイムアウトを60秒に延長
      
      console.log(`--- キーワード検索テスト開始: ${user.id} ---`);

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
      // Step 2: メインコンテンツへ移動 (変更なし)
      // ---------------------------------------------------
      await test.step('Navigate to Main Content', async () => {
        await page.getByRole('button', { name: 'Content' }).click();
        await page.getByRole('menuitem', { name: 'ECN/Dual Use List (Quick' }).click();
        await page.goto(BASE_URL + '/gtm/aspx?href=%2FContent%2FfugECCNDetail.aspx');
      });

      // ---------------------------------------------------
      // Step 3: "View As" 等のオーバーレイ対策 (変更なし)
      // ---------------------------------------------------
      await test.step('Handle Overlays', async () => {
        const viewAsOverlay = page.locator('text=Select group(s) to view as');
        if (await viewAsOverlay.count() > 0 && await viewAsOverlay.isVisible()) {
            console.log('⚠️ "View As" パネルを検出。必要に応じて閉じる処理を行ってください。');
        }
      });

      // ---------------------------------------------------
      // Step 4: 検索ボタンの確認とキーワード検索実行
      // ---------------------------------------------------
      await test.step('Verify and Perform Keyword Search', async () => {
        const mainFrameElement = page.locator('iframe[name="legacy-outlet"]');
        await mainFrameElement.waitFor({ state: 'attached', timeout: 30000 });
        const mainFrame = await mainFrameElement.contentFrame();
        if (!mainFrame) throw new Error('メインフレームの中身にアクセスできませんでした。');

        const statusBar = mainFrame.locator('#ctl00_MainContent_pnlStatusBarItems');
        const searchBtn = statusBar.locator('a[href*="lnxbtnStatusBarSearch"]');

        await expect(statusBar).toBeVisible({ timeout: 60000 });

        if (await searchBtn.isVisible()) {
          console.log(`✅ ${user.id}: Searchボタンを検出。キーワード検索をテストします。`);
          
          await mainFrame.locator('#ctl00_MainContent_cmbxRegulationList_Arrow').click();
          await mainFrame.getByText('US - UNITED STATES OF AMERICA - EAR/Commerce Control List').click();
          
          const randomKeyword = keywordList[Math.floor(Math.random() * keywordList.length)];
          console.log(`ℹ️ 使用するキーワード: ${randomKeyword}`);
          const searchTextbox = mainFrame.getByRole('textbox', { name: 'Enter an ECN Number or' });
          
          // --- オートコンプリート候補は iframe の中に表示される ---
          const dropdownContainer = mainFrame.locator('#ctl00_MainContent_cmbxStatusBarECN_DropDown');

          // ネットワーク応答を待つ準備 (これが最も堅牢な方法)
          const responsePromise = page.waitForResponse(
            resp => resp.url().includes('fugECCNDetail.aspx') && resp.status() === 200,
            { timeout: 15000 } // ネットワーク要求のタイムアウトを15秒に設定
          );

          await searchTextbox.click();
          await searchTextbox.type(randomKeyword, { delay: 100 });
          
          try {
            console.log('ℹ️ オートコンプリートのネットワーク応答を待っています...');
            await responsePromise;
            console.log('✅ ネットワーク応答を受信しました。');
            
            // ネットワーク応答後、UIが更新されるのを待つ (タイムアウトを15秒に延長)
            await expect(dropdownContainer).toBeVisible({ timeout: 15000 });
            
            // ドロップダウンコンテナ内の候補で、"ECN Quick Search Keywords" を含まないものを特定する
            // キーワード自体は結果のテキストに含まれない場合があるため、キーワードでの絞り込みは削除
            const suggestionLocator = dropdownContainer.locator(`li.rcbItem:not(:has-text("ECN Quick Search Keywords"))`);

            console.log('ℹ️ リトライループ開始前に5秒間待機します...');
            await page.waitForTimeout(5000);

            let count = 0;
            const maxRetries = 5;
            for (let i = 0; i < maxRetries; i++) {
              console.log(`ℹ️ クリック可能要素の待機中... (${i + 1}/${maxRetries})`);
              count = await suggestionLocator.count();
              if (count > 0) {
                break; // 候補が見つかったらループを抜ける
              }

              // デバッグ情報の出力
              console.log('⚠️ 候補が見つかりませんでした。コンテナのHTMLを確認します。');
              try {
                const containerHTML = await dropdownContainer.innerHTML();
                console.log('--- dropDownContainer HTML ---');
                console.log(containerHTML);
                console.log('------------------------------');
              } catch (htmlError) {
                console.log('⚠️ コンテナのHTML取得中にエラーが発生しました。', htmlError);
              }

              console.log('ℹ️ 3秒待機してリトライします。');
              await page.waitForTimeout(3000); // 3秒待ってからリトライ
            }

            console.log(`ℹ️ ${count}件のクリック可能な候補が見つかりました。`);

            if (count > 0) {
              // 安定化のため、ランダム性を排除し、常に2番目の候補(index: 1)を選択する
              const randomIndex = count > 1 ? 1 : 0;

              const chosenSuggestion = suggestionLocator.nth(randomIndex);
              
              const suggestionText = await chosenSuggestion.textContent();
              if (!suggestionText) throw new Error('候補のテキストが空でした。');
              
              // 候補のテキストと実際の結果でECCNのフォーマットが違うため、'.'以降を削除して比較する
              const eccnFromResult = suggestionText.trim().split(' ')[0].split('.')[0];
              console.log(`ℹ️ ランダムに選択した候補: "${suggestionText.trim()}"。ECCN "${eccnFromResult}" の結果を待ちます。`);
              
              // Playwrightのアクション可能性チェックに任せるため、forceオプションを削除
              console.log(`ℹ️ ${randomIndex + 1}番目の候補をクリックします。`);
              await chosenSuggestion.click();

              await expect(suggestionLocator.first()).not.toBeVisible({ timeout: 10000 });
              console.log('ℹ️ オートコンプリート候補が消え、検索が開始されました。');

              // ユーザー提供のセレクターに基づき、検索結果の要素を特定
              const resultLocator = mainFrame.locator('#ctl00_MainContent_lblShortDescription');

              // その要素が表示され、かつ、期待したECCNのテキストを含んでいることを確認する
              await expect(resultLocator).toBeVisible({ timeout: 30000 });
              await expect(resultLocator).toContainText(eccnFromResult, { timeout: 5000 });

              const headerText = await resultLocator.textContent();
              console.log(`ℹ️ 確認した結果ヘッダー: ${headerText?.trim()}`);

              // ユーザーが結果を目視確認するための3秒間の停止
              await page.waitForTimeout(3000);

              await test.step('Verify Tabs Content', async () => {
                const tabsToVerify = [
                  { name: /Header/, content: '#ctl00_MainContent_rgdMain_ctl00_ctl06_dtvwHeader' },
                  { name: /Control Reasons/, content: '#ctl00_MainContent_rgdMain_ctl00_ctl06_PageView2' },
                  { name: /Available Licenses\/Exceptions/, content: '#ctl00_MainContent_rgdMain_ctl00_ctl06_PageView3' }, // 検証文字列を安定的なものに修正
                  { name: /Notes/, content: '#ctl00_MainContent_rgdMain_ctl00_ctl06_PageView4' },
                  { name: /Related Definitions/, content: '#ctl00_MainContent_rgdMain_ctl00_ctl06_PageView5' },
                  { name: /HS Related Controls/, content: '#ctl00_MainContent_rgdMain_ctl00_ctl06_pnlRelatedControls' },
                  { name: /Export Related Controls/, content: '#ctl00_MainContent_rgdMain_ctl00_ctl06_PageView7' },
                  { name: /Documents/, content: '#ctl00_MainContent_rgdMain_ctl00_ctl06_rgdMainDocuments' }
                ];

                const failedTabs: string[] = [];

                for (const tab of tabsToVerify) {
                  const tabName = tab.name.source.replace(/[^a-zA-Z0-9]/g, ''); // ファイル名用にサニタイズ
                  try {
                    console.log(`ℹ️ タブ "${tab.name.source}" を検証中...`);
                    await mainFrame.getByRole('link', { name: tab.name }).click();
                    const contentLocator = mainFrame.locator(tab.content);
                    await expect(contentLocator).toBeVisible({ timeout: 15000 });
                    console.log(`✅ タブ "${tab.name.source}" のコンテンツが表示されました。`);

                    // 各タブの検証直後にスクリーンショットを撮影
                    console.log(`ℹ️ タブ "${tab.name.source}" のスクリーンショットを撮影します。`);
                    const screenshotBuffer = await page.screenshot();
                    await test.info().attach(`${user.id}-tab-${tabName}.png`, {
                      body: screenshotBuffer,
                      contentType: 'image/png',
                    });

                  } catch (error) {
                    console.error(`❌ タブ "${tab.name.source}" の検証中にエラーが発生しました。`);
                    failedTabs.push(tab.name.source);
                  } finally {
                    await page.waitForTimeout(500); // 次のタブへの切り替えのための短い待機
                  }
                }

                if (failedTabs.length > 0) {
                  throw new Error(`以下のタブの検証に失敗しました: ${failedTabs.join(', ')}`);
                }
              });

              const exitButton = mainFrame.getByRole('link', { name: 'Exit' });
              await exitButton.click(); 
              await page.waitForLoadState('domcontentloaded');

            } else {
              throw new Error(`リトライ(${maxRetries}回)後もオートコンプリート候補が見つかりませんでした。`);
            }
          } catch (e) {
            console.log(`ℹ️ "${randomKeyword}" のオートコンプリート候補の処理中にエラーが発生しました。`);
            console.error(e);
            // エラーを再スローして、テストを失敗させる
            throw e;
          }
        } else {
          console.log(`⏭️ ${user.id}: Searchボタンが見つかりません。このロールのテストはスキップします。`);
        }
      });

      // ---------------------------------------------------
      // Step 5: ログアウト処理 (変更なし)
      // ---------------------------------------------------
      await test.step('Logout', async () => {
        const userButtonName = `MC Test${user.id.substring(6)}`;
        await page.getByRole('button', { name: new RegExp(`^${userButtonName}`) }).click();
        await page.getByRole('button', { name: ' Sign out' }).click();
        await expect(page.getByRole('textbox', { name: 'Company' })).toBeVisible({ timeout: 30000 });
      });
      
    });
  }
});

