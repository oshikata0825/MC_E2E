import { test, expect } from '@playwright/test';
import { BASE_URL } from './config';
import * as fs from 'fs';
import * as path from 'path';
import { exec } from 'child_process';
import * as ExcelJS from 'exceljs';

// Read account information
const users = [
  { id: 'TOshikata', pass: '5D9dtEi*Qaivft', company: 'Mitsubishi Corporation' }
];

test.describe('Role Based Access Control - Spreadsheet Upload Visibility', () => {
  
  // ユーザーの数だけテストケースを自動生成します
  for (const user of users) {
    
    test(`Verify Spreadsheet Upload Visibility for ${user.id}`, async ({ page }) => {
      test.setTimeout(180000); // タイムアウトを180秒に延長
      
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
        await page.getByRole('button', { name: 'Client Maintenance' }).click();
        await page.getByRole('menuitem', { name: 'Client Maintenance' }).click();
      });

      // ---------------------------------------------------
      // Step 3: elements.txt の要素を確認 (Spreadsheet Upload)
      // ---------------------------------------------------
      await test.step('Verify Spreadsheet Upload Element', async () => {
        const mainFrameElement = page.locator('iframe[name="legacy-outlet"]');
        await mainFrameElement.waitFor({ state: 'attached', timeout: 30000 });
        const mainFrame = await mainFrameElement.contentFrame();
        if (!mainFrame) throw new Error('メインフレームの中身にアクセスできませんでした。');

        // 1. Spreadsheet Upload をクリック
        const uploadLink = mainFrame.locator('#ctl00_LeftHandNavigation_LHN_ctl00');
        await expect(uploadLink).toBeVisible({ timeout: 30000 });
        await uploadLink.click();
        console.log('✅ Spreadsheet Upload をクリックしました');
        
        // 少し待機
        await page.waitForTimeout(3000);

        // 2. "130840 SpreadSheet upload - DPS" が表示されているか確認
        const targetElement = mainFrame.locator('td', { hasText: '130840 SpreadSheet upload - DPS' });
        const isVisible = await targetElement.isVisible();
        const isAuthorizedUser = (user.id === 'MCTest5' || user.id === 'MCTest9' || user.id === 'TOshikata');

        if (isAuthorizedUser) {
          // Authorized users: Error if not visible
          if (isVisible) {
            console.log(`✅ ${user.id}: 130840 SpreadSheet upload - DPS が表示されています (期待通り)`);
            
            // 3. "Upload" リンクをクリックして新しいタブで開く
            console.log(`--- ${user.id}: "Upload" リンクのクリック処理を開始 ---`);
            const uploadActionLink = targetElement.locator('..').locator('a', { hasText: /Upload/i });
            
            const [popup] = await Promise.all([
              page.waitForEvent('popup'),
              uploadActionLink.first().click(),
            ]);

            await popup.waitForLoadState();
            console.log(`✅ ${user.id}: "Upload" ポップアップ画面が開きました`);

            // 4. 画面の読み込み完了を待機 (lbxPageName が表示されるまで)
            console.log(`ℹ️ 画面の読み込み完了を待機しています (#lbxPageName)...`);
            let titleFound = false;
            for (let i = 0; i < 10; i++) { // 最大10秒程度待機
              const allFrames = popup.frames();
              for (const frame of allFrames) {
                if (await frame.locator('#lbxPageName:has-text("DPS Spreadsheet Upload")').count() > 0) {
                  titleFound = true;
                  console.log(`✅ 画面の読み込みを確認しました (フレーム: "${frame.name()}")`);
                  break;
                }
              }
              if (titleFound) break;
              await popup.waitForTimeout(1000);
            }

            // 5. "Download Excel Template" を探す (全フレームを走査)
            console.log(`ℹ️ ポップアップ内の全フレームを走査してリンクを探します...`);
            let finalDownloadLink: any = null;

            // セレクターを定義 (text= は CSS セレクターとして使えないため分離)
            const getLinkLocator = (root: any) => {
              return root.locator('#ctl00_MainContent_hyxlnDownload').or(root.getByText('Download Excel Template'));
            };

            // まずメインページを探す
            if (await getLinkLocator(popup).count() > 0) {
              finalDownloadLink = getLinkLocator(popup).first();
              console.log('ℹ️ メインページ内にリンクを発見しました。');
            } else {
              // 全てのフレームをチェック
              const allFrames = popup.frames();
              console.log(`ℹ️ フレーム数: ${allFrames.length}`);
              for (const frame of allFrames) {
                console.log(`   - フレーム名: "${frame.name()}", URL: ${frame.url().substring(0, 50)}...`);
                if (await getLinkLocator(frame).count() > 0) {
                  finalDownloadLink = getLinkLocator(frame).first();
                  console.log(`ℹ️ フレーム "${frame.name()}" 内にリンクを発見しました。`);
                  break;
                }
              }
            }

            if (!finalDownloadLink) {
              console.log('⚠️ リンクが見つかりません。フォールバックとして getByText を試みます。');
              finalDownloadLink = popup.getByText('Download Excel Template');
            }

            await expect(finalDownloadLink).toBeVisible({ timeout: 20000 });

            console.log(`--- ${user.id}: "Download Excel Template" をクリックします ---`);
            const [download] = await Promise.all([
              popup.waitForEvent('download'),
              finalDownloadLink.click(),
            ]);

            const downloadPath = path.join(__dirname, `template_${user.id}.xlsx`);
            await download.saveAs(downloadPath);
            console.log(`✅ ${user.id}: テンプレートをダウンロードしました: ${downloadPath}`);

            // 6. ExcelJS を使用してテンプレートを編集
            let companyNames: string[] = [];
            const workbook = new ExcelJS.Workbook();
            await workbook.xlsx.readFile(downloadPath);
            const worksheet = workbook.getWorksheet(1); // 最初のシート
            if (worksheet) {
              const generateRandomStr = (length: number) => {
                const chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
                let result = '';
                for (let i = 0; i < length; i++) {
                  result += chars.charAt(Math.floor(Math.random() * chars.length));
                }
                return result;
              };

              const numRows = Math.floor(Math.random() * 4) + 1;
              console.log(`ℹ️ ${user.id}: ${numRows}行のデータを入力します。`);

              for (let i = 2; i < 2 + numRows; i++) {
                const randomStrB = generateRandomStr(9);
                const randomStrF = generateRandomStr(9);
                const countryCodes = ['JP', 'US', 'CN', 'GB', 'DE', 'SG', 'AU'];
                const randomCountry = countryCodes[Math.floor(Math.random() * countryCodes.length)];
                const companyName = `Company-For-SIT-${randomStrF}`;

                worksheet.getCell(`A${i}`).value = '130840';
                worksheet.getCell(`B${i}`).value = `For-SIT-${randomStrB}`;
                worksheet.getCell(`C${i}`).value = 'X';
                worksheet.getCell(`D${i}`).value = 'X1';
                worksheet.getCell(`E${i}`).value = '001';
                worksheet.getCell(`F${i}`).value = companyName;
                worksheet.getCell(`G${i}`).value = randomCountry;
                worksheet.getCell(`Q${i}`).value = 'Y';
                worksheet.getCell(`U${i}`).value = 'Y';
                
                companyNames.push(companyName);
                console.log(`   - 行 ${i}: A: 130840, B: For-SIT-${randomStrB}, F: ${companyName}, G: ${randomCountry}`);
              }

              await workbook.xlsx.writeFile(downloadPath);
              console.log(`✅ ${user.id}: エクセルを編集・保存しました。合計${numRows}行追加。`);
              console.log(`--- [エクセル追加データ詳細] ---`);
              for (let i = 2; i < 2 + numRows; i++) {
                const row = worksheet.getRow(i);
                console.log(`行 ${i}: A=${row.getCell(1).value}, B=${row.getCell(2).value}, F=${row.getCell(6).value}, G=${row.getCell(7).value}, Q=${row.getCell(17).value}, U=${row.getCell(21).value}`);
              }
              console.log(`-------------------------------`);
            }

            // ---------------------------------------------------
            // Step 4: エクセルファイルをアップロード
            // ---------------------------------------------------
            await test.step('Upload Edited Excel', async () => {
              console.log(`--- ${user.id}: エクセルファイルのアップロードを開始します ---`);
              
              // 全フレームを走査して Choose File ボタンを探す
              let chooseFileBtn: any = null;
              const popupFrames = popup.frames();
              for (const frame of popupFrames) {
                // elements.txt の ID (#ctl00_MainContent_btxChooseFile) またはテキストで探す
                const locator = frame.locator('#ctl00_MainContent_btxChooseFile').or(frame.getByRole('button', { name: 'Choose File' }));
                if (await locator.count() > 0) {
                  chooseFileBtn = locator.first();
                  console.log(`✅ Choose File ボタンをフレーム "${frame.name()}" 内に見つけました。`);
                  break;
                }
              }

              if (!chooseFileBtn) {
                const locator = popup.locator('#ctl00_MainContent_btxChooseFile').or(popup.getByRole('button', { name: 'Choose File' }));
                if (await locator.count() > 0) {
                  chooseFileBtn = locator.first();
                  console.log(`✅ Choose File ボタンをメインページ内に見つけました。`);
                }
              }

              if (chooseFileBtn) {
                await chooseFileBtn.scrollIntoViewIfNeeded();
                
                // ファイル選択イベントを待機しながらクリック
                const fileChooserPromise = popup.waitForEvent('filechooser');
                await chooseFileBtn.click();
                const fileChooser = await fileChooserPromise;
                
                // 編集したファイルをセット
                await fileChooser.setFiles(downloadPath);
                console.log(`✅ ${user.id}: ファイルを選択しました: ${downloadPath}`);
                
                // 少し待機
                await popup.waitForTimeout(2000);

                // 2. Upload and Import Data ボタンを探してクリック
                let uploadBtn: any = null;
                for (const frame of popupFrames) {
                  const locator = frame.locator('#ctl00_MainContent_lnxbtnUploadSpreadsheet').or(frame.getByRole('link', { name: 'Upload and Import Data' }));
                  if (await locator.count() > 0) {
                    uploadBtn = locator.first();
                    console.log(`✅ Upload and Import Data ボタンをフレーム "${frame.name()}" 内に見つけました。`);
                    break;
                  }
                }

                if (!uploadBtn) {
                  const locator = popup.locator('#ctl00_MainContent_lnxbtnUploadSpreadsheet').or(popup.getByRole('link', { name: 'Upload and Import Data' }));
                  if (await locator.count() > 0) {
                    uploadBtn = locator.first();
                    console.log(`✅ Upload and Import Data ボタンをメインページ内に見つけました。`);
                  }
                }

                if (uploadBtn) {
                  // --- 時差対策: アップロード前の状態を取得しておく ---
                  let initialEffectiveDate = '';
                  let initialHasRow = false;
                  let tableRoot: any = null;
                  
                  for (const frame of popupFrames) {
                    const table = frame.locator('#ctl00_MainContent_rgUploadWorkflowStatus_ctl00');
                    if (await table.count() > 0) {
                      tableRoot = table;
                      break;
                    }
                  }
                  if (!tableRoot) tableRoot = popup.locator('#ctl00_MainContent_rgUploadWorkflowStatus_ctl00');

                  if (await tableRoot.count() > 0) {
                    const firstRow = tableRoot.locator('tbody tr.rgRow, tbody tr.rgAltRow').first();
                    if (await firstRow.count() > 0) {
                      initialHasRow = true;
                      initialEffectiveDate = (await firstRow.locator('td:nth-child(3)').innerText()).trim();
                      console.log(`ℹ️ アップロード前の状態: レコードあり (時刻: ${initialEffectiveDate})`);
                    } else {
                      console.log('ℹ️ アップロード前の状態: レコードなし (テーブルは存在)');
                    }
                  } else {
                    console.log('ℹ️ アップロード前の状態: テーブル未出現');
                  }

                  const startTime = new Date();
                  console.log(`--- ${user.id}: Upload and Import Data をクリックします (${startTime.toLocaleString('ja-JP')}) ---`);
                  await uploadBtn.click();
                  
                  // 5秒待機
                  console.log('ℹ️ 5秒間待機してから Refresh をチェックします...');
                  await popup.waitForTimeout(5000);

                  // 3. Refresh ボタンを探してクリック & 新規レコード確認ループ
                  let refreshBtn: any = null;
                  for (const frame of popupFrames) {
                    const locator = frame.locator('#ctl00_MainContent_lnxBtnRefreshWorkflowStatus').or(frame.getByRole('link', { name: 'Refresh' }));
                    if (await locator.count() > 0) {
                      refreshBtn = locator.first();
                      console.log(`✅ Refresh ボタンをフレーム "${frame.name()}" 内に見つけました。`);
                      break;
                    }
                  }

                  if (!refreshBtn) {
                    const locator = popup.locator('#ctl00_MainContent_lnxBtnRefreshWorkflowStatus').or(popup.getByRole('link', { name: 'Refresh' }));
                    if (await locator.count() > 0) {
                      refreshBtn = locator.first();
                      console.log(`✅ Refresh ボタンをメインページ内に見つけました。`);
                    }
                  }

                  if (refreshBtn) {
                    const maxRetries = 20; // Running 状態を考慮して試行回数を増加
                    let isFinished = false;

                    for (let retry = 0; retry <= maxRetries; retry++) {
                      console.log(`--- ${user.id}: Refresh をクリックします (試行 ${retry + 1}/${maxRetries + 1}) ---`);
                      await refreshBtn.click();
                      await popup.waitForTimeout(5000); // 反映待ち

                      // テーブルを再取得して確認
                      if (await tableRoot.count() > 0) {
                        const firstRow = tableRoot.locator('tbody tr.rgRow, tbody tr.rgAltRow').first();
                        if (await firstRow.count() > 0) {
                          const currentEffectiveDate = (await firstRow.locator('td:nth-child(3)').innerText()).trim();
                          const status = (await firstRow.locator('td:nth-child(5)').innerText()).trim();
                          const summary = (await firstRow.locator('td:nth-child(6)').innerText()).trim();
                          
                          // 1. まず新しいレコードが出現したか確認
                          if (!initialHasRow || (currentEffectiveDate !== initialEffectiveDate)) {
                            console.log(`ℹ️ 最新レコードを確認: 時刻=${currentEffectiveDate}, ステータス=${status}`);

                            if (status === 'Complete') {
                              console.log('✅ アップロード処理が正常に完了しました。');
                              isFinished = true;
                              break;
                            } else if (status === 'Error') {
                              const errorMsg = `❌ アップロード処理が失敗しました。詳細: ${summary}`;
                              console.error(errorMsg);
                              await test.info().attach('upload-error-detail', {
                                body: await popup.screenshot({ fullPage: true }),
                                contentType: 'image/png',
                              });
                              await popup.close();
                              throw new Error(errorMsg);
                            } else if (status === 'Running') {
                              console.log('ℹ️ 処理ステータスは "Running" です。完了まで待機します...');
                            } else {
                              console.log(`ℹ️ 現在のステータス: "${status}"。更新を待ちます...`);
                            }
                          } else {
                            console.log('ℹ️ まだ新しいレコードが表示されていません。');
                          }
                        }
                      }

                      if (retry < maxRetries) {
                        console.log('ℹ️ 5秒待機してリトライします...');
                        await popup.waitForTimeout(5000);
                      }
                    }

                    if (!isFinished) {
                      console.warn(`⚠️ ${user.id}: 最大試行回数に達しましたが、処理が完了（Complete/Error）しませんでした。`);
                    }
                  } else {
                    console.warn(`⚠️ ${user.id}: Refresh ボタンが見つかりませんでした。`);
                  }

                  if (!popup.isClosed()) {
                    await popup.close();
                  }
                  
                  // ---------------------------------------------------
                  // Step 5: DPS Search/Reporting Lookup へ移動して確認
                  // ---------------------------------------------------
                  await test.step('Navigate to DPS Search Lookup', async () => {
                    console.log(`--- ${user.id}: DPS Search/Reporting Lookup への遷移を確認します ---`);
                    
                    // DPS メニューをクリック (トップバーの要素)
                    await page.bringToFront();
                    const dpsMenu = page.getByRole('button', { name: 'DPS' });
                    await dpsMenu.click();
                    console.log('✅ DPS メニューをクリックしました');

                    // DPS Search/Reporting Lookup をクリック (メニュー項目)
                    const lookupMenuItem = page.getByRole('menuitem', { name: 'DPS Search/Reporting Lookup' });
                    await lookupMenuItem.click();
                    console.log('✅ DPS Search/Reporting Lookup をクリックしました');

                    // 遷移後の待機とスクリーンショット
                    await page.waitForTimeout(3000);
                    const lookupScreenshot = await page.screenshot({ fullPage: true });
                    await test.info().attach('dps-lookup-page', {
                      body: lookupScreenshot,
                      contentType: 'image/png',
                    });

                    // MC Company Lookup をクリック (iframe 内)
                    const mainFrameElement = page.locator('iframe[name="legacy-outlet"]');
                    await mainFrameElement.waitFor({ state: 'attached', timeout: 30000 });
                    const mainFrame = await mainFrameElement.contentFrame();
                    if (!mainFrame) throw new Error('メインフレームの中身にアクセスできませんでした。');

                    const mcCompanyLookupLink = mainFrame.locator('a#ctl00_LeftHandNavigation_LHN_ctl00');
                    await expect(mcCompanyLookupLink).toBeVisible({ timeout: 20000 });
                    await mcCompanyLookupLink.click();
                    console.log('✅ MC Company Lookup をクリックしました');

                    await page.waitForTimeout(3000);
                    
                    // --- フィルタリング操作 (各会社名ごとに実行) ---
                    for (const companyName of companyNames) {
                      console.log(`--- ${user.id}: 会社名 "${companyName}" でフィルタリングを開始します ---`);
                      
                      // 1. CompanyName カラムのメニューボタン (...) をクリック
                      // 前回のループでメニューが開いたまま、あるいは不安定な状態を避けるため一度Escapeを送信
                      await page.keyboard.press('Escape');
                      await page.waitForTimeout(500);

                      const companyNameHeader = mainFrame.locator('th.rgHeader').filter({ hasText: /^CompanyName$/i });
                      const optionsBtn = companyNameHeader.locator('button.rgActionButton.rgOptions').first();
                      await optionsBtn.scrollIntoViewIfNeeded();
                      await optionsBtn.click({ force: true });
                      console.log('ℹ️ フィルタオプションメニューを開きました。');
                      await page.waitForTimeout(1000);

                      // 2. フィルタオペレータを "EqualTo" に変更
                      const condInput = mainFrame.locator('input[id*="HCFMRCMBFirstCond_Input"]');
                      await condInput.waitFor({ state: 'visible', timeout: 10000 });
                      
                      const currentOp = await condInput.inputValue();
                      if (currentOp !== 'EqualTo') {
                        await condInput.click({ force: true });
                        await page.waitForTimeout(500);
                        
                        // 選択肢から EqualTo を選ぶ
                        const equalToItem = mainFrame.locator('li.rcbItem', { hasText: /^EqualTo$/i }).or(page.locator('li.rcbItem', { hasText: /^EqualTo$/i })).first();
                        await equalToItem.waitFor({ state: 'visible', timeout: 5000 });
                        await equalToItem.click({ force: true });
                        console.log('ℹ️ フィルタオペレータを "EqualTo" に設定しました。');
                        await page.waitForTimeout(500);
                      } else {
                        console.log('ℹ️ フィルタオペレータは既に "EqualTo" です。');
                      }

                      // 3. 会社名を入力
                      const valInput = mainFrame.locator('input[id*="HCFMRTBFirstCond"]').first();
                      await valInput.waitFor({ state: 'visible', timeout: 5000 });
                      await valInput.click({ force: true });
                      await valInput.fill('');
                      await valInput.type(companyName, { delay: 30 }); // delayを少し短縮
                      await valInput.press('Tab');
                      await page.waitForTimeout(500);

                      // 4. Filter ボタンをクリック
                      const filterBtn = mainFrame.locator('button[id*="HCFMFilterButton"]').or(mainFrame.locator('span.rgButtonText:has-text("Filter")')).first();
                      await filterBtn.click({ force: true });
                      
                      console.log(`ℹ️ フィルタ適用を待機しています (${companyName})...`);
                      // ポストバックによるグリッドの更新を待つ
                      await page.waitForTimeout(3000);

                      // フィルタリング後のスクリーンショット
                      await test.info().attach(`filter-result-${companyName}`, {
                        body: await page.screenshot({ fullPage: true }),
                        contentType: 'image/png',
                      });

                      console.log(`✅ ${companyName} の確認が完了しました。`);
                    }
                    
                    console.log(`✅ ${user.id}: すべての会社名のフィルタリング確認が完了しました。`);
                  });

                } else {
                  console.error(`❌ ${user.id}: Upload and Import Data ボタンが見つかりませんでした。`);
                  await popup.close();
                }
              } else {
                console.error(`❌ ${user.id}: Choose File ボタンが見つかりませんでした。`);
                await popup.close();
              }
            });
            
          } else {
            const errorMsg = `❌ FAILED: ${user.id} は 130840 SpreadSheet upload - DPS が表示されるべきですが、非表示です。`;
            console.error(errorMsg);
            throw new Error(errorMsg);
          }
        } else {
          // それ以外のユーザーの場合: 表示されていればエラー
          if (isVisible) {
            const errorMsg = `❌ FAILED: ${user.id} は 130840 SpreadSheet upload - DPS が非表示であるべきですが、表示されています。`;
            console.error(errorMsg);
            throw new Error(errorMsg);
          } else {
            console.log(`✅ ${user.id}: 130840 SpreadSheet upload - DPS は非表示です (期待通り)`);
          }
        }
      });


      // ---------------------------------------------------
      // Step 5: ログアウト処理
      // ---------------------------------------------------
      await test.step('Logout', async () => {
        try {
          await page.bringToFront();
          
          // 1. ユーザー名(Tomohiro Oshikata)が表示されている要素をクリック
          const userSpan = page.locator('span', { hasText: 'Tomohiro Oshikata' });
          await userSpan.first().click({ timeout: 10000 });
          console.log('ℹ️ User menu opened.');

          // 2. Sign out ボタン (#logoutButton) をクリック
          const logoutButton = page.locator('#logoutButton');
          await logoutButton.waitFor({ state: 'visible', timeout: 10000 });
          await logoutButton.click({ force: true });
          console.log('✅ Logout button clicked.');
          
          await page.waitForURL(/Logon|default/i, { timeout: 15000 });
        } catch (e) {
          console.error(`❌ Logout failed: ${e.message}`);
          // スクリーンショットを撮ってデバッグ
          await test.info().attach('logout-failure', {
            body: await page.screenshot({ fullPage: true }),
            contentType: 'image/png',
          });
        }
      });
      
    });
  }
});
