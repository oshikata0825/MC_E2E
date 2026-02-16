import { test, expect } from '@playwright/test';
import { BASE_URL } from './config';
import * as fs from 'fs';
import * as path from 'path';
import * as ExcelJS from 'exceljs';

// Read account information from account.dat
const users = JSON.parse(fs.readFileSync(path.join(__dirname, 'account.dat'), 'utf8'));

test.describe('Role Based Access Control - Spreadsheet Upload Visibility', () => {
  
  for (const user of users) {
    
    test(`Verify Spreadsheet Upload Visibility for ${user.id}`, async ({ page }) => {
      test.setTimeout(300000); // タイムアウトを300秒に延長
      
      console.log(`--- テスト開始: ${user.id} ---`);

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
      // Step 2: メインコンテンツへ移動
      // ---------------------------------------------------
      await test.step('Navigate to Main Content', async () => {
        await page.getByRole('button', { name: 'Client Maintenance' }).click();
        await page.getByRole('menuitem', { name: 'Client Maintenance' }).click();
      });

      // ---------------------------------------------------
      // Step 3: メインロジック
      // ---------------------------------------------------
      await test.step('Spreadsheet Upload Test Flow', async () => {
        const mainFrameElement = page.locator('iframe[name="legacy-outlet"]');
        await mainFrameElement.waitFor({ state: 'attached', timeout: 30000 });
        const mainFrame = await mainFrameElement.contentFrame();
        if (!mainFrame) throw new Error('メインフレームの中身にアクセスできませんでした。');

        const uploadLink = mainFrame.locator('#ctl00_LeftHandNavigation_LHN_ctl00');
        await expect(uploadLink).toBeVisible({ timeout: 30000 });
        await uploadLink.click();
        await page.waitForTimeout(3000);

        const targetElement = mainFrame.locator('td', { hasText: '130840 SpreadSheet upload - DPS' });
        const isVisible = await targetElement.isVisible();
        const isAuthorizedUser = (user.id === 'MCTest5' || user.id === 'MCTest9');

        if (isVisible) {
          if (!isAuthorizedUser) throw new Error(`❌ FAILED: ${user.id} は非表示であるべきですが、表示されています。`);
          console.log(`✅ ${user.id}: Spreadsheet Upload 項目を確認。テストを開始します。`);

          // --- [1] テストデータの存在確認と準備 ---
          console.log('ℹ️ データ準備状況を確認するため DPS Search/Reporting Lookup へ移動します。');
          await page.bringToFront();
          await page.getByRole('button', { name: 'DPS' }).click();
          await page.getByRole('menuitem', { name: 'DPS Search/Reporting Lookup' }).click();
          await page.waitForTimeout(5000);

          let dpsFrame = await (page.locator('iframe[name="legacy-outlet"]')).contentFrame();
          if (!dpsFrame) throw new Error('DPSフレームが見つかりません');
          await dpsFrame.locator('a#ctl00_LeftHandNavigation_LHN_ctl00').click();
          await page.waitForTimeout(5000);

          const applyContainsFilter = async (frame: any, value: string) => {
            await page.keyboard.press('Escape');
            const header = frame.locator('th.rgHeader').filter({ hasText: /^CompanyName$/i });
            const optionsBtn = header.locator('button.rgActionButton.rgOptions').first();
            await optionsBtn.scrollIntoViewIfNeeded();
            await optionsBtn.click({ force: true });
            await page.waitForTimeout(3000);

            const cond = frame.locator('input[id*="HCFMRCMBFirstCond_Input"]');
            await cond.waitFor({ state: 'attached', timeout: 10000 });
            await cond.evaluate(el => (el as HTMLElement).click());
            await page.waitForTimeout(1000);

            for(let j=0; j<10; j++) await page.keyboard.press('ArrowUp');
            for(let i=0; i<15; i++){
              const val = await cond.inputValue();
              if(val && val.toLowerCase() === 'contains'){
                await page.keyboard.press('Enter');
                break;
              }
              await page.keyboard.press('ArrowDown');
              await page.waitForTimeout(200);
            }
            const valInput = frame.locator('input[id*="HCFMRTBFirstCond"]').first();
            await valInput.click({ force: true });
            await valInput.fill(value);
            await page.keyboard.press('Enter');
            await page.waitForTimeout(5000);
          };

          await applyContainsFilter(dpsFrame, 'Company-For-SIT');
          let rowCount = await dpsFrame.locator('tbody tr.rgRow, tbody tr.rgAltRow').count();

          if (rowCount === 0) {
            console.log('⚠️ テストデータが0件のため、セットアップ用のアップロードを行います。');
            // データ準備のためのアップロード (1件)
            await page.bringToFront();
            await page.getByRole('button', { name: 'Client Maintenance' }).click();
            await page.getByRole('menuitem', { name: 'Client Maintenance' }).click();
            const mfElement = page.locator('iframe[name="legacy-outlet"]');
            await mfElement.waitFor({ state: 'attached', timeout: 30000 });
            const mf = await mfElement.contentFrame();
            if (!mf) throw new Error('メインフレーム(mf)にアクセスできません。');
            
            const uploadLinkSetup = mf.locator('#ctl00_LeftHandNavigation_LHN_ctl00');
            await expect(uploadLinkSetup).toBeVisible({ timeout: 20000 });
            await uploadLinkSetup.click();
            
            const uploadRow = mf.locator('tr').filter({ has: mf.locator('td', { hasText: '130840 SpreadSheet upload - DPS' }) }).first();
            await expect(uploadRow).toBeVisible({ timeout: 20000 });
            await uploadRow.scrollIntoViewIfNeeded();

            const uploadActionLink = uploadRow.locator('a', { hasText: /Upload/i }).first();
            await expect(uploadActionLink).toBeVisible({ timeout: 10000 });
            console.log('ℹ️ (Setup) "Upload" リンクを JS でクリックします...');

            const popupPromise = page.waitForEvent('popup', { timeout: 60000 });
            await uploadActionLink.evaluate(el => (el as HTMLElement).click());
            const popup = await popupPromise;
            await popup.waitForLoadState();
            
            // 画面の読み込み完了を待機 (lbxPageName が表示されるまで)
            console.log(`ℹ️ (Setup) ポップアップの読み込み完了を待機しています...`);
            let titleFoundSetup = false;
            for (let i = 0; i < 30; i++) {
              const allFrames = popup.frames();
              for (const frame of allFrames) {
                try {
                  if (await frame.locator('#lbxPageName:has-text("DPS Spreadsheet Upload")').count() > 0) {
                    titleFoundSetup = true;
                    break;
                  }
                } catch (e) { }
              }
              if (titleFoundSetup) break;
              await popup.waitForTimeout(1000);
            }

            // ダウンロード (全フレーム走査)
            let downloadLink: any = null;
            for (const frame of popup.frames()) {
              const loc = frame.locator('#ctl00_MainContent_hyxlnDownload').or(frame.getByText('Download Excel Template'));
              try {
                if (await loc.count() > 0) {
                  downloadLink = loc.first();
                  break;
                }
              } catch (e) { }
            }
            if (!downloadLink) downloadLink = popup.locator('#ctl00_MainContent_hyxlnDownload').or(popup.getByText('Download Excel Template')).first();

            await expect(downloadLink).toBeVisible({ timeout: 20000 });
            const [download] = await Promise.all([popup.waitForEvent('download'), downloadLink.click()]);
            const setupPath = path.join(__dirname, `template_setup_${user.id}.xlsx`);
            if (fs.existsSync(setupPath)) { try { fs.unlinkSync(setupPath); } catch (e) { console.log(`ℹ️ 古いファイルの削除に失敗しました: ${setupPath}`); } }
            await download.saveAs(setupPath);

            const wb = new ExcelJS.Workbook();
            await wb.xlsx.readFile(setupPath);
            const ws = wb.getWorksheet(1);
            const suffix = Math.random().toString(36).substring(2, 7).toUpperCase();
            ws!.getCell('A2').value = '130840';
            ws!.getCell('B2').value = `For-SIT-SETUP-${suffix}`;
            ws!.getCell('C2').value = 'X';
            ws!.getCell('D2').value = 'X1';
            ws!.getCell('E2').value = '001';
            ws!.getCell('F2').value = `Company-For-SIT-SETUP-${suffix}`;
            ws!.getCell('G2').value = 'JP';
            ws!.getCell('Q2').value = 'Y';
            ws!.getCell('U2').value = 'Y';
            await wb.xlsx.writeFile(setupPath);

            // アップロード (全フレーム走査)
            let chooseFileBtn: any = null;
            for (const frame of popup.frames()) {
              const loc = frame.locator('#ctl00_MainContent_btxChooseFile').or(frame.getByRole('button', { name: 'Choose File' }));
              try {
                if (await loc.count() > 0) {
                  chooseFileBtn = loc.first();
                  break;
                }
              } catch (e) { }
            }
            if (!chooseFileBtn) chooseFileBtn = popup.locator('#ctl00_MainContent_btxChooseFile').or(popup.getByRole('button', { name: 'Choose File' })).first();

            const fileChooserPromise = popup.waitForEvent('filechooser');
            await chooseFileBtn.click();
            const fileChooser = await fileChooserPromise;
            await fileChooser.setFiles(setupPath);
            
            let uploadBtn: any = null;
            for (const frame of popup.frames()) {
              const loc = frame.locator('#ctl00_MainContent_lnxbtnUploadSpreadsheet').or(frame.getByRole('link', { name: 'Upload and Import Data' }));
              try {
                if (await loc.count() > 0) {
                  uploadBtn = loc.first();
                  break;
                }
              } catch (e) { }
            }
            if (!uploadBtn) uploadBtn = popup.locator('#ctl00_MainContent_lnxbtnUploadSpreadsheet').or(popup.getByRole('link', { name: 'Upload and Import Data' })).first();
            // --- 時差対策: アップロード前の最新レコードの時刻を取得 ---
            let initialEffectiveDateSetup = '';
            let initialHasRowSetup = false;
            let tableRootSetup: any = null;
            for (const frame of popup.frames()) {
              const table = frame.locator('#ctl00_MainContent_rgUploadWorkflowStatus_ctl00');
              if (await table.count() > 0) {
                tableRootSetup = table;
                break;
              }
            }
            if (tableRootSetup && await tableRootSetup.count() > 0) {
              const firstRow = tableRootSetup.locator('tbody tr.rgRow, tbody tr.rgAltRow').first();
              if (await firstRow.count() > 0) {
                initialHasRowSetup = true;
                initialEffectiveDateSetup = (await firstRow.locator('td:nth-child(3)').innerText()).trim();
                console.log(`ℹ️ (Setup) アップロード前の最新レコード時刻: ${initialEffectiveDateSetup}`);
              }
            }

            await uploadBtn.click();
            
            // Refresh ボタンの特定と完了待機
            let refreshBtn: any = null;
            for (const frame of popup.frames()) {
              const loc = frame.locator('#ctl00_MainContent_lnxBtnRefreshWorkflowStatus').or(frame.getByRole('link', { name: 'Refresh' }));
              try {
                if (await loc.count() > 0) {
                  refreshBtn = loc.first();
                  break;
                }
              } catch (e) { }
            }
            if (!refreshBtn) refreshBtn = popup.locator('#ctl00_MainContent_lnxBtnRefreshWorkflowStatus').or(popup.getByRole('link', { name: 'Refresh' })).first();

            for (let i = 0; i < 20; i++) {
              await popup.waitForTimeout(5000);
              await refreshBtn.click();
              
              let status = '';
              let currentEffectiveDate = '';
              let isNewRecord = false;

              for (const frame of popup.frames()) {
                const row = frame.locator('#ctl00_MainContent_rgUploadWorkflowStatus_ctl00 tbody tr.rgRow, #ctl00_MainContent_rgUploadWorkflowStatus_ctl00 tbody tr.rgAltRow').first();
                if (await row.count() > 0) {
                  currentEffectiveDate = (await row.locator('td:nth-child(3)').innerText()).trim();
                  status = (await row.locator('td:nth-child(5)').innerText()).trim();
                  
                  // 時刻が変わっているか、最初が空だった場合に「新規」と判定
                  if (!initialHasRowSetup || (currentEffectiveDate !== initialEffectiveDateSetup)) {
                    isNewRecord = true;
                  }
                  break;
                }
              }

              if (isNewRecord) {
                console.log(`   - 試行 ${i+1}: ステータス = ${status} (時刻: ${currentEffectiveDate})`);
                if (status === 'Complete') break;
                if (status === 'Error') {
                  const errorMsg = '❌ (Setup) アップロード処理が "Error" ステータスで終了しました。';
                  console.error(errorMsg);
                  await test.info().attach('setup-upload-error', { body: await popup.screenshot({ fullPage: true }), contentType: 'image/png' });
                  throw new Error(errorMsg);
                }
              } else {
                console.log(`   - 試行 ${i+1}: 新しいレコードを待機中...`);
              }
            }
            await popup.close();

            // 再度検索画面へ
            await page.bringToFront();
            await page.getByRole('button', { name: 'DPS' }).click();
            await page.getByRole('menuitem', { name: 'DPS Search/Reporting Lookup' }).click();
            dpsFrame = await (page.locator('iframe[name="legacy-outlet"]')).contentFrame();
            await dpsFrame!.locator('a#ctl00_LeftHandNavigation_LHN_ctl00').click();
            await applyContainsFilter(dpsFrame, 'Company-For-SIT');
            rowCount = await dpsFrame!.locator('tbody tr.rgRow, tbody tr.rgAltRow').count();
          }

          // --- [2] メインテスト: 全 CompanyID 記憶 & 一括アップロード (1回) ---
          const rows = dpsFrame!.locator('tbody tr.rgRow, tbody tr.rgAltRow');
          let allCompanyIds: string[] = [];
          for (let i = 0; i < rowCount; i++) {
            const id = await rows.nth(i).locator('td:nth-child(4)').innerText();
            allCompanyIds.push(id.trim());
          }

          // 5件以上の場合はランダムに5件選別
          let companyIds: string[] = [];
          if (allCompanyIds.length > 5) {
            console.log(`ℹ️ 合計 ${allCompanyIds.length} 件見つかりました。ランダムに 5 件を選択してテストを実施します。`);
            companyIds = allCompanyIds.sort(() => 0.5 - Math.random()).slice(0, 5);
          } else {
            companyIds = allCompanyIds;
          }
          console.log(`✅ ターゲット CompanyID 一覧: ${companyIds.join(', ')}`);

          // Spreadsheet Upload に戻る
          await page.bringToFront();
          await page.getByRole('button', { name: 'Client Maintenance' }).click();
          await page.getByRole('menuitem', { name: 'Client Maintenance' }).click();
          
          // iframe の出現と中身を待機
          const mf2Element = page.locator('iframe[name="legacy-outlet"]');
          await mf2Element.waitFor({ state: 'attached', timeout: 30000 });
          const mf2 = await mf2Element.contentFrame();
          if (!mf2) throw new Error('メインフレーム(mf2)にアクセスできません。');
          
          // Spreadsheet Upload をクリック
          const uploadLink2 = mf2.locator('#ctl00_LeftHandNavigation_LHN_ctl00');
          await expect(uploadLink2).toBeVisible({ timeout: 20000 });
          await uploadLink2.click();
          await page.waitForTimeout(3000);
          
          const uploadRow2 = mf2.locator('tr').filter({ has: mf2.locator('td', { hasText: '130840 SpreadSheet upload - DPS' }) }).first();
          await expect(uploadRow2).toBeVisible({ timeout: 20000 });
          await uploadRow2.scrollIntoViewIfNeeded();
          
          const uploadActionLink2 = uploadRow2.locator('a', { hasText: /Upload/i }).first();
          await expect(uploadActionLink2).toBeVisible({ timeout: 10000 });
          console.log('ℹ️ "Upload" リンクを JS でクリックします...');

          const popupPromise2 = page.waitForEvent('popup', { timeout: 60000 });
          await uploadActionLink2.evaluate(el => (el as HTMLElement).click());
          const popup2 = await popupPromise2;
          await popup2.waitForLoadState();

          // 画面の読み込み完了を待機 (lbxPageName が表示されるまで)
          console.log(`ℹ ポップアップの読み込み完了を待機しています...`);
          let titleFound2 = false;
          for (let i = 0; i < 30; i++) {
            const allFrames = popup2.frames();
            for (const frame of allFrames) {
              try {
                if (await frame.locator('#lbxPageName:has-text("DPS Spreadsheet Upload")').count() > 0) {
                  titleFound2 = true;
                  break;
                }
              } catch (e) { }
            }
            if (titleFound2) break;
            await popup2.waitForTimeout(1000);
          }

          // テンプレートを1回だけダウンロード (全フレーム走査)
          let downloadLink2: any = null;
          for (const frame of popup2.frames()) {
            const loc = frame.locator('#ctl00_MainContent_hyxlnDownload').or(frame.getByText('Download Excel Template'));
            try {
              if (await loc.count() > 0) {
                downloadLink2 = loc.first();
                break;
              }
            } catch (e) { }
          }
          if (!downloadLink2) downloadLink2 = popup2.locator('#ctl00_MainContent_hyxlnDownload').or(popup2.getByText('Download Excel Template')).first();

          await expect(downloadLink2).toBeVisible({ timeout: 20000 });
          const [download2] = await Promise.all([popup2.waitForEvent('download'), downloadLink2.click()]);
          const testPath = path.join(__dirname, `template_test_${user.id}.xlsx`);
          if (fs.existsSync(testPath)) { try { fs.unlinkSync(testPath); } catch (e) { console.log(`ℹ️ 古いファイルの削除に失敗しました: ${testPath}`); } }
          await download2.saveAs(testPath);

          // エクセルに全件記入
          const workbook = new ExcelJS.Workbook();
          await workbook.xlsx.readFile(testPath);
          const worksheet = workbook.getWorksheet(1);
          const genStr = (l: number) => Math.random().toString(36).substring(2, 2 + l).toUpperCase();
          
          // 検証用データの記憶
          const testDataMap = new Map<string, {
            clientId: string, // A
            companyId: string, // B
            colC: string, // C
            colD: string, // D
            colE: string, // E
            companyName: string, // F
            country: string, // G
            typeVal: string, // H
            companyNameJP: string, // I
            shortNameVal: string, // J
            addressLineVal: string, // K
            cityVal: string, // L
            stateVal: string, // M
            address2: string, // N
            address3: string, // O
            zipCode: string, // P
            qVal: string, // Q
            colR: string, // R
            colS: string, // S
            colT: string, // T
            uVal: string  // U
          }>();
          const countryCodes = ['JP', 'US', 'CN', 'GB', 'DE', 'SG', 'AU'];
          const stateCodes = ['8', 'AZ', 'WW', '65', '52', '51', '53', '54', '55', '56', '50', '79', '77', '78', '75', '74', '76'];
          
          // ランダムな都市名を生成する関数
          const genCityName = () => {
            const cities = ['London', 'New York', 'Tokyo', 'Paris', 'Berlin', 'Singapore', 'Sydney'];
            return `${cities[Math.floor(Math.random() * cities.length)]} ${genAlphaNum(3)}`;
          };

          // ランダムな英数字を生成する関数
          const genAlphaNum = (l: number) => {
            const chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789';
            let res = '';
            for(let i=0; i<l; i++) res += chars.charAt(Math.floor(Math.random() * chars.length));
            return res;
          };

          // ランダムな日本語（ひらがな）を生成する関数
          const genJpStr = (l: number) => {
            let res = '';
            for(let i=0; i<l; i++) res += String.fromCharCode(0x3041 + Math.floor(Math.random() * 80));
            return res;
          };

          for (let i = 0; i < companyIds.length; i++) {
            const r = i + 2;
            const companyId = companyIds[i];
            const clientId = '130840';
            const colC = 'X';
            const colD = 'X1';
            const colE = '001';
            const companyName = `Company-For-SIT-TEST-${genStr(5)}`;
            const randomCountry = countryCodes[Math.floor(Math.random() * countryCodes.length)];
            const typeVal = Math.random() > 0.5 ? '1' : '2';
            const companyNameJP = `Company-For-SIT-${genJpStr(6)}`;
            const shortNameVal = genAlphaNum(6);
            const addressLineVal = `${Math.floor(Math.random() * 999) + 1} Random Street ${genAlphaNum(4)}`;
            const cityVal = genCityName(); // Restore missing
            const stateVal = stateCodes[Math.floor(Math.random() * stateCodes.length)]; // Restore missing
            const zipCode = `${Math.floor(Math.random() * 900) + 100}-${Math.floor(Math.random() * 9000) + 1000}`; // N
            const subsidiaryVal = (Math.floor(Math.random() * 7) + 1).toString(); // O (1-7)
            const officeNameVal = genAlphaNum(10); // P (8+ chars)
            const qVal = Math.random() > 0.5 ? 'Y' : 'N';
            const colR = (Math.floor(Math.random() * 5) + 1).toString(); // R (1-5)
            const colT = Math.random() > 0.5 ? 'Y' : 'N'; // T (Y/N)
            const uVal = 'Y';

            worksheet!.getCell(`A${r}`).value = clientId;
            worksheet!.getCell(`B${r}`).value = companyId;
            worksheet!.getCell(`C${r}`).value = colC;
            worksheet!.getCell(`D${r}`).value = colD;
            worksheet!.getCell(`E${r}`).value = colE;
            worksheet!.getCell(`F${r}`).value = companyName;
            worksheet!.getCell(`G${r}`).value = randomCountry;
            worksheet!.getCell(`H${r}`).value = typeVal;
            worksheet!.getCell(`I${r}`).value = companyNameJP;
            worksheet!.getCell(`J${r}`).value = shortNameVal;
            worksheet!.getCell(`K${r}`).value = addressLineVal;
            worksheet!.getCell(`L${r}`).value = cityVal;
            worksheet!.getCell(`M${r}`).value = stateVal;
            worksheet!.getCell(`N${r}`).value = zipCode;
            worksheet!.getCell(`O${r}`).value = subsidiaryVal;
            worksheet!.getCell(`P${r}`).value = officeNameVal;
            worksheet!.getCell(`Q${r}`).value = qVal;
            worksheet!.getCell(`R${r}`).value = colR;
            // Column S (SanFlag) is not updated via upload per developer confirmation
            worksheet!.getCell(`T${r}`).value = colT;
            worksheet!.getCell(`U${r}`).value = uVal;

            testDataMap.set(companyId, { 
              clientId, companyId, colC, colD, colE, companyName, country: randomCountry, 
              typeVal, companyNameJP, shortNameVal, addressLineVal, cityVal, stateVal,
              zipCode, address2: subsidiaryVal, address3: officeNameVal, qVal, colR, colT, uVal 
            });          }
          await workbook.xlsx.writeFile(testPath);
          console.log(`✅ ${companyIds.length}件のデータをエクセルに記入し、1回アップロードします。`);

          // アップロード (1回)
          let chooseFileBtn2: any = null;
          for (const frame of popup2.frames()) {
            const loc = frame.locator('#ctl00_MainContent_btxChooseFile').or(frame.getByRole('button', { name: 'Choose File' }));
            try {
              if (await loc.count() > 0) {
                chooseFileBtn2 = loc.first();
                break;
              }
            } catch (e) { }
          }
          if (!chooseFileBtn2) chooseFileBtn2 = popup2.locator('#ctl00_MainContent_btxChooseFile').or(popup2.getByRole('button', { name: 'Choose File' })).first();

          const fcPromise = popup2.waitForEvent('filechooser');
          await chooseFileBtn2.click();
          const fc = await fcPromise;
          await fc.setFiles(testPath);
          
          let uploadBtn2: any = null;
          for (const frame of popup2.frames()) {
            const loc = frame.locator('#ctl00_MainContent_lnxbtnUploadSpreadsheet').or(frame.getByRole('link', { name: 'Upload and Import Data' }));
            try {
              if (await loc.count() > 0) {
                uploadBtn2 = loc.first();
                break;
              }
            } catch (e) { }
          }
          if (!uploadBtn2) uploadBtn2 = popup2.locator('#ctl00_MainContent_lnxbtnUploadSpreadsheet').or(popup2.getByRole('link', { name: 'Upload and Import Data' })).first();
          // --- 時差対策: アップロード前の最新レコードの時刻を取得 ---
          let initialEffectiveDate2 = '';
          let initialHasRow2 = false;
          let tableRoot2: any = null;
          for (const frame of popup2.frames()) {
            const table = frame.locator('#ctl00_MainContent_rgUploadWorkflowStatus_ctl00');
            if (await table.count() > 0) {
              tableRoot2 = table;
              break;
            }
          }
          if (tableRoot2 && await tableRoot2.count() > 0) {
            const firstRow = tableRoot2.locator('tbody tr.rgRow, tbody tr.rgAltRow').first();
            if (await firstRow.count() > 0) {
              initialHasRow2 = true;
              initialEffectiveDate2 = (await firstRow.locator('td:nth-child(3)').innerText()).trim();
              console.log(`ℹ️ アップロード前の最新レコード時刻: ${initialEffectiveDate2}`);
            }
          }

          await uploadBtn2.click();

          // Refresh ボタンの特定
          let refreshBtn2: any = null;
          for (const frame of popup2.frames()) {
            const loc = frame.locator('#ctl00_MainContent_lnxBtnRefreshWorkflowStatus').or(frame.getByRole('link', { name: 'Refresh' }));
            try {
              if (await loc.count() > 0) {
                refreshBtn2 = loc.first();
                break;
              }
            } catch (e) { }
          }
          if (!refreshBtn2) refreshBtn2 = popup2.locator('#ctl00_MainContent_lnxBtnRefreshWorkflowStatus').or(popup2.getByRole('link', { name: 'Refresh' })).first();

          for (let i = 0; i < 20; i++) {
            await popup2.waitForTimeout(5000);
            await refreshBtn2.click();
            
            let status = '';
            let currentEffectiveDate = '';
            let isNewRecord = false;

            for (const frame of popup2.frames()) {
              const row = frame.locator('#ctl00_MainContent_rgUploadWorkflowStatus_ctl00 tbody tr.rgRow, #ctl00_MainContent_rgUploadWorkflowStatus_ctl00 tbody tr.rgAltRow').first();
              if (await row.count() > 0) {
                currentEffectiveDate = (await row.locator('td:nth-child(3)').innerText()).trim();
                status = (await row.locator('td:nth-child(5)').innerText()).trim();
                
                if (!initialHasRow2 || (currentEffectiveDate !== initialEffectiveDate2)) {
                  isNewRecord = true;
                }
                break;
              }
            }

            if (isNewRecord) {
              console.log(`   - 試行 ${i+1}: ステータス = ${status} (時刻: ${currentEffectiveDate})`);
              if (status === 'Complete') break;
              if (status === 'Error') {
                const errorMsg = '❌ アップロード処理が "Error" ステータスで終了しました。';
                console.error(errorMsg);
                await test.info().attach('upload-error', { body: await popup2.screenshot({ fullPage: true }), contentType: 'image/png' });
                throw new Error(errorMsg);
              }
            } else {
              console.log(`   - 試行 ${i+1}: 新しいレコードを待機中...`);
            }
          }
          await popup2.close();

          // --- [3] 検証フェーズ: 各 ID を Edit で確認 ---
          console.log('ℹ️ 最終検証のため DPS Search/Reporting Lookup へ移動します。');
          await page.bringToFront();
          await page.getByRole('button', { name: 'DPS' }).click();
          await page.getByRole('menuitem', { name: 'DPS Search/Reporting Lookup' }).click();
          const finalFrame = await (page.locator('iframe[name="legacy-outlet"]')).contentFrame();
          await finalFrame!.locator('a#ctl00_LeftHandNavigation_LHN_ctl00').click();
          await page.waitForTimeout(5000);

          for (const companyId of companyIds) {
            console.log(`ℹ️ CompanyID: "${companyId}" を検証中...`);
            await page.keyboard.press('Escape');
            const h = finalFrame!.locator('th.rgHeader').filter({ hasText: /CompanyID/i });
            const optBtn = h.locator('button.rgActionButton.rgOptions').first();
            await optBtn.scrollIntoViewIfNeeded();
            await optBtn.click({ force: true });
            await page.waitForTimeout(2000);
            
            const cond = finalFrame!.locator('input[id*="HCFMRCMBFirstCond_Input"]');
            await cond.waitFor({ state: 'attached', timeout: 10000 });
            await cond.evaluate(el => (el as HTMLElement).click());
            await page.waitForTimeout(1000);

            for(let j=0; j<10; j++) await page.keyboard.press('ArrowUp');
            for(let i=0; i<15; i++){
              const val = await cond.inputValue();
              if(val && val.toLowerCase() === 'equalto') {
                await page.keyboard.press('Enter');
                break;
              }
              await page.keyboard.press('ArrowDown');
              await page.waitForTimeout(200);
            }
            const valInputFinal = finalFrame!.locator('input[id*="HCFMRTBFirstCond"]').first();
            await valInputFinal.click({ force: true });
            await valInputFinal.fill(companyId);
            await page.keyboard.press('Enter');
            
            // フィルタ適用（ポストバック）の完了を待機
            console.log(`   - フィルタ適用待ち...`);
            await page.waitForTimeout(5000);

            const editLink = finalFrame!.locator('tr.rgRow td, tr.rgAltRow td').locator('a', { hasText: /^Edit$/ }).first();
            await expect(editLink).toBeVisible({ timeout: 10000 });
            await editLink.scrollIntoViewIfNeeded();

            console.log(`   - Edit リンクをクリックします`);
            const [editPopup] = await Promise.all([
              page.waitForEvent('popup', { timeout: 30000 }),
              editLink.click({ force: true })
            ]);
            await editPopup.waitForLoadState();
            await editPopup.waitForTimeout(8000);

              // --- 値の検証 (最大10秒待機) ---
              const expected = testDataMap.get(companyId);
              if (expected) {
                console.log(`   - [検証] CompanyID: ${companyId}`);
                
                let actualCompID = '';
                let actualName = '';
                let actualCountry = '';
                let actualType = '';
                let actualNameJP = '';
                let actualShortName = '';
                let actualAddress1 = '';
                let actualSubsidiaryClass = '';
                let actualAddress2 = '';
                let actualAddress3 = '';
                let actualCity = '';
                let actualState = '';
                let actualZip = '';
                let actualQ = '';
                let actualR = '';
                let actualT = '';
                let actualU = '';

                for (let k = 0; k < 20; k++) {
                  for (const frame of editPopup.frames()) {
                    try {
                      if (!actualCompID) {
                        const loc = frame.locator('#lblCompanyID, #txtCompanyID');
                        if (await loc.count() > 0) actualCompID = (await loc.first().innerText() || await loc.first().inputValue()).trim();
                      }
                      if (!actualName) {
                        const nameLoc = frame.locator('input[id*="txtCompanyName"]');
                        if (await nameLoc.count() > 0) actualName = (await nameLoc.first().inputValue()).trim();
                      }
                      if (!actualCountry) {
                        const countryLoc = frame.locator('div[id*="drpCompanyCountryCode"], [id*="s2id_"][id*="drpCompanyCountryCode"]').locator('.select2-chosen');
                        if (await countryLoc.count() > 0) actualCountry = (await countryLoc.innerText()).trim();
                      }
                      if (!actualType) {
                        const typeLoc = frame.locator('div[id*="drpHQDivisionFlag"], [id*="s2id_"][id*="drpHQDivisionFlag"]').locator('.select2-chosen');
                        if (await typeLoc.count() > 0) actualType = (await typeLoc.innerText()).trim();
                      }
                      if (!actualNameJP) {
                        const nameJPLoc = frame.locator('#txtCompanyNameKanji');
                        if (await nameJPLoc.count() > 0) actualNameJP = (await nameJPLoc.first().inputValue()).trim();
                      }
                      if (!actualShortName) {
                        const shortNameLoc = frame.locator('#txtShortName');
                        if (await shortNameLoc.count() > 0) actualShortName = (await shortNameLoc.first().inputValue()).trim();
                      }
                      if (!actualAddress1) {
                        const addr1Loc = frame.locator('#txtCompanyAddress1');
                        if (await addr1Loc.count() > 0) actualAddress1 = (await addr1Loc.first().inputValue()).trim();
                      }
                      if (!actualSubsidiaryClass) {
                        const subLoc = frame.locator('div[id*="drpSubSidiaryClassification"], [id*="s2id_"][id*="drpSubSidiaryClassification"]').locator('.select2-chosen');
                        if (await subLoc.count() > 0) actualSubsidiaryClass = (await subLoc.innerText()).trim();
                      }
                      if (!actualAddress3) {
                        // P列: OfficeName
                        const officeLoc = frame.locator('#txtOfficeName');
                        if (await officeLoc.count() > 0) actualAddress3 = (await officeLoc.first().inputValue()).trim();
                      }
                      if (!actualCity) {
                        const cityLoc = frame.locator('#txtCompanyCity');
                        if (await cityLoc.count() > 0) actualCity = (await cityLoc.first().inputValue()).trim();
                      }
                      if (!actualState) {
                        const stateLoc = frame.locator('div[id*="drpCompanyState"], [id*="s2id_"][id*="drpCompanyState"]').locator('.select2-chosen');
                        if (await stateLoc.count() > 0) actualState = (await stateLoc.innerText()).trim();
                      }
                      if (!actualZip) {
                        const loc = frame.locator('#txtCompanyPostalCode');
                        if (await loc.count() > 0) actualZip = (await loc.first().inputValue()).trim();
                      }
                      if (!actualQ) {
                        const qLoc = frame.locator('div[id*="drpDTSSearchFlag"], [id*="s2id_"][id*="drpDTSSearchFlag"]').locator('.select2-chosen');
                        if (await qLoc.count() > 0) actualQ = (await qLoc.innerText()).trim();
                      }
                      if (!actualR) {
                        const rLoc = frame.locator('div[id*="drpInternalJudgement"], [id*="s2id_"][id*="drpInternalJudgement"]').locator('.select2-chosen');
                        if (await rLoc.count() > 0) actualR = (await rLoc.innerText()).trim();
                      }
                      if (!actualT) {
                        const tLoc = frame.locator('div[id*="drpCompanyWide"], [id*="s2id_"][id*="drpCompanyWide"]').locator('.select2-chosen');
                        if (await tLoc.count() > 0) actualT = (await tLoc.innerText()).trim();
                      }
                      if (!actualU) {
                        const uLoc = frame.locator('div[id*="drpActiveFlag"], [id*="s2id_"][id*="drpActiveFlag"]').locator('.select2-chosen');
                        if (await uLoc.count() > 0) actualU = (await uLoc.innerText()).trim();
                      }
                    } catch (e) { }
                  }
                  if (actualCompID && actualName && actualCountry && actualType && actualNameJP && actualShortName && actualAddress1 && actualSubsidiaryClass && actualCity && actualState && actualZip && actualQ && actualR && actualT && actualU) break;
                  await editPopup.waitForTimeout(500);
                }

                // --- 判定 (B列からU列まですべて) ---
                console.log(`     B列 (CompanyID): 期待值="${expected.companyId}", 実際="${actualCompID}"`);
                expect.soft(actualCompID).toBe(expected.companyId);

                console.log(`     F列 (Name): 期待值="${expected.companyName}", 実際="${actualName}"`);
                expect.soft(actualName).toBe(expected.companyName);

                console.log(`     G列 (Country): 期待值="${expected.country}", 実際="${actualCountry}"`);
                expect.soft(actualCountry.startsWith(expected.country)).toBe(true);

                const typeText = expected.typeVal === '1' ? '1-本社' : '2-事業所';
                console.log(`     H列 (HQDivisionFlag): 期待值="${typeText}", 実際="${actualType}"`);
                expect.soft(actualType.includes(typeText) || actualType.startsWith(expected.typeVal)).toBe(true);

                console.log(`     I列 (CompanyNameJP): 期待值="${expected.companyNameJP}", 実際="${actualNameJP}"`);
                expect.soft(actualNameJP).toBe(expected.companyNameJP);

                console.log(`     J列 (ShortName): 期待值="${expected.shortNameVal}", 実際="${actualShortName}"`);
                expect.soft(actualShortName).toBe(expected.shortNameVal);

                console.log(`     K列 (Address1): 期待值="${expected.addressLineVal}", 実際="${actualAddress1}"`);
                expect.soft(actualAddress1).toBe(expected.addressLineVal);

                // O列 (Subsidiary Classification) の日本語名称マッピング
                const subsidiaryMap: { [key: string]: string } = {
                  '1': 'MC単体',
                  '2': '現地法人',
                  '3': '子会社',
                  '4': '共同支配企業',
                  '5': '共同支配事業',
                  '6': '関連会社',
                  '7': 'その他'
                };
                const expectedSubsidiaryLabel = subsidiaryMap[expected.address2] || expected.address2;
                console.log(`     O列 (SubsidiaryClass): 期待值コード="${expected.address2}" (${expectedSubsidiaryLabel}), 実際="${actualSubsidiaryClass}"`);
                const isSubsidiaryMatch = actualSubsidiaryClass.includes(expectedSubsidiaryLabel) || actualSubsidiaryClass.startsWith(expected.address2);
                expect.soft(isSubsidiaryMatch, `O列(SubsidiaryClass)不一致: ID=${companyId}, 期待=${expectedSubsidiaryLabel}, 実際=${actualSubsidiaryClass}`).toBe(true);

                console.log(`     P列 (OfficeName): 期待值="${expected.address3}", 実際="${actualAddress3}"`);
                expect.soft(actualAddress3).toBe(expected.address3);

                console.log(`     N列 (ZipCode): 期待值="${expected.zipCode}", 実際="${actualZip}"`);
                expect.soft(actualZip).toBe(expected.zipCode);

                console.log(`     L列 (City): 期待值="${expected.cityVal}", 実際="${actualCity}"`);
                expect.soft(actualCity).toBe(expected.cityVal);

                console.log(`     M列 (State): 期待值コード="${expected.stateVal}", 実際="${actualState}"`);
                if (expected.stateVal) {
                  expect.soft(actualState.startsWith(expected.stateVal) || actualState.includes(`${expected.stateVal}-`)).toBe(true);
                }

                const qText = expected.qVal === 'Y' ? 'Yes' : 'No';
                console.log(`     Q列 (DTSSearchFlag): 期待值="${qText}", 実際="${actualQ}"`);
                expect.soft(actualQ).toBe(qText);

                // R列 (Internal Judgement) の名称マッピング
                const judgementMap: { [key: string]: string } = {
                  '1': '1-外国ユーザーリスト',
                  '2': '2-制裁対象者(禁止先を除く)',
                  '3': '3-禁止先(SDN等)',
                  '4': '4-Entity List等',
                  '5': '5-誤検知'
                };
                const expectedJudgementLabel = judgementMap[expected.colR] || expected.colR;
                console.log(`     R列 (InternalJudgement): 期待值="${expectedJudgementLabel}", 実際="${actualR}"`);
                expect.soft(actualR.includes(expectedJudgementLabel), `R列(InternalJudgement)不一致: ID=${companyId}, 期待=${expectedJudgementLabel}, 実際=${actualR}`).toBe(true);

                const tText = expected.colT === 'Y' ? 'Yes' : 'No';
                console.log(`     T列 (CompanyWide): 期待值="${tText}", 実際="${actualT}"`);
                expect.soft(actualT.includes(tText), `T列(CompanyWide)不一致: ID=${companyId}, 期待=${tText}, 実際=${actualT}`).toBe(true);

                const uText = expected.uVal === 'Y' ? 'はい' : 'いいえ';
                console.log(`     U列 (ActiveFlag): 期待值="${uText}", 実際="${actualU}"`);
                expect.soft(actualU.includes(uText), `U列(ActiveFlag)不一致: ID=${companyId}, 期待=${uText}, 実際=${actualU}`).toBe(true);

                // --- スクリーンショット撮影 (全項目が見えるようにスクロール) ---
                console.log(`   - スクリーンショットを撮影します...`);
                await editPopup.waitForTimeout(3000);
                // 基本情報タブの先頭へ
                await editPopup.evaluate(() => window.scrollTo(0, 0));
                await test.info().attach(`verify-${companyId}-top`, { 
                  body: await editPopup.screenshot({ fullPage: true }), 
                  contentType: 'image/png' 
                });

                // 下部が見えるようにスクロール (特にシステム設定などのドロップダウン)
                const lastElement = editPopup.locator('div[id*="pnlSystem"], div[id*="pnlAudit"]').last();
                if (await lastElement.count() > 0) {
                  await lastElement.scrollIntoViewIfNeeded();
                  await test.info().attach(`verify-${companyId}-bottom`, { 
                    body: await editPopup.screenshot(), 
                    contentType: 'image/png' 
                  });
                }
              }

            await editPopup.close();
            console.log(`   ✅ "${companyId}" の検証が完了しました。`);
          }

        } else if (isAuthorizedUser) {
          throw new Error(`❌ FAILED: ${user.id} は表示されるべきです。`);
        } else {
          console.log(`✅ ${user.id}: Spreadsheet Upload は非表示です (期待通り)`);
        }
      });

      // Step 4: Logout
      await test.step('Logout', async () => {
        const userButtonName = `MC Test${user.id.substring(6)}`;
        await page.getByRole('button', { name: new RegExp(`^${userButtonName}`) }).click();
        await page.getByRole('button', { name: ' Sign out' }).click();
        await page.waitForURL(/Logon|default/i);
      });
      
    });
  }
});
