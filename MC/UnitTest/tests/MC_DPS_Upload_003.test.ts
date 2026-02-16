import { test, expect } from '@playwright/test';
import { BASE_URL } from './config';
import * as fs from 'fs';
import * as path from 'path';
import * as ExcelJS from 'exceljs';

// Read account information from account.dat
const allUsers = JSON.parse(fs.readFileSync(path.join(__dirname, 'account.dat'), 'utf8'));
const users = allUsers.filter(u => u.id === 'MCTest5' || u.id === 'MCTest9');

test.describe('Spreadsheet Upload - Mandatory Fields Error Handling', () => {
  
  for (const user of users) {
    
    test(`Verify Mandatory Fields Error Handling for ${user.id}`, async ({ page }) => {
      // 9回のアップロードを行うため、余裕を持ってタイムアウトを設定
      test.setTimeout(1200000); 
      
      console.log(`--- テスト開始: ${user.id} ---`);

      // Step 1 & 2: Login and Navigate
      await page.goto(BASE_URL + '/Logon.aspx?ReturnUrl=%2fgtm%2fhome');
      await page.getByRole('textbox', { name: 'Company' }).fill(user.company);
      await page.getByRole('textbox', { name: 'Username' }).fill(user.id);
      await page.getByRole('textbox', { name: 'Password' }).fill(user.pass);
      await page.getByRole('link', { name: 'Login' }).click();

      await page.getByRole('button', { name: 'Client Maintenance' }).click();
      await page.getByRole('menuitem', { name: 'Client Maintenance' }).click();

      // Step 3: Main Logic
      const mainFrameElement = page.locator('iframe[name="legacy-outlet"]');
      await mainFrameElement.waitFor({ state: 'attached', timeout: 30000 });
      const mainFrame = await mainFrameElement.contentFrame();
      if (!mainFrame) throw new Error('メインフレームにアクセスできません。');

      const uploadLink = mainFrame.locator('#ctl00_LeftHandNavigation_LHN_ctl00');
      await expect(uploadLink).toBeVisible({ timeout: 30000 });
      await uploadLink.click();
      await page.waitForTimeout(3000);

      // Upload ポップアップを開く
      const uploadRow = mainFrame.locator('tr').filter({ has: mainFrame.locator('td', { hasText: '130840 SpreadSheet upload - DPS' }) }).first();
      const uploadActionLink = uploadRow.locator('a', { hasText: /Upload/i }).first();
      const popupPromise = page.waitForEvent('popup', { timeout: 60000 });
      await uploadActionLink.evaluate(el => (el as HTMLElement).click());
      const popup = await popupPromise;
      await popup.waitForLoadState();
      
      // ポップアップ内のタイトル確認
      let titleFound = false;
      for (let i = 0; i < 30; i++) {
        const allFrames = popup.frames();
        for (const frame of allFrames) {
          try {
            if (await frame.locator('#lbxPageName:has-text("DPS Spreadsheet Upload")').count() > 0) {
              titleFound = true;
              break;
            }
          } catch (e) { } 
        }
        if (titleFound) break;
        await popup.waitForTimeout(1000);
      }

      // テンプレートをダウンロード (初回のみ)
      let downloadLink: any = null;
      for (const frame of popup.frames()) {
        const loc = frame.locator('#ctl00_MainContent_hyxlnDownload').or(frame.getByText('Download Excel Template'));
        try { if (await loc.count() > 0) { downloadLink = loc.first(); break; } } catch (e) { } 
      }
      if (!downloadLink) downloadLink = popup.locator('#ctl00_MainContent_hyxlnDownload').or(popup.getByText('Download Excel Template')).first();
      const [download] = await Promise.all([popup.waitForEvent('download'), downloadLink.click()]);
      const templatePath = path.join(__dirname, `template_base_${user.id}.xlsx`);
      await download.saveAs(templatePath);

      // テスト項目リスト
      const testCols = [
        { col: 'A', name: 'ClientID', expected: 'Error' },
        { col: 'B', name: 'CompanyID', expected: 'Complete' }, 
        { col: 'C', name: 'ColC', expected: 'Error' },
        { col: 'D', name: 'ColD', expected: 'Error' },
        { col: 'E', name: 'ColE', expected: 'Error' },
        { col: 'F', name: 'CompanyName', expected: 'Error' },
        { col: 'G', name: 'Country', expected: 'Error' },
        { col: 'Q', name: 'DTSSearchFlag', expected: 'Error' },
        { col: 'U', name: 'ActiveFlag', expected: 'Error' },
      ];

      for (const target of testCols) {
        // ユニークなファイル名を生成 (秒単位のタイムスタンプ)
        const ts = new Date().toISOString().replace(/[:.-]/g, '').substring(0, 14);
        const fileName = `MT_ERR_${target.col}_${ts}.xlsx`;
        const testFilePath = path.join(__dirname, fileName);
        
        console.log(`\n--- テスト対象: ${target.name} (${target.col}列) を空にします ---`);
        console.log(`ℹ️ ファイル名: ${fileName}`);

        const wb = new ExcelJS.Workbook();
        await wb.xlsx.readFile(templatePath);
        const ws = wb.getWorksheet(1)!;
        
        // 2行目以降を完全にクリア
        ws.eachRow((row, rowNumber) => {
          if (rowNumber > 1) {
            row.eachCell((cell) => { cell.value = null; });
          }
        });
        while (ws.rowCount > 1) { ws.spliceRows(2, 1); } 
        
        // 2行目を新規作成して値をセット
        const row2 = ws.getRow(2);
        if (target.col !== 'A') row2.getCell(1).value = '130840';
        if (target.col !== 'B') row2.getCell(2).value = `MT-TEST-${target.col}-${ts}`;
        if (target.col !== 'C') row2.getCell(3).value = 'X';
        if (target.col !== 'D') row2.getCell(4).value = 'X1';
        if (target.col !== 'E') row2.getCell(5).value = '001';
        if (target.col !== 'F') row2.getCell(6).value = `Test-Company-${target.name}`;
        if (target.col !== 'G') row2.getCell(7).value = 'JP';
        row2.getCell(8).value = '1';
        if (target.col !== 'Q') row2.getCell(17).value = 'Y';
        if (target.col !== 'U') row2.getCell(21).value = 'Y';
        
        row2.commit();
        
        // デバッグログ: Q列とU列の状態をログ出力
        const qVal = ws.getCell('Q2').value;
        const uVal = ws.getCell('U2').value;
        const targetVal = ws.getCell(`${target.col}2`).value;
        console.log(`   [DEBUG] 書き込み直前: Q2="${qVal}", U2="${uVal}", Target(${target.col}2)="${targetVal}"`);

        await wb.xlsx.writeFile(testFilePath);

        // フレームとボタンの特定
        let tableFrame: any = null;
        for (const frame of popup.frames()) {
          if (await frame.locator('#ctl00_MainContent_rgUploadWorkflowStatus_ctl00').count() > 0) { tableFrame = frame; break; }
        }
        const chooseFileBtn = tableFrame.locator('#ctl00_MainContent_btxChooseFile').or(tableFrame.getByRole('button', { name: 'Choose File' })).first();
        const uploadBtn = tableFrame.locator('#ctl00_MainContent_lnxbtnUploadSpreadsheet').or(tableFrame.getByRole('link', { name: 'Upload and Import Data' })).first();
        const refreshBtn = tableFrame.locator('#ctl00_MainContent_lnxBtnRefreshWorkflowStatus').or(tableFrame.getByRole('link', { name: 'Refresh' })).first();

        // アップロード実行
        const fileChooserPromise = popup.waitForEvent('filechooser');
        await chooseFileBtn.click();
        const fileChooser = await fileChooserPromise;
        await fileChooser.setFiles(testFilePath);
        await uploadBtn.click();

        // ファイル名でレコードを特定してステータスを確認
        let finalStatus = '';
        let targetRow: any = null;

        for (let i = 0; i < 30; i++) {
          await popup.waitForTimeout(5000);
          await refreshBtn.click();
          await popup.waitForTimeout(2000); 
          
          const rows = tableFrame.locator('#ctl00_MainContent_rgUploadWorkflowStatus_ctl00 tbody tr');
          const count = await rows.count();
          for (let j = 0; j < count; j++) {
            const rowText = await rows.nth(j).innerText();
            if (rowText.includes(fileName)) {
              targetRow = rows.nth(j);
              finalStatus = (await targetRow.locator('td:nth-child(5)').innerText()).trim();
              break;
            }
          }

          if (targetRow) {
            console.log(`   - 試行 ${i+1}: レコード発見! ステータス = "${finalStatus}"`);
            if (finalStatus.includes('Error') || finalStatus.includes('Complete')) break;
          } else {
            console.log(`   - 試行 ${i+1}: ファイル名 "${fileName}" のレコードを待機中...`);
          }
        }

        // 判定と詳細確認
        if (finalStatus.includes(target.expected)) {
          console.log(`   ✅ OK: 期待通り "${target.expected}" を検知しました。`);
          if (target.expected === 'Error') {
            const errorCountLink = targetRow.locator('a[datafield="ErrorCount"]');
            if (await errorCountLink.count() > 0) {
              await errorCountLink.click();
              console.log(`   ℹ️ リンクをクリックしました。詳細の表示を10秒間待機します...`);
              await popup.waitForTimeout(10000); // 表示の安定待ちを10秒に延長
              
              await test.info().attach(`error-details-${target.name}`, { body: await popup.screenshot({ fullPage: true }), contentType: 'image/png' });
              console.log(`   📸 エラー詳細のスクリーンショットを撮影しました。`);
            }
          }
        } else {
          console.error(`   ❌ FAILED: ステータス不一致 (期待="${target.expected}", 実際="${finalStatus}")`);
        }

        // --- 徹底的なクリーンアップ (全フレーム/メインページから RadWindow とオーバーレイを除去) ---
        console.log(`   ℹ️ クリーンアップを実行します...`);
        const contexts = [popup, ...popup.frames()];
        for (const ctx of contexts) {
          try {
            if (ctx.isDetached()) continue;
            await ctx.evaluate(() => {
              // 1. Telerik API で閉じる
              try {
                // @ts-ignore
                const wm = (window as any).Telerik?.Web?.UI?.RadWindowUtils?.get_windowManager();
                if (wm) wm.get_windows().forEach((w: any) => { if (w.is_visible()) w.close(); });
              } catch (e) {}

              // 2. DOM から物理削除
              const selectors = [
                '[id*="RadWindowWrapper"]', 
                '[id*="RadModalBackground"]', 
                '.TelerikModalOverlay', 
                '.RadModalBackground', 
                '.rwModalMask', 
                '.rwModalBackground',
                'iframe[name="rwDetail"]'
              ];
              selectors.forEach(sel => {
                document.querySelectorAll(sel).forEach(el => {
                  console.log('Cleaning up:', sel);
                  el.remove();
                });
              });
            }).catch(() => {});
          } catch (e) {}
        }
        await popup.waitForTimeout(2000);
        
        expect.soft(finalStatus.includes(target.expected)).toBe(true);
        try { fs.unlinkSync(testFilePath); } catch (e) { } 
      }

      // Logout
      const userButtonName = `MC Test${user.id.substring(6)}`;
      await page.getByRole('button', { name: new RegExp(`^${userButtonName}`) }).click();
      await page.getByRole('button', { name: ' Sign out' }).click();
      
    });
  }
});
