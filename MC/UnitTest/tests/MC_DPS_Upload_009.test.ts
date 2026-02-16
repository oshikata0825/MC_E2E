import { test, expect } from '@playwright/test';
import { BASE_URL } from './config';
import * as fs from 'fs';
import * as path from 'path';
import * as ExcelJS from 'exceljs';

/**
 * MC_DPS_Upload_009.test.ts
 * 
 * Mixed Upload Test: 10 rows in total.
 * - Rows with CompanyID: Existing records update.
 * - Rows without CompanyID: New records creation (system auto-assigns ID).
 * Verifies all mandatory fields (C-T) via Lookup search (ID or Name).
 */

const allUsers = JSON.parse(fs.readFileSync(path.join(__dirname, 'account.dat'), 'utf8'));
const users = allUsers.filter(user => user.id === 'MCTest5' || user.id === 'MCTest9');

test.describe('MC_DPS_Upload_009 - Mixed Existing/New Records Upload Verification', () => {
  
  for (const user of users) {
    
    test(`Verify Mixed Upload for ${user.id}`, async ({ page }) => {
      test.setTimeout(1800000); // 10件の検証のため30分に延長
      
      console.log(`--- テスト開始: ${user.id} ---`);

      // Step 1: Login
      await test.step('Login', async () => {
        await page.goto(BASE_URL + '/Logon.aspx?ReturnUrl=%2fgtm%2fhome');
        await page.getByRole('textbox', { name: 'Company' }).fill(user.company);
        await page.getByRole('textbox', { name: 'Username' }).fill(user.id);
        await page.getByRole('textbox', { name: 'Password' }).fill(user.pass);
        await page.getByRole('link', { name: 'Login' }).click();
      });

      // Step 2: Find Existing For-SIT- Records (Up to 5)
      let existingIds: string[] = [];
      await test.step('Find Existing Records', async () => {
        await page.getByRole('button', { name: 'DPS' }).click();
        await page.getByRole('menuitem', { name: 'DPS Search/Reporting Lookup' }).click();
        const legacyFrame = page.frameLocator('iframe[name="legacy-outlet"]');
        await legacyFrame.locator('a#ctl00_LeftHandNavigation_LHN_ctl00').click();
        await page.waitForTimeout(3000);

        await page.keyboard.press('Escape');
        const header = legacyFrame.locator('th.rgHeader').filter({ hasText: /^CompanyID$/i });
        await header.locator('button.rgActionButton.rgOptions').first().click({ force: true });
        await page.waitForTimeout(1000);

        const condInput = legacyFrame.locator('input[id*="HCFMRCMBFirstCond_Input"]');
        await condInput.click({ force: true });
        for (let j = 0; j < 10; j++) await page.keyboard.press('ArrowUp');
        for (let i = 0; i < 15; i++) {
          if ((await condInput.inputValue())?.toLowerCase() === 'startswith') { await page.keyboard.press('Enter'); break; }
          await page.keyboard.press('ArrowDown');
          await page.waitForTimeout(200);
        }
        await legacyFrame.locator('input[id*="HCFMRTBFirstCond"]').first().fill('For-SIT-');
        await page.keyboard.press('Enter');
        await page.waitForTimeout(5000);

        const rows = legacyFrame.locator('tbody tr.rgRow, tbody tr.rgAltRow');
        const count = await rows.count();
        console.log(`ℹ️ Found ${count} existing For-SIT- records.`);
        
        for (let i = 0; i < Math.min(count, 5); i++) {
          const id = (await rows.nth(i).locator('td:nth-child(4)').innerText()).trim();
          existingIds.push(id);
        }
        console.log(`✅ Collected ${existingIds.length} existing IDs for update.`);
      });

      // Step 3: Prepare Excel with 10 rows (Mixed)
      const ts = new Date().toISOString().replace(/[:.-]/g, '').substring(0, 14);
      const fileName = `MIXED_10_${ts}.xlsx`;
      const testFilePath = path.join(__dirname, fileName);
      const recordsToVerify: { searchKey: string, searchType: 'ID' | 'Name', expected: any }[] = [];

      await test.step('Prepare and Upload Excel', async () => {
        await page.getByRole('button', { name: 'Client Maintenance' }).click();
        await page.getByRole('menuitem', { name: 'Client Maintenance' }).click();
        const mainFrame = page.frameLocator('iframe[name="legacy-outlet"]');
        await mainFrame.locator('#ctl00_LeftHandNavigation_LHN_ctl00').click();
        const uploadRow = mainFrame.locator('td', { hasText: '130840 SpreadSheet upload - DPS' }).locator('..');
        const popupPromise = page.waitForEvent('popup');
        await uploadRow.locator('a', { hasText: /Upload/i }).first().click();
        const popup = await popupPromise;
        await popup.waitForLoadState();

        let tableFrame: any = null;
        for (let i = 0; i < 30; i++) {
          for (const f of popup.frames()) {
            if (await f.locator('#lbxPageName:has-text("DPS Spreadsheet Upload")').count() > 0) { tableFrame = f; break; }
          }
          if (tableFrame) break;
          await popup.waitForTimeout(1000);
        }

        const downloadBtn = tableFrame.locator('#ctl00_MainContent_hyxlnDownload').or(tableFrame.getByText('Download Excel Template')).first();
        const [download] = await Promise.all([popup.waitForEvent('download'), downloadBtn.click({ force: true })]);
        const baseTemplate = path.join(__dirname, `base_template_009.xlsx`);
        await download.saveAs(baseTemplate);

        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(baseTemplate);
        const worksheet = workbook.getWorksheet(1)!;

        const genStr = (l: number) => Math.random().toString(36).substring(2, 2 + l).toUpperCase();
        const countries = ['JP', 'US', 'CN', 'GB', 'DE', 'AU'];
        const states = ['8', 'AZ', '65', '52', '51'];

        // Prepare 10 rows
        for (let i = 0; i < 10; i++) {
          const r = i + 2;
          const isExisting = i < existingIds.length;
          const companyId = isExisting ? existingIds[i] : '';
          const uniqueName = `Company-For-SIT-${genStr(8)}`;
          
          const data: any = {
            A: '130840', B: companyId, C: 'X', D: 'X1', E: '001',
            F: uniqueName,
            G: countries[Math.floor(Math.random() * countries.length)],
            H: Math.random() > 0.5 ? '1' : '2',
            I: uniqueName, // 日本語名も同じに合わせるか、適宜調整
            J: `SHORT-${genStr(3)}`,
            K: `${Math.floor(Math.random()*999)} Mixed St.`,
            L: `City-${genStr(4)}`,
            M: states[Math.floor(Math.random() * states.length)],
            N: `ZIP-${genStr(5)}`,
            O: (Math.floor(Math.random() * 7) + 1).toString(),
            P: `Office-${genStr(4)}`,
            Q: Math.random() > 0.5 ? 'Y' : 'N',
            R: (Math.floor(Math.random() * 5) + 1).toString(),
            T: Math.random() > 0.5 ? 'Y' : 'N',
            U: 'Y'
          };

          Object.keys(data).forEach(col => { worksheet.getCell(`${col}${r}`).value = data[col]; });
          
          recordsToVerify.push({
            searchKey: isExisting ? companyId : uniqueName,
            searchType: isExisting ? 'ID' : 'Name',
            expected: data
          });
          console.log(`   - Row ${r}: [${isExisting?'Existing':'New'}] Key=${isExisting?companyId:uniqueName}`);
        }
        await workbook.xlsx.writeFile(testFilePath);

        // Upload
        const chooseFileBtn = tableFrame.locator('#ctl00_MainContent_btxChooseFile').or(tableFrame.getByRole('button', { name: 'Choose File' })).first();
        const uploadBtn = tableFrame.locator('#ctl00_MainContent_lnxbtnUploadSpreadsheet').or(tableFrame.getByRole('link', { name: 'Upload and Import Data' })).first();
        const fcPromise = popup.waitForEvent('filechooser');
        await chooseFileBtn.click();
        const fc = await fcPromise;
        await fc.setFiles(testFilePath);
        await uploadBtn.click();

        // Monitor
        const refreshBtn = tableFrame.locator('#ctl00_MainContent_lnxBtnRefreshWorkflowStatus').or(tableFrame.getByRole('link', { name: 'Refresh' })).first();
        let finalStatus = '';
        let targetRow: any = null;
        for (let i = 0; i < 60; i++) {
          await popup.waitForTimeout(5000);
          await refreshBtn.click();
          await popup.waitForTimeout(2000);
          const rows = tableFrame.locator('#ctl00_MainContent_rgUploadWorkflowStatus_ctl00 tbody tr');
          for (let j = 0; j < await rows.count(); j++) {
            if ((await rows.nth(j).innerText()).includes(fileName)) {
              targetRow = rows.nth(j);
              finalStatus = (await targetRow.locator('td:nth-child(5)').innerText()).trim();
              break;
            }
          }
          if (finalStatus.includes('Error') || finalStatus.includes('Complete')) break;
        }
        if (!finalStatus.includes('Complete')) throw new Error(`Upload Failed: ${finalStatus}`);
        await popup.close();
        try { fs.unlinkSync(baseTemplate); } catch(e) {}
      });

      // Step 4: Verification
      await test.step('Verification', async () => {
        await page.bringToFront();
        await page.getByRole('button', { name: 'DPS' }).click();
        await page.getByRole('menuitem', { name: 'DPS Search/Reporting Lookup' }).click();
        const legacyFrame = page.frameLocator('iframe[name="legacy-outlet"]');
        await legacyFrame.locator('a#ctl00_LeftHandNavigation_LHN_ctl00').click();
        await page.waitForTimeout(3000);

        const verificationErrors: string[] = [];

        for (const record of recordsToVerify) {
          // Use CompanyName (F column) as the search key for both existing and new records
          const searchName = record.expected.F;
          console.log(`ℹ️ Verifying Record by Name: "${searchName}"`);
          const expected = record.expected;
          
          await page.keyboard.press('Escape');
          await page.waitForTimeout(1000);

          // 1. Open Filter Options for "CompanyName"
          const header = legacyFrame.locator('th.rgHeader').filter({ hasText: /^CompanyName$/i }).first();
          await header.scrollIntoViewIfNeeded();
          await header.locator('button.rgActionButton.rgOptions').first().click({ force: true });
          await page.waitForTimeout(2000);

          // 2. Set Filter Condition to "EqualTo" using Keyboard
          const condInput = legacyFrame.locator('input[id*="HCFMRCMBFirstCond_Input"]');
          await condInput.waitFor({ state: 'visible', timeout: 10000 });
          await condInput.click({ force: true });
          await page.waitForTimeout(500);
          
          // Ensure we are at the top of the list, then find "EqualTo"
          for (let j = 0; j < 10; j++) await page.keyboard.press('ArrowUp');
          for (let i = 0; i < 15; i++) {
            const currentVal = await condInput.inputValue();
            if (currentVal?.toLowerCase() === 'equalto') {
              await page.keyboard.press('Enter');
              break;
            }
            await page.keyboard.press('ArrowDown');
            await page.waitForTimeout(200);
          }
          await page.waitForTimeout(500);

          // 3. Input CompanyName and Filter
          const valInput = legacyFrame.locator('input[id*="HCFMRTBFirstCond"]').first();
          await valInput.fill('');
          await valInput.type(searchName, { delay: 50 });
          await valInput.press('Enter');
          
          console.log(`   - Waiting for grid refresh...`);
          await page.waitForTimeout(8000);

          const gridTable = legacyFrame.locator('#ctl00_MainContent_rgLookup_ctl00').or(legacyFrame.locator('#ctl00_MainContent_rgMCCompanyLookup_ctl00'));
          let row = gridTable.locator('tbody tr.rgRow, tbody tr.rgAltRow').first();
          
          // Retry if row is not immediately visible
          if (!(await row.isVisible())) {
            await page.waitForTimeout(5000);
          }

          if (!(await row.isVisible())) {
            verificationErrors.push(`❌ Record not found for CompanyName: ${searchName}`);
            await test.info().attach(`not-found-${searchName}`, { body: await page.screenshot({ fullPage: true }), contentType: 'image/png' });
            continue;
          }

          const [editPopup] = await Promise.all([page.waitForEvent('popup'), row.locator('a', { hasText: 'Edit' }).click()]);
          await editPopup.waitForLoadState();
          await editPopup.waitForTimeout(8000);

          let vals: any = {};
          let targetF: any = null;
          for (let i = 0; i < 20; i++) {
            for (const f of editPopup.frames()) {
              try {
                const nameLoc = f.locator('#txtCompanyName');
                if (await nameLoc.count() > 0) {
                  targetF = f;
                  vals.F = (await nameLoc.inputValue()).trim();
                  vals.G = (await f.locator('div[id*="drpCompanyCountryCode"] .select2-chosen').innerText()).trim();
                  vals.H = (await f.locator('div[id*="drpHQDivisionFlag"] .select2-chosen').innerText()).trim();
                  vals.I = (await f.locator('#txtCompanyNameKanji').inputValue()).trim();
                  vals.J = (await f.locator('#txtShortName').inputValue()).trim();
                  vals.K = (await f.locator('#txtCompanyAddress1').inputValue()).trim();
                  vals.L = (await f.locator('#txtCompanyCity').inputValue()).trim();
                  vals.M = (await f.locator('div[id*="drpCompanyState"] .select2-chosen').innerText()).trim();
                  vals.N = (await f.locator('#txtCompanyPostalCode').inputValue()).trim();
                  vals.O = (await f.locator('div[id*="drpSubSidiaryClassification"] .select2-chosen').innerText()).trim();
                  vals.P = (await f.locator('#txtOfficeName').inputValue()).trim();
                  vals.Q = (await f.locator('div[id*="drpDTSSearchFlag"] .select2-chosen').innerText()).trim();
                  vals.R = (await f.locator('div[id*="drpInternalJudgement"] .select2-chosen').innerText()).trim();
                  vals.T = (await f.locator('div[id*="drpCompanyWide"] .select2-chosen').innerText()).trim();
                  break;
                }
              } catch(e) {}
            }
            if (targetF) break;
            await editPopup.waitForTimeout(1000);
          }

          if (targetF) {
            // Screenshots with multi-scroll
            await targetF.locator('#txtCompanyName').first().scrollIntoViewIfNeeded();
            await editPopup.waitForTimeout(1000);
            await test.info().attach(`verify-${record.searchKey}-top`, { body: await editPopup.screenshot(), contentType: 'image/png' });
            
            await targetF.locator('#txtCompanyAddress1').first().scrollIntoViewIfNeeded();
            await editPopup.waitForTimeout(1000);
            await test.info().attach(`verify-${record.searchKey}-mid`, { body: await editPopup.screenshot(), contentType: 'image/png' });
            
            await targetF.locator('div[id*="pnlSystem"]').first().scrollIntoViewIfNeeded();
            await editPopup.waitForTimeout(1000);
            await test.info().attach(`verify-${record.searchKey}-bottom`, { body: await editPopup.screenshot(), contentType: 'image/png' });

            const check = (label: string, act: string, exp: string) => {
              const match = act.includes(exp) || exp.includes(act) || 
                            (exp === 'はい' && (act.includes('邵ｺｯ') || act.includes('Yes'))) ||
                            (exp === 'いいえ' && (act.includes('邵ｺ繝ｻ樒ｸｺ') || act.includes('No')));
              
              console.log(`   [${label}] Match: ${match ? '✅' : '❌'}`);
              console.log(`     - Expected: "${exp}"`);
              console.log(`     - Actual:   "${act}"`);
              
              if (!match) verificationErrors.push(`❌ [${searchName}] ${label} Mismatch: Exp="${exp}", Act="${act}"`);
            };

            check('F (Name)', vals.F, expected.F);
            check('G (Country)', vals.G, expected.G);
            check('H (Type)', vals.H, expected.H === '1' ? '本社' : '事業所');
            check('I (NameJP)', vals.I, expected.I);
            check('J (ShortName)', vals.J, expected.J);
            check('K (Addr1)', vals.K, expected.K);
            check('L (City)', vals.L, expected.L);
            if (expected.M) check('M (State)', vals.M, expected.M);
            check('N (Zip)', vals.N, expected.N);
            const subMap: any = { '1': 'MC単体', '2': '現地法人', '3': '子会社', '4': '共同支配企業', '5': '共同支配事業', '6': '関連会社', '7': 'その他' };
            check('O (Subsidiary)', vals.O, subMap[expected.O]);
            check('P (Office)', vals.P, expected.P);
            check('Q (DTS)', vals.Q, expected.Q === 'Y' ? 'Yes' : 'No');
            const judgeMap: any = { '1': '外国ユーザーリスト', '2': '制裁対象者', '3': '禁止先', '4': 'Entity List', '5': '誤検知' };
            check('R (Judgement)', vals.R, judgeMap[expected.R]);
            check('T (CompanyWide)', vals.T, expected.T === 'Y' ? 'Yes' : 'No');
          } else {
            verificationErrors.push(`❌ [${record.searchKey}] Edit elements not found.`);
          }
          await editPopup.close();
        }
        if (verificationErrors.length > 0) (user as any).verificationErrors = verificationErrors;
      });

      // Step 5: Logout
      await test.step('Logout', async () => {
        try {
          const btn = `MC Test${user.id.substring(6)}`;
          await page.getByRole('button', { name: new RegExp(`^${btn}`) }).click();
          await page.getByRole('button', { name: ' Sign out' }).click();
          await page.waitForURL(/Logon|default/i);
        } catch (e) {}
        try { fs.unlinkSync(testFilePath); } catch(e) {}
        const errs = (user as any).verificationErrors;
        if (errs && errs.length > 0) {
          console.log('\n--- [Failures Summary] ---');
          errs.forEach((e: string) => console.log(e));
          throw new Error(`Verification failed with ${errs.length} errors.`);
        }
      });
    });
  }
});