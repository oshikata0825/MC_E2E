import { test, expect } from '@playwright/test';
import { BASE_URL } from './config';
import * as fs from 'fs';
import * as path from 'path';
import * as ExcelJS from 'exceljs';

/**
 * MC_DPS_Upload_007.test.ts
 * 
 * Tests spreadsheet upload error handling for restricted columns (H, M, O, R).
 * Performs 4 separate uploads with one invalid value in each target column.
 */

const allUsers = JSON.parse(fs.readFileSync(path.join(__dirname, 'account.dat'), 'utf8'));
const users = allUsers.filter(u => u.id === 'MCTest5' || u.id === 'MCTest9');

test.describe('Spreadsheet Upload - Invalid Option Validation (H, M, O, R)', () => {
  
  for (const user of users) {
    
    test(`Verify Error Handling for ${user.id}`, async ({ page }) => {
      test.setTimeout(1200000); // 4ケースあるため長めに設定
      
      console.log(`--- テスト開始: ${user.id} ---`);

      // Step 1: Login
      await page.goto(BASE_URL + '/Logon.aspx?ReturnUrl=%2fgtm%2fhome');
      await page.getByRole('textbox', { name: 'Company' }).fill(user.company);
      await page.getByRole('textbox', { name: 'Username' }).fill(user.id);
      await page.getByRole('textbox', { name: 'Password' }).fill(user.pass);
      await page.getByRole('link', { name: 'Login' }).click();

      // Step 2: Navigate to Spreadsheet Upload
      await page.getByRole('button', { name: 'Client Maintenance' }).click();
      await page.getByRole('menuitem', { name: 'Client Maintenance' }).click();
      const mf = await page.locator('iframe[name="legacy-outlet"]').contentFrame();
      if (!mf) throw new Error('Client Maintenance frame not found');
      await mf.locator('#ctl00_LeftHandNavigation_LHN_ctl00').click();
      await page.waitForTimeout(2000);

      const uploadActionLink = mf.locator('td', { hasText: '130840 SpreadSheet upload - DPS' }).locator('..').locator('a', { hasText: /Upload/i });

      // テストケースの定義
      const scenarios = [
        { col: 'H', name: 'Type (HQDivisionFlag)', invalid: '1234567890' },
        { col: 'O', name: 'Subsidiary (SubSidiaryClassification)', invalid: '1234567890' },
        { col: 'R', name: 'Judgement (InternalJudgement)', invalid: '1234567890' },
        { col: 'T', name: 'CompanyWide (CompanyWide)', invalid: '1234567890' },
        { col: 'U', name: 'ActiveFlag (ActiveFlag)', invalid: '1234567890' }
      ];

      const validationFailures: string[] = [];

      for (const scenario of scenarios) {
        await test.step(`Testing Invalid Value in Column ${scenario.col} (${scenario.name})`, async () => {
          console.log(`
>>> Scenario: ${scenario.name} (Col ${scenario.col}) with value "${scenario.invalid}"`);
          
          try {
            const popupPromise = page.waitForEvent('popup');
            await uploadActionLink.first().click();
            const popup = await popupPromise;
            await popup.waitForLoadState();

            let tableFrame: any = null;
            for (let i = 0; i < 30; i++) {
              for (const f of popup.frames()) {
                if (await f.locator('#lbxPageName:has-text("DPS Spreadsheet Upload")').count() > 0) {
                  tableFrame = f; break;
                }
              }
              if (tableFrame) break;
              await popup.waitForTimeout(1000);
            }

            // Download Template
            let downloadBtn: any = null;
            for (const f of popup.frames()) {
              const loc = f.locator('#ctl00_MainContent_hyxlnDownload').or(f.getByText('Download Excel Template')).first();
              if (await loc.count() > 0) { downloadBtn = loc; break; }
            }
            const [download] = await Promise.all([popup.waitForEvent('download', { timeout: 60000 }), downloadBtn.click({ force: true })]);
            const baseTemplate = path.join(__dirname, `base_template_007_${user.id}_${scenario.col}.xlsx`);
            await download.saveAs(baseTemplate);

            // Prepare Excel (1 row)
            const ts = new Date().toISOString().replace(/[:.-]/g, '').substring(0, 14);
            const fileName = `ERR_007_${scenario.col}_${ts}.xlsx`;
            const testFilePath = path.join(__dirname, fileName);          
            const wb = new ExcelJS.Workbook();
            await wb.xlsx.readFile(baseTemplate);
            const ws = wb.getWorksheet(1)!;

            const genAlphaNum = (l: number) => Math.random().toString(36).substring(2, 2 + l).toUpperCase();
            const compID = `For-SIT-${genAlphaNum(8)}`;
            const compName = `ErrTest-007-${scenario.col}-${genAlphaNum(4)}`;

            // 共通の正しい値 (H列は 1 または 2、O列は 1-7 をランダムに使用)
            const rowData: any = {
              A: '130840', B: compID, C: 'X', D: 'X1', E: '001', F: compName, G: 'JP',
              H: Math.random() > 0.5 ? '1' : '2', 
              I: compName, J: 'SHORT', K: 'Address', L: 'City', M: '8', N: '100-0001',
              O: (Math.floor(Math.random() * 7) + 1).toString(), 
              P: 'Office', Q: 'Y', R: '5', S: '1', T: 'Y', U: 'Y'
            };

            // ターゲットカラムのみ不正な値に上書き
            rowData[scenario.col] = scenario.invalid;

            // 書き込み
            const r = 2;
            Object.keys(rowData).forEach(key => {
              ws.getCell(`${key}${r}`).value = rowData[key];
            });

            await wb.xlsx.writeFile(testFilePath);
            console.log(`   - File created: ${fileName}`);

            // Upload
            let chooseFileBtn: any = null;
            let uploadBtn: any = null;
            let refreshBtn: any = null;
            for (const f of popup.frames()) {
              const cLoc = f.locator('#ctl00_MainContent_btxChooseFile').or(f.getByRole('button', { name: 'Choose File' }));
              if (await cLoc.count() > 0) chooseFileBtn = cLoc.first();
              const uLoc = f.locator('#ctl00_MainContent_lnxbtnUploadSpreadsheet').or(f.getByRole('link', { name: 'Upload and Import Data' }));
              if (await uLoc.count() > 0) uploadBtn = uLoc.first();
              const rLoc = f.locator('#ctl00_MainContent_lnxBtnRefreshWorkflowStatus').or(f.getByRole('link', { name: 'Refresh' }));
              if (await rLoc.count() > 0) refreshBtn = rLoc.first();
            }

            const fileChooserPromise = popup.waitForEvent('filechooser');
            await chooseFileBtn.click();
            const fileChooser = await fileChooserPromise;
            await fileChooser.setFiles(testFilePath);
            await uploadBtn.click();
            console.log(`   - Uploaded.`);

            // Monitor Status
            let finalStatus = '';
            let targetRow: any = null;
            for (let i = 0; i < 30; i++) {
              await popup.waitForTimeout(5000);
              await refreshBtn.click();
              await popup.waitForTimeout(2000);
              
              let workflowFrame: any = popup;
              for(const f of popup.frames()){
                if(await f.locator('#ctl00_MainContent_rgUploadWorkflowStatus_ctl00 tbody tr').count() > 0) { workflowFrame = f; break; }
              }

              const rows = workflowFrame.locator('#ctl00_MainContent_rgUploadWorkflowStatus_ctl00 tbody tr');
              const count = await rows.count();
              for (let j = 0; j < count; j++) {
                if ((await rows.nth(j).innerText()).includes(fileName)) {
                  targetRow = rows.nth(j);
                  finalStatus = (await targetRow.locator('td:nth-child(5)').innerText()).trim();
                  break;
                }
              }
              if (finalStatus.includes('Error') || finalStatus.includes('Complete')) break;
            }

            console.log(`   - Final Status: "${finalStatus}"`);

            // Verification
            if (finalStatus.includes('Error')) {
              console.log(`   ✅ Success: Correctly identified invalid input in Col ${scenario.col}.`);
              const errorCountLink = targetRow.locator('a[datafield="ErrorCount"]');
              if (await errorCountLink.count() > 0) {
                await errorCountLink.click();
                await popup.waitForTimeout(5000);
                await test.info().attach(`error-col-${scenario.col}-${user.id}`, { 
                  body: await popup.screenshot({ fullPage: true }), 
                  contentType: 'image/png' 
                });
              }
            } else {
              const msg = `❌ FAILED: Upload reached "${finalStatus}" despite invalid input in Col ${scenario.col}.`;
              console.error(msg);
              validationFailures.push(`Column ${scenario.col} (${scenario.name}) - Status: ${finalStatus}`);
              await test.info().attach(`fail-col-${scenario.col}-${user.id}`, { body: await popup.screenshot({ fullPage: true }), contentType: 'image/png' });
            }

            await popup.close();
            try { fs.unlinkSync(testFilePath); fs.unlinkSync(baseTemplate); } catch(e) {}
          } catch (error) {
            console.error(`❌ Unexpected error in scenario ${scenario.col}: ${error.message}`);
            validationFailures.push(`Column ${scenario.col} (${scenario.name}) - Execution Error: ${error.message}`);
          }
        });
      }

      if (validationFailures.length > 0) {
        console.log(`
=========================================`);
        console.log(`❌ Validation FAILED for the following columns:`);
        validationFailures.forEach(f => console.log(`   - ${f}`));
        console.log(`=========================================
`);
        throw new Error(`Spreadsheet validation failed for ${validationFailures.length} columns.`);
      } else {
        console.log(`
✅ All invalid input scenarios correctly handled (returned Error status).`);
      }

      // Logout
      const userButtonName = `MC Test${user.id.substring(6)}`;
      await page.getByRole('button', { name: new RegExp(`^${userButtonName}`) }).click();
      await page.getByRole('button', { name: ' Sign out' }).click();
    });
  }
});