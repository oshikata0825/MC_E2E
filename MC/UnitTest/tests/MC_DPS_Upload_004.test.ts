import { test, expect } from '@playwright/test';
import { BASE_URL } from './config';
import * as fs from 'fs';
import * as path from 'path';
import * as ExcelJS from 'exceljs';

/**
 * MC_DPS_Upload_004.test.ts
 * 
 * Tests spreadsheet upload error handling with a mix of New and Existing records.
 * Uses unique filenames and robust status monitoring from MC_DPS_Upload_003.
 */

const allUsers = JSON.parse(fs.readFileSync(path.join(__dirname, 'account.dat'), 'utf8'));
const users = allUsers.filter(u => u.id === 'MCTest5' || u.id === 'MCTest9');

test.describe('Spreadsheet Upload - New/Existing Mix Error Handling', () => {
  
  for (const user of users) {
    
    test(`Verify Error Handling for ${user.id}`, async ({ page }) => {
      test.setTimeout(600000); 
      
      console.log(`--- テスト開始: ${user.id} ---`);

      // Step 1: Login
      await page.goto(BASE_URL + '/Logon.aspx?ReturnUrl=%2fgtm%2fhome');
      await page.getByRole('textbox', { name: 'Company' }).fill(user.company);
      await page.getByRole('textbox', { name: 'Username' }).fill(user.id);
      await page.getByRole('textbox', { name: 'Password' }).fill(user.pass);
      await page.getByRole('link', { name: 'Login' }).click();

      // Step 2: Navigate and Check Visibility
      const mainFrameElement = page.locator('iframe[name="legacy-outlet"]');
      await page.getByRole('button', { name: 'Client Maintenance' }).click();
      await page.getByRole('menuitem', { name: 'Client Maintenance' }).click();
      await mainFrameElement.waitFor({ state: 'attached', timeout: 30000 });
      const mainFrame = await mainFrameElement.contentFrame();
      if (!mainFrame) throw new Error('Main frame not found');

      const uploadLink = mainFrame.locator('#ctl00_LeftHandNavigation_LHN_ctl00');
      const isAuthorizedUser = (user.id === 'MCTest5' || user.id === 'MCTest9');
      
      if (isAuthorizedUser) {
        await uploadLink.waitFor({ state: 'visible', timeout: 30000 }).catch(() => {});
      }

      if (!(await uploadLink.isVisible())) {
        if (isAuthorizedUser) throw new Error(`❌ FAILED: ${user.id} should see Spreadsheet Upload.`);
        console.log(`✅ ${user.id}: Spreadsheet Upload is hidden (expected).`);
        return;
      }

      // --- [A] Get an existing CompanyID (For-SIT-) from DPS Lookup ---
      console.log('ℹ️ Fetching an existing "For-SIT-" CompanyID...');
      await page.getByRole('button', { name: 'DPS' }).click();
      await page.getByRole('menuitem', { name: 'DPS Search/Reporting Lookup' }).click();
      const dpsFrame = await page.locator('iframe[name="legacy-outlet"]').contentFrame();
      await dpsFrame!.locator('a#ctl00_LeftHandNavigation_LHN_ctl00').click();
      await page.waitForTimeout(5000);

      const header = dpsFrame!.locator('th.rgHeader').filter({ hasText: /^CompanyID$/i });
      await header.locator('button.rgActionButton.rgOptions').first().click({ force: true });
      await page.waitForTimeout(2000);
      const cond = dpsFrame!.locator('input[id*="HCFMRCMBFirstCond_Input"]');
      await cond.click({ force: true });
      for(let j=0; j<10; j++) await page.keyboard.press('ArrowUp');
      for(let i=0; i<15; i++){
        const val = await cond.inputValue();
        if(val?.toLowerCase() === 'contains'){ await page.keyboard.press('Enter'); break; }
        await page.keyboard.press('ArrowDown');
        await page.waitForTimeout(200);
      }
      await dpsFrame!.locator('input[id*="HCFMRTBFirstCond"]').first().fill('For-SIT-');
      await page.keyboard.press('Enter');
      await page.waitForTimeout(5000);
      
      const existingId = (await dpsFrame!.locator('tbody tr.rgRow td:nth-child(4), tbody tr.rgAltRow td:nth-child(4)').first().innerText()).trim();
      if (!existingId) throw new Error('Could not find a "For-SIT-" CompanyID to test.');
      console.log(`✅ Found target CompanyID: ${existingId}`);

      // Back to Upload
      await page.getByRole('button', { name: 'Client Maintenance' }).click();
      await page.getByRole('menuitem', { name: 'Client Maintenance' }).click();
      const mf = await page.locator('iframe[name="legacy-outlet"]').contentFrame();
      await mf!.locator('#ctl00_LeftHandNavigation_LHN_ctl00').click();
      await page.waitForTimeout(2000);

      const uploadActionLink = mf!.locator('td', { hasText: '130840 SpreadSheet upload - DPS' }).locator('..').locator('a', { hasText: /Upload/i });
      
      const cases = [
        { name: 'Case 1: [Row 1: New, Row 2: Existing]', order: ['NEW', 'EXISTING'] },
        { name: 'Case 2: [Row 1: Existing, Row 2: New]', order: ['EXISTING', 'NEW'] }
      ];

      for (const testCase of cases) {
        await test.step(testCase.name, async () => {
          console.log(`
--- Starting ${testCase.name} ---`);
          
          const popupPromise = page.waitForEvent('popup');
          await uploadActionLink.first().click();
          const popup = await popupPromise;
          await popup.waitForLoadState();

          let titleFound = false;
          let tableFrame: any = null;
          for (let i = 0; i < 30; i++) {
            for (const f of popup.frames()) {
              if (await f.locator('#lbxPageName:has-text("DPS Spreadsheet Upload")').count() > 0) {
                tableFrame = f; titleFound = true; break;
              }
            }
            if (titleFound) break;
            await popup.waitForTimeout(1000);
          }

          // Download Template
          const downloadBtn = tableFrame.locator('#ctl00_MainContent_hyxlnDownload').or(tableFrame.getByText('Download Excel Template')).first();
          const [download] = await Promise.all([popup.waitForEvent('download'), downloadBtn.click()]);
          const baseTemplate = path.join(__dirname, `base_template_${user.id}.xlsx`);
          await download.saveAs(baseTemplate);

          // Prepare Test File
          const ts = new Date().toISOString().replace(/[:.-]/g, '').substring(0, 14);
          const fileName = `MIX_${testCase.order[0].substring(0,1)}_${ts}.xlsx`;
          const testFilePath = path.join(__dirname, fileName);          
          const wb = new ExcelJS.Workbook();
          await wb.xlsx.readFile(baseTemplate);
          const ws = wb.getWorksheet(1)!;
          const genStr = (l: number) => Math.random().toString(36).substring(2, 2 + l).toUpperCase();

          for (let i = 0; i < 2; i++) {
            const r = i + 2;
            const type = testCase.order[i];
            const compId = (type === 'NEW') ? '' : existingId;
            const grpCode = `Invalid-${genStr(3)}`;
            const compName = `Test-004-${type}-${genStr(5)}`;
            const country = 'JP';

            ws.getCell(`A${r}`).value = '130840';
            ws.getCell(`B${r}`).value = compId;
            ws.getCell(`C${r}`).value = grpCode;
            ws.getCell(`D${r}`).value = 'X1';
            ws.getCell(`E${r}`).value = '001';
            ws.getCell(`F${r}`).value = compName;
            ws.getCell(`G${r}`).value = country;
            ws.getCell(`Q${r}`).value = 'Y';
            ws.getCell(`U${r}`).value = 'Y';
            
            console.log(`   - Row ${r} [${type}]:`);
            console.log(`     A (ClientID): 130840`);
            console.log(`     B (CompanyID): "${compId}"`);
            console.log(`     C (GroupCode): "${grpCode}"`);
            console.log(`     F (CompanyName): "${compName}"`);
            console.log(`     G (Country): "${country}"`);
            console.log(`     Q (DTSSearchFlag): "Y"`);
            console.log(`     U (ActiveFlag): "Y"`);
          }
          await wb.xlsx.writeFile(testFilePath);

          // Upload
          const chooseFileBtn = tableFrame.locator('#ctl00_MainContent_btxChooseFile').or(tableFrame.getByRole('button', { name: 'Choose File' })).first();
          const uploadBtn = tableFrame.locator('#ctl00_MainContent_lnxbtnUploadSpreadsheet').or(tableFrame.getByRole('link', { name: 'Upload and Import Data' })).first();
          const refreshBtn = tableFrame.locator('#ctl00_MainContent_lnxBtnRefreshWorkflowStatus').or(tableFrame.getByRole('link', { name: 'Refresh' })).first();

          const fileChooserPromise = popup.waitForEvent('filechooser');
          await chooseFileBtn.click();
          const fileChooser = await fileChooserPromise;
          await fileChooser.setFiles(testFilePath);
          await uploadBtn.click();
          console.log(`ℹ️ Uploaded: ${fileName}`);

          // Monitor Status
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
              console.log(`   - Attempt ${i+1}: Status = "${finalStatus}"`);
              if (finalStatus.includes('Error') || finalStatus.includes('Complete')) break;
            } else {
              console.log(`   - Attempt ${i+1}: Waiting for record "${fileName}"...`);
            }
          }

          // Verify result and capture Error Details
          if (finalStatus.includes('Error')) {
            console.log(`✅ Expected Error captured for ${testCase.name}`);
            const errorCountLink = targetRow.locator('a[datafield="ErrorCount"]');
            if (await errorCountLink.count() > 0) {
              await errorCountLink.click();
              await popup.waitForTimeout(8000);
              await test.info().attach(`error-detail-${testCase.order[0]}-${user.id}`, { 
                body: await popup.screenshot({ fullPage: true }), 
                contentType: 'image/png' 
              });
              console.log(`📸 Error details captured.`);
            }
          } else if (finalStatus.includes('Complete')) {
            throw new Error(`❌ FAILED: ${testCase.name} reached "Complete" despite invalid data.`);
          } else {
            throw new Error(`❌ FAILED: ${testCase.name} did not reach a final state in time.`);
          }

          // Cleanup RadWindows
          await popup.evaluate(() => {
            const wm = (window as any).Telerik?.Web?.UI?.RadWindowUtils?.get_windowManager();
            if (wm) wm.get_windows().forEach((w: any) => { if (w.is_visible()) w.close(); });
            document.querySelectorAll('[id*="RadWindowWrapper"], [id*="RadModalBackground"], .TelerikModalOverlay').forEach(el => el.remove());
          }).catch(() => {});
          
          await popup.close();
          try { fs.unlinkSync(testFilePath); fs.unlinkSync(baseTemplate); } catch(e) {}
        });
      }

      // Logout
      const userButtonName = `MC Test${user.id.substring(6)}`;
      await page.getByRole('button', { name: new RegExp(`^${userButtonName}`) }).click();
      await page.getByRole('button', { name: ' Sign out' }).click();
    });
  }
});
