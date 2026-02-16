import { test, expect } from '@playwright/test';
import { BASE_URL } from './config';
import * as fs from 'fs';
import * as path from 'path';
import * as ExcelJS from 'exceljs';

/**
 * MC_DPS_Upload_005.test.ts
 * 
 * Tests spreadsheet upload for a single NEW record using values extracted from an existing record.
 */

const allUsers = JSON.parse(fs.readFileSync(path.join(__dirname, 'account.dat'), 'utf8'));
const users = allUsers.filter(u => u.id === 'MCTest5' || u.id === 'MCTest9');

test.describe('Spreadsheet Upload - Single New Record Validation', () => {
  
  for (const user of users) {
    
    test(`Verify Upload for ${user.id}`, async ({ page }) => {
      test.setTimeout(600000); 
      
      console.log(`--- テスト開始: ${user.id} ---`);

      // Step 1: Login
      await page.goto(BASE_URL + '/Logon.aspx?ReturnUrl=%2fgtm%2fhome');
      await page.getByRole('textbox', { name: 'Company' }).fill(user.company);
      await page.getByRole('textbox', { name: 'Username' }).fill(user.id);
      await page.getByRole('textbox', { name: 'Password' }).fill(user.pass);
      await page.getByRole('link', { name: 'Login' }).click();

      // Step 2: Access DPS Lookup to find an existing record
      console.log('ℹ️ Fetching an existing "For-SIT-" CompanyID...');
      await page.getByRole('button', { name: 'DPS' }).click();
      await page.getByRole('menuitem', { name: 'DPS Search/Reporting Lookup' }).click();
      
      const mainFrameElement_dps = page.locator('iframe[name="legacy-outlet"]');
      await mainFrameElement_dps.waitFor({ state: 'attached', timeout: 30000 });
      const dpsFrame = await mainFrameElement_dps.contentFrame();
      if (!dpsFrame) throw new Error('DPS main frame not found');

      await dpsFrame.getByRole('link', { name: 'MC Company Lookup' }).click();
      await page.waitForTimeout(3000);

      const header = dpsFrame.locator('th.rgHeader').filter({ hasText: /^CompanyID$/i });
      await header.locator('button.rgActionButton.rgOptions').first().click({ force: true });
      await page.waitForTimeout(2000);
      
      const cond = dpsFrame.locator('input[id*="HCFMRCMBFirstCond_Input"]');
      await cond.click({ force: true });
      await page.waitForTimeout(1000);
      
      const startsWithItem = dpsFrame.locator('li.rcbItem', { hasText: /^StartsWith$/i }).or(page.locator('li.rcbItem', { hasText: /^StartsWith$/i })).first();
      await startsWithItem.click();
      await page.waitForTimeout(1000);

      const valInput = dpsFrame.locator('input[id*="HCFMRTBFirstCond"]').first();
      await valInput.fill('For-SIT-');
      await page.keyboard.press('Enter');
      await page.waitForTimeout(5000);
      
      const firstRow = dpsFrame.locator('tbody tr.rgRow, tbody tr.rgAltRow').first();
      const existingId = (await firstRow.locator('td:nth-child(4)').innerText()).trim();
      if (!existingId) throw new Error('Could not find a "For-SIT-" CompanyID to test.');

      // Open Edit popup to extract details
      const editLink = firstRow.locator('a:has-text("Edit")');
      const editPopupPromise = page.context().waitForEvent('page', { timeout: 30000 });
      await editLink.click({ force: true });
      const editPopup = await editPopupPromise;
      await editPopup.waitForLoadState('load');
      await editPopup.waitForTimeout(5000);

      const popupIframeElement = editPopup.locator('iframe[name="legacy-outlet"]');
      let popupRoot: any = editPopup;
      if (await popupIframeElement.count() > 0) {
        popupRoot = editPopup.frameLocator('iframe[name="legacy-outlet"]');
      }

      const getText = async (selector: string) => {
        const loc = popupRoot.locator(selector);
        return (await loc.count() > 0) ? (await loc.innerText()).trim() : '';
      };

      const groupCodeSelector = '#s2id_ctl00_MainContent_tabs_tabCompanyInfo_pnlBasic_Info_drpGroupOrgCode .select2-chosen';
      const divCodeSelector = '#s2id_ctl00_MainContent_tabs_tabCompanyInfo_pnlBasic_Info_drpDivisionOrgCode .select2-chosen';
      const buCodeSelector = '#s2id_ctl00_MainContent_tabs_tabCompanyInfo_pnlBasic_Info_drpBUCode .select2-chosen';
      const countryCodeSelector = '#s2id_ctl00_MainContent_tabs_tabCompanyInfo_pnlBasic_Info_drpCompanyCountryCode .select2-chosen';

      const existingGroupCode = (await getText(groupCodeSelector)).split(':')[0];
      const existingDivCode = (await getText(divCodeSelector)).split(':')[0];
      const existingBUCode = (await getText(buCodeSelector)).split(':')[0];
      const existingCountryCode = (await getText(countryCodeSelector)).split('-')[0];
      const existingCompName = await popupRoot.locator('#txtCompanyName').inputValue();

      console.log(`
=========================================
`);
      console.log(`Target Record: ${existingId}`);
      console.log(`Group: ${existingGroupCode}, Div: ${existingDivCode}, BU: ${existingBUCode}`);
      console.log(`Country: ${existingCountryCode}, Name: ${existingCompName}`);
      console.log(`=========================================
`);

      await editPopup.close();

      // Step 3: Navigate to Spreadsheet Upload
      await page.getByRole('button', { name: 'Client Maintenance' }).click();
      await page.getByRole('menuitem', { name: 'Client Maintenance' }).click();
      const mf = await page.locator('iframe[name="legacy-outlet"]').contentFrame();
      if (!mf) throw new Error('Client Maintenance frame not found');
      await mf.locator('#ctl00_LeftHandNavigation_LHN_ctl00').click();
      await page.waitForTimeout(2000);

      const uploadActionLink = mf.locator('td', { hasText: '130840 SpreadSheet upload - DPS' }).locator('..').locator('a', { hasText: /Upload/i });
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
      if (!tableFrame) throw new Error('Spreadsheet Upload frame not found');

      // Download Template
      let downloadBtn: any = null;
      for (const f of popup.frames()) {
        const loc = f.locator('#ctl00_MainContent_hyxlnDownload').or(f.getByText('Download Excel Template')).first();
        if (await loc.count() > 0) { downloadBtn = loc; break; }
      }
      if (!downloadBtn) downloadBtn = popup.locator('#ctl00_MainContent_hyxlnDownload').or(popup.getByText('Download Excel Template')).first();

      const [download] = await Promise.all([popup.waitForEvent('download', { timeout: 60000 }), downloadBtn.click({ force: true })]);
      const baseTemplate = path.join(__dirname, `base_template_${user.id}_005.xlsx`);
      await download.saveAs(baseTemplate);

      // Step 4: Prepare Excel and Upload
      const ts = new Date().toISOString().replace(/[:.-]/g, '').substring(0, 14);
      const fileName = `NEW_RECORD_${ts}.xlsx`;
      const testFilePath = path.join(__dirname, fileName);          
      const wb = new ExcelJS.Workbook();
      await wb.xlsx.readFile(baseTemplate);
      const ws = wb.getWorksheet(1)!;

      const genAlphaNum = (l: number) => {
        const chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789';
        let res = '';
        for(let i=0; i<l; i++) res += chars.charAt(Math.floor(Math.random() * chars.length));
        return res;
      };

      ws.getCell('A2').value = '130840';
      ws.getCell('B2').value = ''; // New Record
      ws.getCell('C2').value = existingGroupCode;
      ws.getCell('D2').value = existingDivCode;
      ws.getCell('E2').value = existingBUCode;
      ws.getCell('F2').value = existingCompName;
      ws.getCell('G2').value = existingCountryCode;
      ws.getCell('H2').value = '1';
      ws.getCell('I2').value = existingCompName;
      ws.getCell('J2').value = genAlphaNum(6);
      ws.getCell('K2').value = 'Test Address 005';
      ws.getCell('L2').value = 'Test City';
      ws.getCell('M2').value = '8';
      ws.getCell('N2').value = '100-0001';
      ws.getCell('O2').value = '1';
      ws.getCell('P2').value = 'Test Office';
      ws.getCell('Q2').value = 'Y';
      ws.getCell('R2').value = '5';
      ws.getCell('S2').value = '1';
      ws.getCell('T2').value = 'Y';
      ws.getCell('U2').value = 'Y';

      await wb.xlsx.writeFile(testFilePath);

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
      console.log(`ℹ️ Uploaded: ${fileName}`);

      // Step 5: Monitor Status
      let finalStatus = '';
      let targetRow: any = null;
      for (let i = 0; i < 30; i++) {
        await popup.waitForTimeout(5000);
        await refreshBtn.click();
        await popup.waitForTimeout(2000);
        
        let workflowFrame: any = popup;
        for(const f of popup.frames()){
           if(await f.locator('#ctl00_MainContent_rgUploadWorkflowStatus_ctl00 tbody tr').count() > 0) {
             workflowFrame = f; break;
           }
        }

        const rows = workflowFrame.locator('#ctl00_MainContent_rgUploadWorkflowStatus_ctl00 tbody tr');
        const count = await rows.count();
        for (let j = 0; j < count; j++) {
          const rowText = await rows.nth(j).innerText();
          if (rowText.includes(fileName)) {
            targetRow = rows.nth(j);
            finalStatus = (await targetRow.locator('td:nth-child(5)').innerText()).trim();
            break;
          }
        }
        console.log(`   - Attempt ${i+1}: Status = "${finalStatus}"`);
        if (finalStatus.includes('Error') || finalStatus.includes('Complete')) break;
      }

      // Verify result: We expect "Error" because we are trying to add a duplicate CompanyName as a New Record
      if (finalStatus.includes('Error')) {
        console.log(`✅ Expected Result: Upload failed with "Error" (Validation worked).`);
        const errorCountLink = targetRow.locator('a[datafield="ErrorCount"]');
        if (await errorCountLink.count() > 0) {
          await errorCountLink.click();
          await popup.waitForTimeout(8000);
          await test.info().attach(`error-detail-${user.id}`, { 
            body: await popup.screenshot({ fullPage: true }), 
            contentType: 'image/png' 
          });
          console.log(`📸 Error details captured in report.`);
        }
      } else if (finalStatus.includes('Complete')) {
        const failMsg = `❌ FAILED: Upload reached "Complete". Duplicate CompanyName should have been rejected for a NEW record.`;
        console.error(failMsg);
        await test.info().attach(`unexpected-complete-${user.id}`, { 
          body: await popup.screenshot({ fullPage: true }), 
          contentType: 'image/png' 
        });
        throw new Error(failMsg);
      } else {
        throw new Error(`❌ FAILED: Did not reach a final state in time. Current status: "${finalStatus}"`);
      }

      await popup.close();
      try { fs.unlinkSync(testFilePath); fs.unlinkSync(baseTemplate); } catch(e) {}

      // Logout
      const userButtonName = `MC Test${user.id.substring(6)}`;
      await page.getByRole('button', { name: new RegExp(`^${userButtonName}`) }).click();
      await page.getByRole('button', { name: ' Sign out' }).click();
    });
  }
});