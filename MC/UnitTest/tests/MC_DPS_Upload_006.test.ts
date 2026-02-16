import { test, expect } from '@playwright/test';
import { BASE_URL } from './config';
import * as fs from 'fs';
import * as path from 'path';
import * as ExcelJS from 'exceljs';

/**
 * MC_DPS_Upload_006.test.ts
 * 
 * Tests spreadsheet upload for 1-3 random new records.
 * Column B is set to "SS-Uplaod-XXXXXXXX" (8 random chars).
 */

const allUsers = JSON.parse(fs.readFileSync(path.join(__dirname, 'account.dat'), 'utf8'));
const users = allUsers.filter(u => u.id === 'MCTest5' || u.id === 'MCTest9');

test.describe('Spreadsheet Upload - Random New Records (1-3 rows)', () => {
  
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

      // Step 2: Navigate to Spreadsheet Upload
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
      const baseTemplate = path.join(__dirname, `base_template_${user.id}_006.xlsx`);
      await download.saveAs(baseTemplate);

      // Step 3: Prepare Excel with 1-3 Random Rows
      const ts = new Date().toISOString().replace(/[:.-]/g, '').substring(0, 14);
      const fileName = `NEW_RECORD_006_${ts}.xlsx`;
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

      const rowCount = Math.floor(Math.random() * 3) + 1; // 1 to 3 rows
      console.log(`ℹ️ Generating ${rowCount} rows for upload.`);

      const uploadedData: any[] = [];

      for (let i = 0; i < rowCount; i++) {
        const r = i + 2;
        const randomID = `SS-Uplaod-${genAlphaNum(8)}`;
        const compName = `Test-006-${genAlphaNum(6)}`;
        const data = {
          clientId: '130840',
          companyId: randomID,
          groupCode: 'X',
          divCode: 'X1',
          buCode: '001',
          compName: compName,
          country: 'JP',
          type: '1',
          shortName: genAlphaNum(6),
          address: 'Test Address 006',
          city: 'Test City',
          state: '8',
          zip: '100-0001',
          subsidiary: '1',
          office: 'Test Office',
          dtsSearch: 'Y',
          judgement: '5',
          companyWide: 'Y',
          active: 'Y'
        };

        ws.getCell(`A${r}`).value = data.clientId;
        ws.getCell(`B${r}`).value = data.companyId;
        ws.getCell(`C${r}`).value = data.groupCode;
        ws.getCell(`D${r}`).value = data.divCode;
        ws.getCell(`E${r}`).value = data.buCode;
        ws.getCell(`F${r}`).value = data.compName;
        ws.getCell(`G${r}`).value = data.country;
        ws.getCell(`H${r}`).value = data.type;
        ws.getCell(`I${r}`).value = data.compName;
        ws.getCell(`J${r}`).value = data.shortName;
        ws.getCell(`K${r}`).value = data.address;
        ws.getCell(`L${r}`).value = data.city;
        ws.getCell(`M${r}`).value = data.state;
        ws.getCell(`N${r}`).value = data.zip;
        ws.getCell(`O${r}`).value = data.subsidiary;
        ws.getCell(`P${r}`).value = data.office;
        ws.getCell(`Q${r}`).value = data.dtsSearch;
        ws.getCell(`R${r}`).value = data.judgement;
        ws.getCell(`S${r}`).value = '1';
        ws.getCell(`T${r}`).value = data.companyWide;
        ws.getCell(`U${r}`).value = data.active;

        uploadedData.push(data);
        console.log(`   - Row ${r}: B=${randomID}, F=${compName}`);
      }

      await wb.xlsx.writeFile(testFilePath);

      // Step 4: Upload
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

      // Verify upload status
      if (finalStatus.includes('Complete')) {
        console.log(`✅ Success: Upload reached "Complete".`);
      } else {
        if (finalStatus.includes('Error')) {
          const errorCountLink = targetRow.locator('a[datafield="ErrorCount"]');
          if (await errorCountLink.count() > 0) {
            await errorCountLink.click();
            await popup.waitForTimeout(8000);
            await test.info().attach(`error-detail-${user.id}`, { body: await popup.screenshot({ fullPage: true }), contentType: 'image/png' });
          }
        }
        throw new Error(`❌ FAILED: Upload did not reach "Complete". Status: "${finalStatus}"`);
      }

      await popup.close();

      // --- Step 6: Verification Phase ---
      console.log('ℹ️ Starting verification phase in DPS Search/Reporting Lookup...');
      await page.bringToFront();
      await page.getByRole('button', { name: 'DPS' }).click();
      await page.getByRole('menuitem', { name: 'DPS Search/Reporting Lookup' }).click();

      const dpsFrameElement = page.locator('iframe[name="legacy-outlet"]');
      await dpsFrameElement.waitFor({ state: 'attached', timeout: 30000 });
      const dpsFrame = await dpsFrameElement.contentFrame();
      if (!dpsFrame) throw new Error('DPS frame not found');

      // Click "MC Company Lookup" link inside the frame
      await dpsFrame.getByRole('link', { name: 'MC Company Lookup' }).click();
      await page.waitForTimeout(3000);

      for (const data of uploadedData) {
        console.log(`ℹ️ Verifying CompanyID: "${data.companyId}"`);
        
        // 1. Filter by CompanyID
        await page.keyboard.press('Escape');
        const header = dpsFrame.locator('th.rgHeader').filter({ hasText: /^CompanyID$/i });
        await header.locator('button.rgActionButton.rgOptions').first().click({ force: true });
        await page.waitForTimeout(2000);

        const cond = dpsFrame.locator('input[id*="HCFMRCMBFirstCond_Input"]');
        await cond.click({ force: true });
        for(let j=0; j<10; j++) await page.keyboard.press('ArrowUp');
        for(let i=0; i<15; i++){
          const val = await cond.inputValue();
          if(val?.toLowerCase() === 'equalto'){ await page.keyboard.press('Enter'); break; }
          await page.keyboard.press('ArrowDown');
          await page.waitForTimeout(200);
        }
        await dpsFrame.locator('input[id*="HCFMRTBFirstCond"]').first().fill(data.companyId);
        await page.keyboard.press('Enter');
        await page.waitForTimeout(5000);

        // 2. Take Screenshot
        const row = dpsFrame.locator('tbody tr.rgRow, tbody tr.rgAltRow').first();
        if (!(await row.isVisible())) throw new Error(`❌ FAILED: Record with CompanyID "${data.companyId}" not found.`);
        
        await test.info().attach(`lookup-result-${data.companyId}`, { 
          body: await page.screenshot({ fullPage: true }), 
          contentType: 'image/png' 
        });

        // 3. Open Edit popup
        const editLink = row.locator('a:has-text("Edit")');
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

        // 4. Robust extraction helper
        const safeGet = async (selector: string, isInput = false) => {
          const loc = popupRoot.locator(selector).first();
          try {
            await loc.waitFor({ state: 'visible', timeout: 5000 });
            if (isInput) return (await loc.inputValue()).trim();
            return (await loc.innerText()).trim();
          } catch (e) {
            console.log(`⚠️ Warning: Could not get value for ${selector}`);
            return '';
          }
        };

        const actualCompID = await safeGet('#lblCompanyID, #txtCompanyID', true);
        const actualName = await safeGet('#txtCompanyName', true);
        const actualCountry = await safeGet('#s2id_ctl00_MainContent_tabs_tabCompanyInfo_pnlBasic_Info_drpCompanyCountryCode .select2-chosen');
        const actualType = await safeGet('#s2id_ctl00_MainContent_tabs_tabCompanyInfo_pnlBasic_Info_drpHQDivisionFlag .select2-chosen');
        const actualNameJP = await safeGet('#txtCompanyNameKanji', true);
        const actualShortName = await safeGet('#txtShortName', true);
        const actualAddress = await safeGet('#txtCompanyAddress1', true);
        const actualCity = await safeGet('#txtCompanyCity', true);
        const actualZip = await safeGet('#txtCompanyPostalCode', true);
        const actualSubsidiary = await safeGet('#s2id_ctl00_MainContent_tabs_tabCompanyInfo_pnlBasic_Info_drpSubSidiaryClassification .select2-chosen');
        const actualOffice = await safeGet('#txtOfficeName', true);
        const actualDTS = await safeGet('#s2id_ctl00_MainContent_tabs_tabCompanyInfo_pnlSystem_drpDTSSearchFlag .select2-chosen');
        const actualJudgement = await safeGet('#s2id_ctl00_MainContent_tabs_tabCompanyInfo_pnlSystem_drpInternalJudgement .select2-chosen');
        
        // Audit パネルへスクロール
        await popupRoot.locator('[id*="pnlAudit"]').last().scrollIntoViewIfNeeded().catch(() => {});

        const actualCompanyWide = await safeGet('#s2id_ctl00_MainContent_tabs_tabCompanyInfo_pnlAudit_drpCompanyWide .select2-chosen');
        const actualActive = await safeGet('#s2id_ctl00_MainContent_tabs_tabCompanyInfo_pnlAudit_drpActiveFlag .select2-chosen');

        // 5. Verification logic (B-U)
        const mismatches: string[] = [];
        const checkValue = (label: string, actual: string, expected: string) => {
          let isMatch = false;
          if (expected === '' && actual === '') {
            isMatch = true;
          } else if (expected !== '' && actual === '') {
            isMatch = false;
          } else {
            isMatch = actual.toLowerCase().includes(expected.toLowerCase()) || expected.toLowerCase().includes(actual.toLowerCase());
          }

          console.log(`     [Check] ${label}: Expected="${expected}", Actual="${actual}" ${isMatch ? '✅' : '❌'}`);
          if (!isMatch) {
            mismatches.push(`${label}: Expected "${expected}", but found "${actual}"`);
          }
        };

        checkValue('B (CompanyID)', actualCompID, data.companyId);
        checkValue('F (CompanyName)', actualName, data.compName);
        checkValue('G (Country)', actualCountry, data.country);
        checkValue('H (Type)', actualType, data.type === '1' ? '本社' : '事業所');
        checkValue('I (CompanyNameJP)', actualNameJP, data.compName);
        checkValue('J (ShortName)', actualShortName, data.shortName);
        checkValue('K (Address)', actualAddress, data.address);
        checkValue('L (City)', actualCity, data.city);
        checkValue('N (Zip)', actualZip, data.zip);
        checkValue('O (Subsidiary)', actualSubsidiary, data.subsidiary === '1' ? 'MC単体' : '');
        checkValue('P (Office)', actualOffice, data.office);
        checkValue('Q (DTSSearch)', actualDTS, data.dtsSearch === 'Y' ? 'Yes' : 'No');
        checkValue('R (Judgement)', actualJudgement, '5-誤検知');
        checkValue('T (CompanyWide)', actualCompanyWide, data.companyWide === 'Y' ? 'Yes' : 'No');
        checkValue('U (Active)', actualActive, data.active === 'Y' ? 'はい' : 'いいえ');

        await test.info().attach(`edit-screen-verify-${data.companyId}`, { 
          body: await editPopup.screenshot({ fullPage: true }), 
          contentType: 'image/png' 
        });

        await editPopup.close();

        if (mismatches.length > 0) {
          console.log(`\n[!!!] Verification FAILED for CompanyID: ${data.companyId}`);
          mismatches.forEach(m => console.log(`      ${m}`));
          throw new Error(`❌ FAILED: Field verification failed for ${data.companyId} with ${mismatches.length} mismatches.`);
        }

        console.log(`   ✅ Verification for "${data.companyId}" complete.`);
      }

      try { fs.unlinkSync(testFilePath); fs.unlinkSync(baseTemplate); } catch(e) {}

      // Logout
      const userButtonName = `MC Test${user.id.substring(6)}`;
      await page.getByRole('button', { name: new RegExp(`^${userButtonName}`) }).click();
      await page.getByRole('button', { name: ' Sign out' }).click();
    });
  }
});
