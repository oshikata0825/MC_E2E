import { test, expect } from '@playwright/test';
import { BASE_URL } from './config';
import * as fs from 'fs';
import * as path from 'path';
import * as ExcelJS from 'exceljs';

/**
 * MC_DPS_Upload_008.test.ts
 * 
 * Tests spreadsheet upload and verifies the automatic setting of the SanFlag (Sanctioned Flag)
 * based on whether the Company Name starts with "San_".
 */

const allUsers = JSON.parse(fs.readFileSync(path.join(__dirname, 'account.dat'), 'utf8'));
const users = allUsers.filter(user => user.id === 'MCTest5' || user.id === 'MCTest9');

test.describe('MC_DPS_Upload_008 - SanFlag Automatic Setting Verification', () => {
  
  for (const user of users) {
    
    test(`Verify SanFlag for ${user.id}`, async ({ page }) => {
      test.setTimeout(600000); 
      
      console.log(`--- テスト開始: ${user.id} ---`);

      // ---------------------------------------------------
      // Step 1: Login
      // ---------------------------------------------------
      await test.step('Login', async () => {
        await page.goto(BASE_URL + '/Logon.aspx?ReturnUrl=%2fgtm%2fhome');
        await page.getByRole('textbox', { name: 'Company' }).fill(user.company);
        await page.getByRole('textbox', { name: 'Username' }).fill(user.id);
        await page.getByRole('textbox', { name: 'Password' }).fill(user.pass);
        await page.getByRole('link', { name: 'Login' }).click();
        console.log('✅ Login successful');
      });

      // ---------------------------------------------------
      // Step 2: Navigate to Spreadsheet Upload
      // ---------------------------------------------------
      await test.step('Navigate to Spreadsheet Upload', async () => {
        await page.getByRole('button', { name: 'Client Maintenance' }).click();
        await page.getByRole('menuitem', { name: 'Client Maintenance' }).click();
        
        const mainFrameElement = page.locator('iframe[name="legacy-outlet"]');
        await mainFrameElement.waitFor({ state: 'attached', timeout: 30000 });
        const mainFrame = await mainFrameElement.contentFrame();
        if (!mainFrame) throw new Error('Main frame not found');

        const uploadLink = mainFrame.locator('#ctl00_LeftHandNavigation_LHN_ctl00');
        await expect(uploadLink).toBeVisible({ timeout: 30000 });
        await uploadLink.click();
        
        const targetElement = mainFrame.locator('td', { hasText: '130840 SpreadSheet upload - DPS' });
        await expect(targetElement).toBeVisible({ timeout: 20000 });
        console.log('✅ Navigated to Spreadsheet Upload');
      });

      // ---------------------------------------------------
      // Step 3: Prepare Excel Data
      // ---------------------------------------------------
      const ts = new Date().toISOString().replace(/[:.-]/g, '').substring(0, 14);
      const fileName = `SAN_CHECK_008_${ts}.xlsx`;
      const testFilePath = path.join(__dirname, fileName);
      let excelData: { companyName: string, companyId: string, isSan: boolean }[] = [];

      await test.step('Prepare Excel Data', async () => {
        const mainFrame = page.frameLocator('iframe[name="legacy-outlet"]');
        const uploadActionLink = mainFrame.locator('td', { hasText: '130840 SpreadSheet upload - DPS' }).locator('..').locator('a', { hasText: /Upload/i });
        
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
        if (!tableFrame) throw new Error('Spreadsheet Upload popup frame not found');

        // Download Template
        let downloadBtn: any = null;
        for (const f of popup.frames()) {
          const loc = f.locator('#ctl00_MainContent_hyxlnDownload').or(f.getByText('Download Excel Template')).first();
          if (await loc.count() > 0) { downloadBtn = loc; break; }
        }
        if (!downloadBtn) downloadBtn = popup.locator('#ctl00_MainContent_hyxlnDownload').or(popup.getByText('Download Excel Template')).first();

        const [download] = await Promise.all([popup.waitForEvent('download', { timeout: 60000 }), downloadBtn.click({ force: true })]);
        const baseTemplate = path.join(__dirname, `base_template_${user.id}_008.xlsx`);
        await download.saveAs(baseTemplate);

        // Edit Excel
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(baseTemplate);
        const worksheet = workbook.getWorksheet(1)!;

        const genAlphaNum = (l: number) => Math.random().toString(36).substring(2, 2 + l).toUpperCase();
        const rowCount = Math.floor(Math.random() * 3) + 2; // 2 to 4 rows
        console.log(`ℹ️ Generating ${rowCount} rows for upload.`);

        for (let i = 0; i < rowCount; i++) {
          const r = i + 2;
          const companyId = `For-SIT-${genAlphaNum(8)}`;
          const isSan = (i % 2 === 1); // Alternating San_ and non-San_
          const companyName = (isSan ? 'San_' : '') + `Test-008-${genAlphaNum(6)}`;
          
          worksheet.getCell(`A${r}`).value = '130840';
          worksheet.getCell(`B${r}`).value = companyId;
          worksheet.getCell(`C${r}`).value = 'X';
          worksheet.getCell(`D${r}`).value = 'X1';
          worksheet.getCell(`E${r}`).value = '001';
          worksheet.getCell(`F${r}`).value = companyName;
          worksheet.getCell(`G${r}`).value = 'JP';
          worksheet.getCell(`Q${r}`).value = 'Y';
          worksheet.getCell(`U${r}`).value = 'Y';
          
          excelData.push({ companyName, companyId, isSan });
          console.log(`   - Row ${r}: ID=${companyId}, Name=${companyName}, ExpectedSan=${isSan}`);
        }
        await workbook.xlsx.writeFile(testFilePath);
        try { fs.unlinkSync(baseTemplate); } catch(e) {}
        
        // Upload
        let chooseFileBtn: any = null;
        let uploadBtn: any = null;
        for (const f of popup.frames()) {
          const cLoc = f.locator('#ctl00_MainContent_btxChooseFile').or(f.getByRole('button', { name: 'Choose File' }));
          if (await cLoc.count() > 0) chooseFileBtn = cLoc.first();
          const uLoc = f.locator('#ctl00_MainContent_lnxbtnUploadSpreadsheet').or(f.getByRole('link', { name: 'Upload and Import Data' }));
          if (await uLoc.count() > 0) uploadBtn = uLoc.first();
        }

        const fileChooserPromise = popup.waitForEvent('filechooser');
        await chooseFileBtn.click();
        const fileChooser = await fileChooserPromise;
        await fileChooser.setFiles(testFilePath);
        await uploadBtn.click();
        console.log(`ℹ️ Uploaded: ${fileName}`);

        // Monitor Status
        let refreshBtn: any = null;
        for (const f of popup.frames()) {
          const rLoc = f.locator('#ctl00_MainContent_lnxBtnRefreshWorkflowStatus').or(f.getByRole('link', { name: 'Refresh' }));
          if (await rLoc.count() > 0) { refreshBtn = rLoc.first(); break; }
        }

        let finalStatus = '';
        let targetRow: any = null;
        for (let i = 0; i < 40; i++) {
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

        if (!finalStatus.includes('Complete')) {
          if (finalStatus.includes('Error')) {
            const errorCountLink = targetRow.locator('a[datafield="ErrorCount"]');
            if (await errorCountLink.count() > 0) {
              await errorCountLink.click();
              await popup.waitForTimeout(8000);
              await test.info().attach(`upload-error-${user.id}`, { body: await popup.screenshot({ fullPage: true }), contentType: 'image/png' });
            }
          }
          throw new Error(`❌ Upload FAILED or timed out. Final Status: ${finalStatus}`);
        }
        await popup.close();
        console.log('✅ Spreadsheet upload complete');
      });

      // ---------------------------------------------------
      // Step 4: Verification in DPS Lookup
      // ---------------------------------------------------
      await test.step('Verification in DPS Lookup', async () => {
        await page.bringToFront();
        await page.getByRole('button', { name: 'DPS' }).click();
        await page.getByRole('menuitem', { name: 'DPS Search/Reporting Lookup' }).click();
        
        const legacyFrame = page.frameLocator('iframe[name="legacy-outlet"]');
        await legacyFrame.locator('a#ctl00_LeftHandNavigation_LHN_ctl00').click(); // MC Company Lookup
        await page.waitForTimeout(3000);

        const verificationErrors: string[] = [];

        for (const data of excelData) {
          console.log(`ℹ️ Verifying CompanyID: "${data.companyId}"`);
          try {
            // Filter by CompanyID
            await page.keyboard.press('Escape');
            const header = legacyFrame.locator('th.rgHeader').filter({ hasText: /^CompanyID$/i });
            await header.locator('button.rgActionButton.rgOptions').first().click({ force: true });
            await page.waitForTimeout(1000);

            const condInput = legacyFrame.locator('input[id*="HCFMRCMBFirstCond_Input"]');
            await condInput.click({ force: true });
            for(let j=0; j<10; j++) await page.keyboard.press('ArrowUp');
            for(let i=0; i<15; i++){
              const val = await condInput.inputValue();
              if(val?.toLowerCase() === 'equalto'){ await page.keyboard.press('Enter'); break; }
              await page.keyboard.press('ArrowDown');
              await page.waitForTimeout(200);
            }

            const valInput = legacyFrame.locator('input[id*="HCFMRTBFirstCond"]').first();
            await valInput.fill(data.companyId);
            await valInput.press('Enter');
            await page.waitForTimeout(5000);

            const row = legacyFrame.locator('tbody tr.rgRow, tbody tr.rgAltRow').first();
            if (!(await row.isVisible())) {
              const msg = `❌ [CompanyID: ${data.companyId}] Record NOT FOUND in lookup.`;
              console.log(msg);
              verificationErrors.push(msg);
              continue;
            }

            const editLink = row.locator('a', { hasText: 'Edit' });
            const [editPopup] = await Promise.all([
              page.waitForEvent('popup'),
              editLink.click(),
            ]);
            await editPopup.waitForLoadState();
            await editPopup.waitForTimeout(5000);

            // Extract values from Edit popup
            let actualSanFlag = '';
            let actualCompanyName = '';
            let found = false;
            let targetFrame: any = null;

            for (let i = 0; i < 15; i++) {
              for (const f of editPopup.frames()) {
                const sanFlagLoc = f.locator('#s2id_ctl00_MainContent_tabs_tabCompanyInfo_pnlSystem_drpSanFlag .select2-chosen').or(f.locator('#ctl00_MainContent_tabs_tabCompanyInfo_pnlSystem_lblSanFlag').locator('..').locator('.select2-chosen'));
                const nameLoc = f.locator('#txtCompanyName');
                if (await sanFlagLoc.count() > 0 && await nameLoc.count() > 0) {
                  actualSanFlag = (await sanFlagLoc.innerText()).trim();
                  actualCompanyName = (await nameLoc.inputValue()).trim();
                  targetFrame = f;
                  found = true; break;
                }
              }
              if (found) break;
              await editPopup.waitForTimeout(1000);
            }

            // Scroll to Company Name area (Top) and take screenshot
            if (found && targetFrame) {
              const nameInput = targetFrame.locator('#txtCompanyName').first();
              if (await nameInput.count() > 0) {
                await nameInput.scrollIntoViewIfNeeded();
                await editPopup.waitForTimeout(1000);
              }
            }
            await test.info().attach(`verify-${data.companyId}-name`, {
              body: await editPopup.screenshot(),
              contentType: 'image/png',
            });

            // Scroll to SanFlag area (Bottom) and take screenshot
            if (found && targetFrame) {
              const systemPanel = targetFrame.locator('div[id*="pnlSystem"], #ctl00_MainContent_tabs_tabCompanyInfo_pnlSystem_lblSanFlag').first();
              if (await systemPanel.count() > 0) {
                await systemPanel.scrollIntoViewIfNeeded();
                await editPopup.waitForTimeout(1000);
              }
            }
            await test.info().attach(`verify-${data.companyId}-sanflag`, {
              body: await editPopup.screenshot(),
              contentType: 'image/png',
            });

            if (!found) {
              const msg = `❌ [CompanyID: ${data.companyId}] Fields not found in Edit popup.`;
              console.log(msg);
              verificationErrors.push(msg);
              await editPopup.close();
              continue;
            }

            // Verify SanFlag
            const isSanMatch = (actual: string, expectedIsSan: boolean) => {
              if (expectedIsSan) {
                return actual.includes('はい') || actual.includes('Yes') || actual.includes('邵ｺｯ');
              } else {
                return actual.includes('いいえ') || actual.includes('No') || actual.includes('邵ｺ繝ｻ樒ｸｺ');
              }
            };

            const expectedSanLabel = data.isSan ? 'Yes/はい' : 'No/いいえ';
            const sanMatch = isSanMatch(actualSanFlag, data.isSan);
            console.log(`   [SanFlag] Match: ${sanMatch ? '✅' : '❌'}`);
            console.log(`     - Expected: "${expectedSanLabel}"`);
            console.log(`     - Actual:   "${actualSanFlag}"`);

            if (!sanMatch) {
              const msg = `❌ [CompanyID: ${data.companyId}] SanFlag MISMATCH: Expected=${expectedSanLabel}, Actual="${actualSanFlag}"`;
              verificationErrors.push(msg);
            }

            // Verify CompanyName
            // New Specification: CompanyName should RETAIN the "San_" prefix in the UI.
            const expectedCompanyName = data.companyName;
            const nameMatch = (actualCompanyName === expectedCompanyName);
            console.log(`   [CompanyName] Match: ${nameMatch ? '✅' : '❌'}`);
            console.log(`     - Expected: "${expectedCompanyName}"`);
            console.log(`     - Actual:   "${actualCompanyName}"`);

            if (!nameMatch) {
              const msg = `❌ [CompanyID: ${data.companyId}] CompanyName MISMATCH: Expected="${expectedCompanyName}", Actual="${actualCompanyName}"`;
              verificationErrors.push(msg);
            }

            await editPopup.close();

          } catch (e: any) {
            const msg = `❌ [CompanyID: ${data.companyId}] Exception during verification: ${e.message}`;
            console.log(msg);
            verificationErrors.push(msg);
          }
        }

        if (verificationErrors.length > 0) {
          // Store errors to throw after logout
          (user as any).verificationErrors = verificationErrors;
        } else {
          console.log('✅ All records verified successfully');
        }
      });

      // ---------------------------------------------------
      // Step 5: Logout
      // ---------------------------------------------------
      await test.step('Logout', async () => {
        try {
          const userButtonName = `MC Test${user.id.substring(6)}`;
          await page.getByRole('button', { name: new RegExp(`^${userButtonName}`) }).click();
          await page.getByRole('button', { name: ' Sign out' }).click();
          await page.waitForURL(/Logon|default/i);
          console.log('✅ Logout successful');
        } catch (logoutErr) {
          console.warn('⚠️ Logout warning:', logoutErr.message);
        }

        try { fs.unlinkSync(testFilePath); } catch(e) {}

        // Finally throw if there were verification errors
        const errors = (user as any).verificationErrors;
        if (errors && errors.length > 0) {
          console.log('\n--- [Verification Failures Summary] ---');
          errors.forEach((e: string) => console.log(e));
          throw new Error(`Verification failed with ${errors.length} errors. Check console for details.`);
        }
      });
    });
  }
});
