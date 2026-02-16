import { test, expect, Page } from '@playwright/test';
import { BASE_URL } from './config';
import * as fs from 'fs';
import * as path from 'path';

// Load user definitions
const allUsers = JSON.parse(fs.readFileSync(path.join(__dirname, 'account.dat'), 'utf8'));
const users = allUsers.filter((user: any) => user.id === 'MCTest9');

/**
 * Common logic to verify base fields are read-only/disabled.
 */
async function verifyBaseFieldsReadOnly(popup: Page, type: string, testInfo: any) {
  console.log(`Verifying fields are read-only for ${type}...`);
  
  const iframe = popup.frameLocator('iframe[name="legacy-outlet"]');
  const isIframe = await popup.locator('iframe[name="legacy-outlet"]').count() > 0;
  const root = isIframe ? iframe : popup;

  // Base Input elements to check for all types
  const readOnlyInputs = [
    { name: 'CompanyID', selector: '#txtCompanyID' },
    { name: 'DTSStatus', selector: '#txtDTSStatus' },
    { name: 'CreatedDate', selector: '#ctl00_MainContent_tabs_tabCompanyInfo_pnlAudit_calCreatedDate_dateInput' }
  ];
  
  for (const input of readOnlyInputs) {
    const el = root.locator(input.selector);
    if (await el.count() === 0) continue;

    // スクロールして撮影
    await el.scrollIntoViewIfNeeded();
    await popup.waitForTimeout(500);
    await testInfo.attach(`readonly_${type}_${input.name}.png`, {
      body: await popup.screenshot(),
      contentType: 'image/png'
    });

    const isReadOnly = await el.getAttribute('readonly') === 'readonly' || await el.getAttribute('aria-readonly') === 'true';
    const hasDisabledClass = (await el.getAttribute('class'))?.includes('riDisabled');
    
    if (!isReadOnly && !hasDisabledClass) {
      const msg = `\n[!!!] FAILED: Field "${input.name}" (${input.selector}) is expected to be read-only for ${type} but is EDITABLE.`;
      console.log(msg);
      throw new Error(msg);
    }
    console.log(`Verified: Field "${input.name}" is read-only.`);
  }

  // Verify Select2 containers (DataSource and SanFlag)
  const select2Containers = [
    { name: 'DataSource', selector: '#s2id_ctl00_MainContent_tabs_tabCompanyInfo_pnlBasic_Info_drpDataSource' },
    { name: 'SanFlag', selector: '#s2id_ctl00_MainContent_tabs_tabCompanyInfo_pnlSystem_drpSanFlag' }
  ];

  for (const container of select2Containers) {
    const el = root.locator(container.selector);
    if (await el.count() > 0) {
      await el.scrollIntoViewIfNeeded();
      await popup.waitForTimeout(500);
      await testInfo.attach(`readonly_${type}_${container.name}.png`, {
        body: await popup.screenshot(),
        contentType: 'image/png'
      });

      const containerClass = await el.getAttribute('class');
      if (!containerClass?.includes('select2-container-disabled')) {
        const msg = `\n[!!!] FAILED: Select2 "${container.name}" (${container.selector}) is expected to be disabled for ${type} but is ENABLED.`;
        console.log(msg);
        throw new Error(msg);
      }
      console.log(`Verified: ${container.name} select2 is disabled.`);
    }
  }
}

/**
 * Helper to verify a single field is read-only.
 */
async function verifySpecificFieldReadOnly(popup: Page, root: any, selector: string, fieldName: string, type: string, testInfo: any) {
  const el = root.locator(selector);
  if (await el.count() === 0) return;

  await el.scrollIntoViewIfNeeded();
  await popup.waitForTimeout(500);
  await testInfo.attach(`readonly_${type}_${fieldName}.png`, {
    body: await popup.screenshot(),
    contentType: 'image/png'
  });

  const isReadOnly = await el.getAttribute('readonly') === 'readonly' || await el.getAttribute('aria-readonly') === 'true';
  const hasDisabledClass = (await el.getAttribute('class'))?.includes('riDisabled');
  
  if (!isReadOnly && !hasDisabledClass) {
    const msg = `\n[!!!] FAILED: Specific field "${fieldName}" (${selector}) is expected to be read-only for ${type} but is EDITABLE.`;
    console.log(msg);
    throw new Error(msg);
  }
  console.log(`Verified: Specific field "${fieldName}" is read-only for ${type}.`);
}

async function verifyFieldsReadOnlyManual(popup: Page, testInfo: any) {
  await verifyBaseFieldsReadOnly(popup, 'Manual', testInfo);
}

async function verifyFieldsReadOnlySS(popup: Page, testInfo: any) {
  await verifyBaseFieldsReadOnly(popup, 'SS', testInfo);
}

async function verifyFieldsReadOnlyMDM(popup: Page, testInfo: any) {
  await verifyBaseFieldsReadOnly(popup, 'MDM', testInfo);
}

test.describe('MC_DPS_UI_003 Denied Party Screening - MC Company Lookup Verification', () => {
  test.describe.configure({ mode: 'parallel' });
  
  users.forEach((user: any) => {
    
    test(`Perform UI verification for ${user.id}`, async ({ page }, testInfo) => {
      test.setTimeout(600000); // タイムアウトをさらに延長

      console.log(`--- Test Start: ${user.id} ---`);

      // Step 1: Login
      await test.step('Login', async () => {
        await page.goto(BASE_URL + '/Logon.aspx?ReturnUrl=%2fgtm%2fhome');
        await page.getByRole('textbox', { name: 'Company' }).fill(user.company);
        await page.getByRole('textbox', { name: 'Username' }).fill(user.id);
        await page.getByRole('textbox', { name: 'Password' }).fill(user.pass);
        await page.getByRole('link', { name: 'Login' }).click();
      });

      // Step 2: Access MC Company Lookup Page
      await test.step('Access MC Company Lookup Page', async () => {
        await page.getByRole('button', { name: 'DPS' }).click();
        await page.locator('a[role="menuitem"]').filter({ hasText: /^DPS Search\/Reporting Lookup$/i }).click();

        const mainFrameElement = page.locator('iframe[name="legacy-outlet"]');
        await mainFrameElement.waitFor({ state: 'attached', timeout: 40000 });
        const frame = await mainFrameElement.contentFrame();
        if (!frame) throw new Error('Could not access main iframe content.');

        await frame.getByRole('link', { name: 'MC Company Lookup' }).click();
        await page.waitForTimeout(3000);
      });

      // Step 3: Iterate Edit Links and Filter MDM
      await test.step('Verify DataSources', async () => {
        const mainFrameElement = page.locator('iframe[name="legacy-outlet"]');
        const frame = await mainFrameElement.contentFrame();
        if (!frame) throw new Error('Could not access main iframe content.');

        const verifiedDataSources = new Set<string>();

        // 1. Manual と SS を1つずつ確認
        let editLinkLocator = frame.locator('a:has-text("Edit")');
        let count = await editLinkLocator.count();
        console.log(`Initial search: Found ${count} links.`);

        for (let i = 0; i < Math.min(count, 30); i++) {
          if (verifiedDataSources.has('Manual') && verifiedDataSources.has('SS')) break;

          console.log(`Checking record ${i + 1} for Manual/SS...`);
          const link = editLinkLocator.nth(i);
          await link.scrollIntoViewIfNeeded();

          let popup: Page;
          try {
            const popupPromise = page.context().waitForEvent('page', { timeout: 15000 });
            await link.click({ force: true });
            popup = await popupPromise;
          } catch (e) {
            try {
              const popupPromise = page.context().waitForEvent('page', { timeout: 15000 });
              await link.evaluate(el => el.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true, view: window })));
              popup = await popupPromise;
            } catch (e2) {
              continue;
            }
          }

          await popup.waitForLoadState('load');
          await popup.waitForTimeout(5000); 

          const popupIframe = popup.frameLocator('iframe[name="legacy-outlet"]');
          const isPopupIframe = await popup.locator('iframe[name="legacy-outlet"]').count() > 0;
          const popupRoot = isPopupIframe ? popupIframe : popup;

          const companyIdIndicator = popupRoot.locator('#ctl00_MainContent_tabs_tabCompanyInfo_pnlBasic_Info_lblCompanyID, #txtCompanyID').first();
          await expect(companyIdIndicator).toBeVisible({ timeout: 20000 });

          const dataSourceChosen = popupRoot.locator('#s2id_ctl00_MainContent_tabs_tabCompanyInfo_pnlBasic_Info_drpDataSource .select2-chosen');
          const dataSourceValue = (await dataSourceChosen.innerText()).trim();
          console.log(`DataSource for entry ${i + 1}: ${dataSourceValue}`);

          if (dataSourceValue === 'Manual' && !verifiedDataSources.has('Manual')) {
            await verifyFieldsReadOnlyManual(popup, testInfo);
            verifiedDataSources.add('Manual');
          } else if (dataSourceValue === 'SS' && !verifiedDataSources.has('SS')) {
            await verifyFieldsReadOnlySS(popup, testInfo);
            verifiedDataSources.add('SS');
          } else if (dataSourceValue === 'MDM' && !verifiedDataSources.has('MDM')) {
            await verifyFieldsReadOnlyMDM(popup, testInfo);
            verifiedDataSources.add('MDM');
          }
          
          await popup.close();
        }

        // 2. MDM がまだ見つかっていなければ、フィルタリングして確認
        if (!verifiedDataSources.has('MDM')) {
          console.log('MDM record not found in initial loop. Filtering for CreatedBy=MDM to find MDM record...');
          
          const headers = await frame.locator('th.rgHeader').allInnerTexts();
          console.log('Available columns:', headers.map(h => h.trim()).join(', '));

          const targetHeader = frame.locator('th.rgHeader').filter({ hasText: /^CreatedBy$/i });
          await targetHeader.locator('button.rgActionButton.rgOptions').first().click();
          await page.waitForTimeout(2000);

          const condInput = frame.locator('input[id*="HCFMRCMBFirstCond_Input"]');
          await condInput.click();
          await page.waitForTimeout(1000);
          
          await frame.locator('li.rcbItem', { hasText: /^EqualTo$/i }).or(page.locator('li.rcbItem', { hasText: /^EqualTo$/i })).first().click();
          await page.waitForTimeout(1000);

          const valInput = frame.locator('input[id*="HCFMRTBFirstCond"]').first();
          await valInput.click();
          await valInput.fill('');
          await valInput.pressSequentially('MDM', { delay: 100 });
          await valInput.press('Tab');
          await page.waitForTimeout(1000);

          const filterBtn = frame.locator('input[type="submit"][value="Filter"], button:has-text("Filter"), .rgFilterButton, span.rgButtonText:has-text("Filter")').first();
          await filterBtn.click({ force: true });
          
          console.log('Waiting for filter to be applied (5s)...');
          await page.waitForTimeout(5000);

          // デバッグ：フィルタ適用後の状態を撮影
          await testInfo.attach('after_filter_applied.png', {
            body: await page.screenshot({ fullPage: true }),
            contentType: 'image/png'
          });

          // フィルターメニューを閉じるために Escape キーを送信
          await page.keyboard.press('Escape');
          await page.waitForTimeout(1000);

          // スクロールリセット
          const gridDataContainer = frame.locator('#ctl00_MainContent_rgLookup_GridData');
          if (await gridDataContainer.count() > 0) {
            console.log('Resetting horizontal scroll for GridData...');
            await gridDataContainer.evaluate(el => { el.scrollLeft = 0; });
            await page.waitForTimeout(1000);
          }

          // フィルタリング後の MDM 行を特定して Edit リンクを取得
          const mdmRows = frame.locator('tr.rgRow, tr.rgAltRow').filter({ hasText: 'MDM' });
          const mdmCount = await mdmRows.count();
          console.log(`After filtering: Found ${mdmCount} rows containing "MDM".`);

          if (mdmCount > 0) {
            const firstMdmRow = mdmRows.first();
            const link = firstMdmRow.locator('a:has-text("Edit")');
            
            console.log('Targeting MDM row Edit link. Bringing into view...');
            await link.evaluate(el => {
              el.scrollIntoView({ behavior: 'instant', block: 'center', inline: 'center' });
            });
            await page.waitForTimeout(1000);

            // デバッグ：クリック直前の状態を撮影
            await testInfo.attach('before_mdm_click.png', {
              body: await page.screenshot({ fullPage: true }),
              contentType: 'image/png'
            });

            let popup: Page;
            console.log('Attempting to open MDM record using Enter key...');
            try {
              const popupPromise = page.context().waitForEvent('page', { timeout: 30000 });
              await link.focus();
              await page.keyboard.press('Enter');
              popup = await popupPromise;
            } catch (e) {
              console.log(`Enter key failed (${e.message}), trying force click...`);
              try {
                const popupPromise = page.context().waitForEvent('page', { timeout: 30000 });
                await link.click({ force: true, timeout: 10000 });
                popup = await popupPromise;
              } catch (e2) {
                console.log(`Force click failed (${e2.message}), trying JS click...`);
                try {
                  const popupPromise = page.context().waitForEvent('page', { timeout: 30000 });
                  await link.evaluate(el => el.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true, view: window })));
                  popup = await popupPromise;
                } catch (e3) {
                  throw new Error(`Failed to open MDM record popup after all attempts: ${e3.message}`);
                }
              }
            }

            await popup.waitForLoadState('load');
            await popup.waitForTimeout(5000); 

            const popupIframe = popup.frameLocator('iframe[name="legacy-outlet"]');
            const isPopupIframe = await popup.locator('iframe[name="legacy-outlet"]').count() > 0;
            const popupRoot = isPopupIframe ? popupIframe : popup;

            const companyIdIndicator = popupRoot.locator('#ctl00_MainContent_tabs_tabCompanyInfo_pnlBasic_Info_lblCompanyID, #txtCompanyID').first();
            await expect(companyIdIndicator).toBeVisible({ timeout: 20000 });

            await verifyFieldsReadOnlyMDM(popup, testInfo);
            verifiedDataSources.add('MDM');
            await popup.close();
          } else {
            console.warn('No MDM records found even after filtering.');
          }
        }

        console.log(`Test completed. Verified DataSources: ${Array.from(verifiedDataSources).join(', ')}`);
      });

      // Step 4: Logout
      await test.step('Logout', async () => {
        const userButtonName = "MC Test9";
        await page.bringToFront();
        const userButton = page.getByRole('button', { name: new RegExp(`^${userButtonName}`) });
        await userButton.hover();
        await page.waitForTimeout(500);
        await userButton.click();
        const logoutButton = page.locator('#logoutButton');
        await expect(logoutButton).toBeVisible({ timeout: 10000 });
        await logoutButton.click({ force: true });
        await expect(page.getByRole('textbox', { name: 'Company' })).toBeVisible({ timeout: 30000 });
      });
    });
  });
});