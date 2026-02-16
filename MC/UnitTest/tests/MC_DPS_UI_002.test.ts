import { test, expect, Page } from '@playwright/test';
import { BASE_URL } from './config';
import * as fs from 'fs';
import * as path from 'path';

// Load user definitions
const allUsers = JSON.parse(fs.readFileSync(path.join(__dirname, 'account.dat'), 'utf8'));
const users = allUsers.filter((user: any) => user.id === 'MCTest9');

test.describe('MC_DPS_UI_002 Denied Party Screening - Hidden Fields Verification', () => {
  test.describe.configure({ mode: 'parallel' });
  
  users.forEach((user: any) => {
    
    test(`Perform Hidden Fields verification for ${user.id}`, async ({ page }, testInfo) => {
      test.setTimeout(600000); 

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
        await page.waitForTimeout(5000); // 遷移後のロード時間を少し長めに確保
      });

      // Step 3: Open an arbitrary record and verify hidden fields
      await test.step('Verify Hidden Fields in Edit Form', async () => {
        const mainFrameElement = page.locator('iframe[name="legacy-outlet"]');
        const frame = await mainFrameElement.contentFrame();
        if (!frame) throw new Error('Could not access main iframe content.');

        // Reset horizontal scroll
        const gridDataContainer = frame.locator('#ctl00_MainContent_rgLookup_GridData');
        if (await gridDataContainer.count() > 0) {
            await gridDataContainer.evaluate(el => { el.scrollLeft = 0; });
            await page.waitForTimeout(1000);
        }

        // Editリンクを探す (ヘッダーのソート用リンクを除外するため、行内を指定)
        const editLink = frame.locator('tr.rgRow, tr.rgAltRow').locator('a:has-text("Edit")').first();
        
        // 可視性待ち
        await expect(editLink).toBeVisible({ timeout: 30000 });
        
        // JSでスクロールして確実に中央へ
        await editLink.evaluate(el => {
            el.scrollIntoView({ behavior: 'instant', block: 'center', inline: 'center' });
        });
        await page.waitForTimeout(2000);

        // デバッグログ
        const linkHtml = await editLink.evaluate(el => el.outerHTML);
        console.log(`Targeting Edit link: ${linkHtml}`);

        await testInfo.attach('before_edit_click.png', {
            body: await page.screenshot({ fullPage: true }),
            contentType: 'image/png'
        });

        console.log('Attempting to open record...');
        
        let popup: Page;
        try {
            const popupPromise = page.context().waitForEvent('page', { timeout: 30000 });
            // Click with force: true
            await editLink.click({ force: true });
            popup = await popupPromise;
        } catch (e) {
            console.log(`Force click failed (${e.message}), trying JS click...`);
            try {
                const popupPromise = page.context().waitForEvent('page', { timeout: 30000 });
                await editLink.evaluate(el => el.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true, view: window })));
                popup = await popupPromise;
            } catch (e2) {
                console.log(`JS click failed (${e2.message}), trying Keyboard Enter...`);
                const popupPromise = page.context().waitForEvent('page', { timeout: 30000 });
                await editLink.focus();
                await page.keyboard.press('Enter');
                popup = await popupPromise;
            }
        }

        await popup.waitForLoadState('load');
        await popup.waitForTimeout(5000);

        const popupIframe = popup.frameLocator('iframe[name="legacy-outlet"]');
        const isPopupIframe = await popup.locator('iframe[name="legacy-outlet"]').count() > 0;
        const root = isPopupIframe ? popupIframe : popup;

        // フォームがロードされたか確認
        const companyIdIndicator = root.locator('#ctl00_MainContent_tabs_tabCompanyInfo_pnlBasic_Info_lblCompanyID, #txtCompanyID').first();
        await expect(companyIdIndicator).toBeVisible({ timeout: 30000 });

        // 確認対象のフィールド（表示されていないことを確認）
        const hiddenFields = [
            { name: 'CompanyType', selectors: ['#drpCompanyType', '#s2id_drpCompanyType', 'label:has-text("Company Type")', 'label:has-text("CompanyType")'] },
            { name: 'FederalID', selectors: ['#txtFederalID', 'label:has-text("Federal ID")', 'label:has-text("FederalID")'] },
            { name: 'DTSOverride', selectors: ['#drpDTSOverride', '#s2id_drpDTSOverride', 'label:has-text("DTS Override")', 'label:has-text("DTSOverride")'] },
            { name: 'LogicalDeletionFlag', selectors: ['#drpLogicalDeletionFlag', '#s2id_drpLogicalDeletionFlag', 'label:has-text("Logical Deletion Flag")', 'label:has-text("LogicalDeletionFlag")'] }
        ];

        console.log('Starting verification for hidden fields...');
        
        for (const field of hiddenFields) {
            console.log(`Checking for field: ${field.name}`);
            let isAnyVisible = false;
            for (const selector of field.selectors) {
                const el = root.locator(selector).first();
                if (await el.isVisible()) {
                    isAnyVisible = true;
                    console.error(`FAILED: Field "${field.name}" (found via selector "${selector}") is VISIBLE, but it should be HIDDEN.`);
                }
            }
            
            if (isAnyVisible) {
                await testInfo.attach(`error_visible_${field.name}.png`, {
                    body: await popup.screenshot({ fullPage: true }),
                    contentType: 'image/png'
                });
                throw new Error(`Field "${field.name}" is visible but should be hidden.`);
            }
            console.log(`Verified: Field "${field.name}" is NOT visible.`);
        }

        await testInfo.attach('final_form_verification.png', {
            body: await popup.screenshot({ fullPage: true }),
            contentType: 'image/png'
        });

        await popup.close();
      });

      // Step 4: Logout
      await test.step('Logout', async () => {
        const userButtonName = `MC Test${user.id.substring(6)}`;
        await page.bringToFront();

        const userButton = page.getByRole('button', { name: new RegExp(`^${userButtonName}`) });
        await userButton.hover();
        await page.waitForTimeout(500);
        await userButton.click();

        const logoutButton = page.locator('#logoutButton');
        await expect(logoutButton).toBeVisible({ timeout: 10000 });
        await logoutButton.click({ force: true });
      });
    });
  });
});
