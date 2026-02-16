import { test, expect } from '@playwright/test';
import { BASE_URL } from './config';
import * as fs from 'fs';
import * as path from 'path';

// Load user definitions
const users = JSON.parse(fs.readFileSync(path.join(__dirname, 'account.dat'), 'utf8'));

test.describe('TC_046 Role Based Access Control - Content Subscriptions Verification', () => {
  test.describe.configure({ mode: 'parallel' });
  
  users.forEach((user: any) => {
    
    test(`Perform Content Subscriptions verification for ${user.id}`, async ({ page }) => {
      test.setTimeout(180000); // Timeoutを延長

      console.log(`--- Test Start: ${user.id} ---`);

      // Step 1: Login
      await test.step('Login', async () => {
        await page.goto(BASE_URL + '/Logon.aspx?ReturnUrl=%2fgtm%2fhome');
        await page.getByRole('textbox', { name: 'Company' }).fill(user.company);
        await page.getByRole('textbox', { name: 'Username' }).fill(user.id);
        await page.getByRole('textbox', { name: 'Password' }).fill(user.pass);
        await page.getByRole('link', { name: 'Login' }).click();
      });

      // Step 2: Access Content Subscriptions Page
      const mainFrame = await test.step('Access Content Subscriptions Page', async () => {
        await page.getByRole('button', { name: 'Content' }).click();
        // "Content Subscriptions"メニューをクリック
        await page.locator('a[role="menuitem"]').filter({ hasText: /^Content Subscriptions$/ }).click();

        const mainFrameElement = page.locator('iframe[name="legacy-outlet"]');
        await mainFrameElement.waitFor({ state: 'attached', timeout: 40000 });
        const frame = await mainFrameElement.contentFrame();
        if (!frame) throw new Error('Could not access main iframe content.');

        const pageNameLabel = frame.locator('#lbxPageName');
        await expect(pageNameLabel).toHaveText('Content Subscriptions', { timeout: 30000 });
        
        return frame;
      });

      // Step 3: Verify and Capture Tabs
      await test.step('Verify and Capture Tabs', async () => {
        // タブヘッダーが表示されるのを待つ
        const tabHeader = mainFrame.locator('#ctl00_MainContent_tcnSubscriptions_header');
        await tabHeader.waitFor({ state: 'visible', timeout: 30000 });
        console.log('Tab header is visible.');

        // タブを動的に読み取ってスクリーンショットを撮る
        const tabs = tabHeader.locator('.ajax__tab_tab');
        const tabCount = await tabs.count();
        console.log(`Found ${tabCount} tabs.`);

        if (tabCount === 0) {
            console.log('No tabs found. Test will pass.');
            return;
        }

        for (let i = 0; i < tabCount; i++) {
            const tab = tabs.nth(i);
            const rawTabText = await tab.innerText();
            // タブ名から括弧と数字を削除
            const tabText = rawTabText.replace(/\s*\(\d+\)/, '').trim();
            console.log(`Processing tab #${i + 1}: ${rawTabText}`);

            await tab.click();
            
            // タブ名から対応するコンテンツパネルのIDサフィックスを決定
            let contentPanelIdSuffix = '';
            if (tabText === 'TARIFF') contentPanelIdSuffix = 'Tariff';
            else if (tabText === 'ANALYZER') contentPanelIdSuffix = 'Analyzer';
            else if (tabText === 'DENIED PARTY SCREENING') contentPanelIdSuffix = 'DPS';
            else if (tabText === 'ECN/ML') contentPanelIdSuffix = 'ECN';
            else if (tabText === 'FREE TRADE AGREEMENTS') contentPanelIdSuffix = 'FTA';
            else if (tabText === 'EXCHANGE RATES') contentPanelIdSuffix = 'ExchangeRate';
            else if (tabText === 'WCO NOTES') contentPanelIdSuffix = 'WCONotes';
            else if (tabText === 'LEGAL TEXT') contentPanelIdSuffix = 'LegalText';
            else if (tabText === 'CONTENT INTEGRATION') contentPanelIdSuffix = 'ContentIntegration';
            else if (tabText === 'BINDING RULINGS') contentPanelIdSuffix = 'BindingRulings';
            
            // 3つのパターンのいずれかが表示されるまで待機 (expect.or() を使用)
            if (contentPanelIdSuffix) {
                const contentPanel = mainFrame.locator(`#ctl00_MainContent_tcnSubscriptions_tab${contentPanelIdSuffix}_updpnl${contentPanelIdSuffix}`);
                
                const dataRow = contentPanel.locator('tr.rgRow').first();
                const noRecordsText = contentPanel.locator('span:has-text("There are no records.")').first();
                const notSubscribedText = contentPanel.locator('span:has-text("The current Partner is not subscribed")').first();

                try {
                    await expect(dataRow.or(noRecordsText).or(notSubscribedText)).toBeVisible({ timeout: 30000 });
                    console.log(`Content for tab "${tabText}" is loaded.`);
                } catch (e) {
                    console.warn(`WARN: Timed out waiting for content of tab "${tabText}". Taking screenshot anyway.`);
                }
            } else {
                // マッピングがない場合は、フォールバックとして固定時間待機
                console.warn(`WARN: No specific content panel selector found for tab: "${tabText}". Falling back to fixed wait.`);
                await page.waitForTimeout(4500);
            }

            // Content Integration タブの場合のみ、追加で3秒待機
            if (tabText === 'CONTENT INTEGRATION') {
                console.log('Applying special 3-second wait for Content Integration tab.');
                await page.waitForTimeout(3000);
            }

            const screenshot = await page.screenshot();
            // ファイル名に使えない文字を置換
            const safeTabText = rawTabText.replace(/[\/\s\(\)]/g, '_');
            await test.info().attach(`content_subscriptions_tab_${safeTabText}_${user.id}.png`, {
                body: screenshot,
                contentType: 'image/png'
            });
        }
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

        await expect(page.getByRole('textbox', { name: 'Company' })).toBeVisible({ timeout: 30000 });
      });
    });
  });
});
