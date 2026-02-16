import { test, expect } from '@playwright/test';
import { BASE_URL } from './config';
import * as fs from 'fs';
import * as path from 'path';

// Load user definitions
const allUsers = JSON.parse(fs.readFileSync(path.join(__dirname, 'account.dat'), 'utf8'));
const users = allUsers.filter((user: any) => user.id === 'MCTest1');

test.describe('MC_DPS_UI_001 Denied Party Screening - UI Verification', () => {
  test.describe.configure({ mode: 'parallel' });
  
  users.forEach((user: any) => {
    
    test(`Perform UI verification for ${user.id}`, async ({ page }) => {
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

      // Step 2: Access Create New Company Page
      const mainFrame = await test.step('Access Create New Company Page', async () => {
        await page.getByRole('button', { name: 'DPS' }).click();
        // "Create New Company"メニューをクリック (大文字小文字を区別しない正規表現を使用)
        await page.locator('a[role="menuitem"]').filter({ hasText: /^Create New Company$/i }).click();

        const mainFrameElement = page.locator('iframe[name="legacy-outlet"]');
        await mainFrameElement.waitFor({ state: 'attached', timeout: 40000 });
        const frame = await mainFrameElement.contentFrame();
        if (!frame) throw new Error('Could not access main iframe content.');

        // tests/elements.txt の内容に基づき、Questionnaireのタイトルが表示されるのを待つ
        const questionnaireTitle = frame.locator('#ctl00_MainContent_lbxQuestionnaireTitle');
        // タイトル表示を待つ (前のステップで追加済みだが、文脈維持のため)
        // await expect(questionnaireTitle).toContainText('Questions', { timeout: 30000 });
        
        // Step 3: Input Customer Name and Submit
        await test.step('Input Customer Name and Submit', async () => {
            const randomSuffix = Math.random().toString(36).substring(2, 8).toUpperCase();
            const randomCustomerName = `Global Trade Solutions ${randomSuffix} Corp`;
            console.log(`Input - Customer Name: ${randomCustomerName}`);

            const nameInput = frame.locator('#ctl00_MainContent_lstQuestionnaire_ctrl0_txtTextResponse');
            await nameInput.fill(randomCustomerName);
            
            const submitButton = frame.locator('#ctl00_MainContent_lnxbtnSingleSubmit');
            await submitButton.click();
            await page.waitForTimeout(3000);

            // Step 4: Select Group and Submit again
            await test.step('Select Group and Submit again', async () => {
                const groupSelect = frame.locator('.select2-choice').first();
                await groupSelect.click();
                await page.waitForTimeout(1000);
                
                await page.keyboard.press('ArrowDown');
                await page.keyboard.press('Enter');
                await page.waitForTimeout(1000);

                const selectedGroup = await groupSelect.locator('.select2-chosen').innerText();
                console.log(`Select - Group: ${selectedGroup}`);

                await submitButton.click();
                await page.waitForTimeout(3000);
            });

            // Step 5: Select Department and Submit again
            await test.step('Select Department and Submit again', async () => {
                const departmentSelect = frame.locator('.select2-choice').nth(1);
                await departmentSelect.click();
                await page.waitForTimeout(1000);
                
                await page.keyboard.press('ArrowDown');
                await page.keyboard.press('Enter');
                await page.waitForTimeout(1000);

                const selectedDept = await departmentSelect.locator('.select2-chosen').innerText();
                console.log(`Select - Department: ${selectedDept}`);

                await submitButton.click();
                await page.waitForTimeout(3000);
            });

            // Step 6: Select Section and Submit again
            await test.step('Select Section and Submit again', async () => {
                const sectionSelect = frame.locator('.select2-choice').nth(2);
                await sectionSelect.click();
                await page.waitForTimeout(1000);
                
                await page.keyboard.press('ArrowDown');
                await page.keyboard.press('Enter');
                await page.waitForTimeout(1000);

                const selectedSection = await sectionSelect.locator('.select2-chosen').innerText();
                console.log(`Select - Section: ${selectedSection}`);

                await submitButton.click();
                await page.waitForTimeout(3000);
            });

            // Step 7: Select Country (Randomly) and Submit again
            await test.step('Select Country and Submit again', async () => {
                const countrySelect = frame.locator('.select2-choice').nth(3);
                await countrySelect.click();
                await page.waitForTimeout(1000);
                
                const randomDowns = Math.floor(Math.random() * 8) + 3;
                for (let i = 0; i < randomDowns; i++) {
                    await page.keyboard.press('ArrowDown');
                    await page.waitForTimeout(100);
                }
                await page.keyboard.press('Enter');
                await page.waitForTimeout(1000);

                const selectedCountry = await countrySelect.locator('.select2-chosen').innerText();
                console.log(`Select - Country (random): ${selectedCountry}`);

                await submitButton.click();
                await page.waitForTimeout(3000);
            });

            // Step 8: Input Address and Submit
            await test.step('Input Address and Submit', async () => {
                const addressInput = frame.locator('#ctl00_MainContent_lstQuestionnaire_ctrl5_txtTextResponse');
                const randomAddress = `${Math.floor(Math.random() * 999) + 1} Business Road, District ${Math.floor(Math.random() * 50) + 1}`;
                await addressInput.fill(randomAddress);
                console.log(`Input - Address: ${randomAddress}`);

                const submitButton = frame.locator('#ctl00_MainContent_lnxbtnSingleSubmit');
                await submitButton.click();
                await page.waitForTimeout(3000);
            });

            // Step 9: Confirm Completion and Submit
            await test.step('Confirm Completion and Submit', async () => {
                const completedMessage = frame.locator('.MessageText');
                await expect(completedMessage).toHaveText('No more questions, questionnaire completed.', { timeout: 30000 });

                const submitButton = frame.locator('#ctl00_MainContent_lnxbtnSingleSubmit');
                await submitButton.click();
                await page.waitForTimeout(3000);

                // 画面読み込みを待機 (elements.txtの1〜3行目の要素が表示されるのを待つ)
                await expect(frame.locator('#ctl00_MainContent_lstQuestionnaire_ctrl0_txtTextResponse')).toBeVisible({ timeout: 10000 });
                await expect(frame.locator('#ctl00_MainContent_lnxbtnSingleSubmit')).toBeVisible({ timeout: 10000 });
                await expect(frame.locator('.select2-choice').first()).toBeVisible({ timeout: 10000 });
                
                // さらに3秒待機
                await page.waitForTimeout(3000);

                // Step 9終了後、Step 10開始前のスクリーンショット
                await test.info().attach('after_address_submit.png', {
                    body: await page.screenshot(),
                    contentType: 'image/png'
                });
            });

            // Step 10: Go To Maintenance Page
            await test.step('Go To Maintenance Page', async () => {
                const maintenanceLink = frame.locator('a:has-text("Go To Maintenance Page")');
                await expect(maintenanceLink).toBeVisible({ timeout: 30000 });

                // 新しいタブで開くリンクをクリック
                const [newPage] = await Promise.all([
                    page.waitForEvent('popup'),
                    maintenanceLink.click(),
                ]);

                await newPage.waitForLoadState();
                
                // メンテナンスページも恐らくiframe構成のため、legacy-outletを探す
                const maintenanceFrame = newPage.frameLocator('iframe[name="legacy-outlet"]');
                
                // メンテナンスページの要素が表示されるのを待つ (elements.txtの内容に基づく)
                // iframe内または直接のいずれかで表示されるのを待つ
                const tabHeader = maintenanceFrame.locator('#ctl00_MainContent_tabs_header');
                const companyInfoTab = maintenanceFrame.locator('#__tab_ctl00_MainContent_tabs_tabCompanyInfo');

                try {
                    await expect(tabHeader.or(newPage.locator('#ctl00_MainContent_tabs_header'))).toBeVisible({ timeout: 30000 });
                    console.log('Maintenance page header detected.');
                } catch (e) {
                    console.warn('WARN: Could not find tab header via standard or iframe locator. Taking screenshot anyway.');
                }

                await newPage.waitForTimeout(3000);

                // 最後のスクリーンショット
                await test.info().attach('maintenance_page_opened.png', {
                    body: await newPage.screenshot(),
                    contentType: 'image/png'
                });

                // Step 11: Validation on Maintenance Page
                await test.step('Validation on Maintenance Page', async () => {
                    const saveButton = maintenanceFrame.locator('#ctl00_MainContent_lnxbtnSave');
                    
                    const validationFields = [
                        { name: 'Company Name', locator: maintenanceFrame.locator('#txtCompanyName'), type: 'input' },
                        { name: 'Address 1', locator: maintenanceFrame.locator('#txtCompanyAddress1'), type: 'input' },
                        { name: 'Country', id: 'ctl00_MainContent_tabs_tabCompanyInfo_pnlBasic_Info_drpCompanyCountryCode', type: 'select2' },
                        { name: 'Active Flag', id: 'ctl00_MainContent_tabs_tabCompanyInfo_pnlAudit_drpActiveFlag', type: 'select2' }
                    ];

                    for (const field of validationFields) {
                        console.log(`--- Validating mandatory field: ${field.name} ---`);
                        
                        let originalValue = '';
                        if (field.type === 'input') {
                            const inputLocator = field.locator!;
                            originalValue = await inputLocator.inputValue();
                            await inputLocator.clear();
                        } else {
                            // select2 の場合
                            const selectEl = maintenanceFrame.locator(`#${field.id}`);
                            const container = maintenanceFrame.locator(`#s2id_${field.id}`);
                            
                            // 現在の表示テキストを保存
                            originalValue = await container.locator('.select2-chosen').innerText();
                            console.log(`Current value of ${field.name}: ${originalValue}`);
                            
                            // セレクトボックスの値を空にする (force: true で非表示要素を操作)
                            await selectEl.selectOption({ value: '' }, { force: true });
                            // select2に反映させるために change イベントを発火
                            await selectEl.evaluate(el => el.dispatchEvent(new Event('change')));
                        }
                        await page.waitForTimeout(1000);
                        console.log(`Clear action performed for ${field.name}`);

                        // Saveをクリック
                        console.log(`Clicking Save for ${field.name}`);
                        await saveButton.click({ force: true, timeout: 10000 });
                        
                        // エラーメッセージが表示されたことを確認
                        const errorMessage = maintenanceFrame.locator('.MessageText');
                        try {
                            // "Please complete..." というメッセージが表示されるまで待つ
                            await expect(errorMessage).toHaveText(/Please complete the required fields before saving/i, { timeout: 15000 });
                            
                            const errorText = await errorMessage.innerText();
                            console.log(`Detected Error Text: "${errorText}"`);
                            
                            // スクリーンショットの前に該当項目までスクロール
                            if (field.type === 'input') {
                                await field.locator!.scrollIntoViewIfNeeded();
                            } else {
                                await maintenanceFrame.locator(`#s2id_${field.id}`).scrollIntoViewIfNeeded();
                            }

                            // スクリーンショットの前に少し待機して描画を確実にする
                            await page.waitForTimeout(3000);

                            await test.info().attach(`validation_error_${field.name.replace(/\s/g, '_')}.png`, {
                                body: await newPage.screenshot(),
                                contentType: 'image/png'
                            });
                        } catch (e) {
                            console.error(`Validation check failed for ${field.name}: ${e.message}`);
                            await test.info().attach(`failed_validation_${field.name.replace(/\s/g, '_')}.png`, {
                                body: await newPage.screenshot(),
                                contentType: 'image/png'
                            });
                            throw e;
                        }

                        // 値を戻す
                        if (field.type === 'input') {
                            await field.locator!.fill(originalValue);
                        } else {
                            const container = maintenanceFrame.locator(`#s2id_${field.id}`);
                            await container.click();
                            await page.waitForTimeout(500);
                            await page.keyboard.type(originalValue);
                            await page.waitForTimeout(500);
                            await page.keyboard.press('Enter');
                        }
                        await page.waitForTimeout(1500);
                        console.log(`Restored original value for ${field.name}: ${originalValue}`);
                    }
                });
            });
        });

        return frame;
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
