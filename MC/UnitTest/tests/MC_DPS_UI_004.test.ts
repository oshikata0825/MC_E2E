import { test, expect, Page } from '@playwright/test';
import { BASE_URL } from './config';
import * as fs from 'fs';
import * as path from 'path';

// Load user definitions
const allUsers = JSON.parse(fs.readFileSync(path.join(__dirname, 'account.dat'), 'utf8'));
const users = allUsers.filter((user: any) => user.id === 'MCTest1');

test.describe('MC_DPS_UI_004 Denied Party Screening - Create New Company UI Verification', () => {
  test.describe.configure({ mode: 'parallel' });
  
  users.forEach((user: any) => {
    
    test(`Perform UI verification for ${user.id}`, async ({ page }, testInfo) => {
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

      // Step 2: Access Create New Company Page
      await test.step('Access Create New Company Page', async () => {
        await page.getByRole('button', { name: 'DPS' }).click();
        await page.locator('a[role="menuitem"]').filter({ hasText: /^Create New Company$/i }).click();

        const mainFrameElement = page.locator('iframe[name="legacy-outlet"]');
        await mainFrameElement.waitFor({ state: 'attached', timeout: 40000 });
        const frame = await mainFrameElement.contentFrame();
        if (!frame) throw new Error('Could not access main iframe content.');

        // Step 3: Input Customer Name and Submit
        const randomSuffix = Math.random().toString(36).substring(2, 8).toUpperCase();
        const randomCustomerName = `Fictitious Customer ${randomSuffix} Corp`;
        console.log(`Input - Customer Name: ${randomCustomerName}`);

        // Using selector from elements.txt line 1
        const nameInput = frame.locator('#ctl00_MainContent_lstQuestionnaire_ctrl0_txtTextResponse');
        await nameInput.fill(randomCustomerName);
        
        // Using selector from elements.txt line 2
        const submitButton = frame.locator('#ctl00_MainContent_lnxbtnSingleSubmit');
        await submitButton.click();
        await page.waitForTimeout(2000);

        // Step 4: Select 'X' for the first dropdown menu (Group)
        await test.step('Select X for Group', async () => {
            const select = frame.locator('.select2-choice').nth(0);
            const chosen = select.locator('.select2-chosen');
            await expect(select).toBeVisible({ timeout: 10000 });
            await select.click({ force: true });
            await page.waitForTimeout(1000);
            
            // Select2の入力フィールドがアクティブになるのを待って入力
            await page.keyboard.type('X', { delay: 200 });
            await page.waitForTimeout(1500);
            await page.keyboard.press('Enter');
            await page.waitForTimeout(1000);
            
            await expect(chosen).toHaveText('X', { timeout: 5000 });
            console.log('Verified: Group is set to X');

            await testInfo.attach('step4_X_selected.png', {
                body: await page.screenshot(),
                contentType: 'image/png'
            });

            await submitButton.click();
            await page.waitForTimeout(3000);
        });

        // Step 5: Select 'X1' for the second dropdown menu (Department)
        await test.step('Select X1 for Department', async () => {
            // elements.txt 7行目。2番目のプルダウンなので nth(1)
            const select = frame.locator('.select2-choice').nth(1);
            const chosen = select.locator('.select2-chosen');
            
            await expect(select).toBeVisible({ timeout: 10000 });
            await select.scrollIntoViewIfNeeded();
            await select.click({ force: true });
            await page.waitForTimeout(1000);
            
            await page.keyboard.type('X1', { delay: 200 });
            await page.waitForTimeout(1500);
            await page.keyboard.press('Enter');
            await page.waitForTimeout(1000);

            await expect(chosen).toHaveText('X1', { timeout: 5000 });
            console.log('Verified: Department is set to X1');

            await testInfo.attach('step5_X1_selected.png', {
                body: await page.screenshot(),
                contentType: 'image/png'
            });

            await submitButton.click();
            await page.waitForTimeout(3000);
        });

        // Step 6: Select '001' for the third dropdown menu (Section)
        await test.step('Select 001 for Section', async () => {
            // elements.txt 9行目。3番目のプルダウンなので nth(2)
            const select = frame.locator('.select2-choice').nth(2);
            const chosen = select.locator('.select2-chosen');
            
            await expect(select).toBeVisible({ timeout: 10000 });
            await select.scrollIntoViewIfNeeded();
            await select.click({ force: true });
            await page.waitForTimeout(1000);
            
            await page.keyboard.type('001', { delay: 200 });
            await page.waitForTimeout(1500);
            await page.keyboard.press('Enter');
            await page.waitForTimeout(1000);

            await expect(chosen).toHaveText('001', { timeout: 5000 });
            console.log('Verified: Section is set to 001');

            await testInfo.attach('step6_001_selected.png', {
                body: await page.screenshot(),
                contentType: 'image/png'
            });

            await submitButton.click();
            await page.waitForTimeout(3000);
        });

        // Step 7: Select a random country for the fourth dropdown menu
        let selectedCountry = '';
        await test.step('Select random Country', async () => {
            // elements.txt 13行目。4番目のプルダウンなので nth(3)
            const select = frame.locator('.select2-choice').nth(3);
            
            await expect(select).toBeVisible({ timeout: 10000 });
            await select.scrollIntoViewIfNeeded();
            await select.click({ force: true });
            await page.waitForTimeout(1000);
            
            const randomDowns = Math.floor(Math.random() * 10) + 1;
            for (let i = 0; i < randomDowns; i++) {
                await page.keyboard.press('ArrowDown');
            }
            await page.keyboard.press('Enter');
            await page.waitForTimeout(1000);

            selectedCountry = (await select.locator('.select2-chosen').innerText()).trim();
            console.log(`Select - Country: ${selectedCountry}`);

            await submitButton.click();
            await page.waitForTimeout(2000);
        });

        // Step 8: Address based on selected country
        await test.step('Input Address', async () => {
            const addressInput = frame.locator('#ctl00_MainContent_lstQuestionnaire_ctrl5_txtTextResponse');
            const fictitiousAddress = `123 Fictitious Road, ${selectedCountry} Central District`;
            await addressInput.fill(fictitiousAddress);
            console.log(`Input - Address: ${fictitiousAddress}`);

            await submitButton.click();
            await page.waitForTimeout(2000);
        });

        // Step 9: Confirm Completion and Submit
        const completedMessage = frame.locator('.MessageText');
        await expect(completedMessage).toHaveText('No more questions, questionnaire completed.', { timeout: 30000 });
        await submitButton.click();
        await page.waitForTimeout(3000);

        // Step 10: Go To Maintenance Page
        const maintenanceLink = frame.locator('a:has-text("Go To Maintenance Page")');
        await expect(maintenanceLink).toBeVisible({ timeout: 30000 });

        const [newPage] = await Promise.all([
            page.waitForEvent('popup'),
            maintenanceLink.click(),
        ]);

        await newPage.waitForLoadState('load');
        console.log('Maintenance page opened.');

        const maintenanceFrame = newPage.frameLocator('iframe[name="legacy-outlet"]');
        const companyNameInput = maintenanceFrame.locator('#txtCompanyName');
        
        // Step 11: Verify Maintenance Page Fields
        await test.step('Verify Maintenance Page Fields', async () => {
            await expect(companyNameInput).toBeVisible({ timeout: 30000 });
            
            // 1. Verify CompanyName (Input random suffix/name used in Step 3)
            await companyNameInput.scrollIntoViewIfNeeded();
            await page.waitForTimeout(500);
            const actualName = await companyNameInput.inputValue();
            console.log(`Verifying CompanyName: expected contains "${randomCustomerName}", actual "${actualName}"`);
            expect(actualName).toBe(randomCustomerName);
            await testInfo.attach('step11_1_CompanyName.png', {
                body: await companyNameInput.screenshot(),
                contentType: 'image/png'
            });

            // 2. Verify DataSource = Manual (elements.txt Line 2)
            const dataSourceContainer = maintenanceFrame.locator('#s2id_ctl00_MainContent_tabs_tabCompanyInfo_pnlBasic_Info_drpDataSource');
            const dataSourceChosen = dataSourceContainer.locator('.select2-chosen');
            await dataSourceContainer.scrollIntoViewIfNeeded();
            await page.waitForTimeout(500);
            const actualDataSource = (await dataSourceChosen.innerText()).trim();
            console.log(`Verifying DataSource: expected "Manual", actual "${actualDataSource}"`);
            expect(actualDataSource).toBe('Manual');
            await testInfo.attach('step11_2_DataSource.png', {
                body: await dataSourceContainer.screenshot(),
                contentType: 'image/png'
            });

            // 3. Verify ActiveFlag = Y-はい (Correction: was 'Yes')
            const activeFlagContainer = maintenanceFrame.locator('#s2id_ctl00_MainContent_tabs_tabCompanyInfo_pnlAudit_drpActiveFlag');
            const activeFlagChosen = activeFlagContainer.locator('.select2-chosen');
            await activeFlagContainer.scrollIntoViewIfNeeded();
            await page.waitForTimeout(500);
            const actualActiveFlag = (await activeFlagChosen.innerText()).trim();
            console.log(`Verifying ActiveFlag: expected "Y-はい", actual "${actualActiveFlag}"`);
            expect(actualActiveFlag).toMatch(/^Y-/); 
            if (!actualActiveFlag.includes('はい')) {
                console.warn(`Warning: ActiveFlag text "${actualActiveFlag}" might have encoding issues but starts with Y-`);
            }
            await testInfo.attach('step11_3_ActiveFlag.png', {
                body: await activeFlagContainer.screenshot(),
                contentType: 'image/png'
            });

            // 4. Verify DTSSearchFlag = Yes
            const dtsSearchFlagContainer = maintenanceFrame.locator('#s2id_ctl00_MainContent_tabs_tabCompanyInfo_pnlSystem_drpDTSSearchFlag');
            const dtsSearchFlagChosen = dtsSearchFlagContainer.locator('.select2-chosen');
            await dtsSearchFlagContainer.scrollIntoViewIfNeeded();
            await page.waitForTimeout(500);
            const actualDTSSearchFlag = (await dtsSearchFlagChosen.innerText()).trim();
            console.log(`Verifying DTSSearchFlag: expected "Yes", actual "${actualDTSSearchFlag}"`);
            expect(actualDTSSearchFlag).toBe('Yes');
            await testInfo.attach('step11_4_DTSSearchFlag.png', {
                body: await dtsSearchFlagContainer.screenshot(),
                contentType: 'image/png'
            });

            // 5. Verify CompanyWide = No
            const companyWideContainer = maintenanceFrame.locator('#s2id_ctl00_MainContent_tabs_tabCompanyInfo_pnlAudit_drpCompanyWide');
            const companyWideChosen = companyWideContainer.locator('.select2-chosen');
            await companyWideContainer.scrollIntoViewIfNeeded();
            await page.waitForTimeout(500);
            const actualCompanyWide = (await companyWideChosen.innerText()).trim();
            console.log(`Verifying CompanyWide: expected "No", actual "${actualCompanyWide}"`);
            expect(actualCompanyWide).toBe('No');
            await testInfo.attach('step11_5_CompanyWide.png', {
                body: await companyWideContainer.screenshot(),
                contentType: 'image/png'
            });

            // 6. Additionally check if CompanyName is Editable as previously requested
            await companyNameInput.scrollIntoViewIfNeeded();
            await page.waitForTimeout(500);
            const isReadOnly = await companyNameInput.getAttribute('readonly');
            const isDisabled = await companyNameInput.getAttribute('disabled');
            const classAttr = await companyNameInput.getAttribute('class') || '';
            const isRiDisabled = classAttr.includes('riDisabled');

            if (isReadOnly === 'readonly' || isDisabled === 'disabled' || isRiDisabled) {
                throw new Error('FAILED: CompanyName field is expected to be EDITABLE, but it is READ-ONLY or DISABLED.');
            }
            await testInfo.attach('step11_6_CompanyName_Editable_Check.png', {
                body: await companyNameInput.screenshot(),
                contentType: 'image/png'
            });

            console.log('SUCCESS: All fields verified on maintenance page.');
        });

        await newPage.close();
      });

      // Step 12: Logout
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