import { test, expect } from '@playwright/test';
import { BASE_URL } from './config';
import * as fs from 'fs';
import * as path from 'path';

// Load user definitions
const users = JSON.parse(fs.readFileSync(path.join(__dirname, 'account.dat'), 'utf8'));

// Define search criteria for random selection (Profile only)

const searchCriteria = [

    { profile: 'WCO - World Customs Organization - Global HS (IMPORT)', hsCode: '8501' },

    { profile: 'JP - JAPAN - Japan Export Tariff (EXPORT)', hsCode: '0101' },

    { profile: 'JP - JAPAN - Japans Tariff Schedule (IMPORT)', hsCode: '8471' },

    { profile: 'US - UNITED STATES OF AMERICA - US HTS (IMPORT and EXPORT)', hsCode: '8542' },

    { profile: 'US - UNITED STATES OF AMERICA - US Schedule B (EXPORT)', hsCode: '8802' }

];



// Keyword list for TC_031

const keywordList = ['system', 'sensor', 'engine', 'laser', 'telecommunications', 'security', 'equipment'];



test.describe('TC_031 Role Based Access Control - Global Tariffs (Classic) Keyword Search', () => {

  test.describe.configure({ mode: 'parallel' });

  

  users.forEach((user: any) => {

    

    test(`Perform Global Tariffs (Classic) keyword search test for ${user.id}`, async ({ page }) => {

      test.setTimeout(210000);



      console.log(`--- Test Start: ${user.id} ---`);



      // Step 1: Login

      await test.step('Login', async () => {

        await page.goto(BASE_URL + '/Logon.aspx?ReturnUrl=%2fgtm%2fhome');

        await page.getByRole('textbox', { name: 'Company' }).fill(user.company);

        await page.getByRole('textbox', { name: 'Username' }).fill(user.id);

        await page.getByRole('textbox', { name: 'Password' }).fill(user.pass);

        await page.getByRole('link', { name: 'Login' }).click();

      });



      // Step 2: Access Global Tariffs (Classic) and Search

      await test.step('Access Global Tariffs (Classic) and Search', async () => {

        let screenshotTaken = false;

        for (let i = 1; i <= 5; i++) {

            console.log(`\n--- Repetition ${i}/5 for user ${user.id} ---`);

            

            await page.getByRole('button', { name: 'Content' }).click();

            // Use regex for exact match

            await page.locator('a[role="menuitem"]').filter({ hasText: /^Global Tariffs \(Classic\)$/ }).click();



            let mainFrameElement = page.locator('iframe[name="legacy-outlet"]');

            await mainFrameElement.waitFor({ state: 'attached', timeout: 30000 });

            let mainFrame = await mainFrameElement.contentFrame();

            if (!mainFrame) throw new Error('Could not access main iframe content.');



            // Verify page name

            const pageNameLabel = mainFrame.locator('#lbxPageName');

            await expect(pageNameLabel).toContainText('Global Tariffs (Classic)', { timeout: 20000 });

            console.log('SUCCESS: Verified page name contains "Global Tariffs (Classic)".');



            // Select a random search criteria (Profile) and Keyword

            const criteria = searchCriteria[Math.floor(Math.random() * searchCriteria.length)];

            const keyword = keywordList[Math.floor(Math.random() * keywordList.length)];

            console.log(`Using criteria: Profile='${criteria.profile}', Keyword='${keyword}'`);



            // Step 2a: Select Tariff Profile

            const profileInput = mainFrame.locator('#ctl00_MainContent_cmbxTariffSchedule_Input');

            await profileInput.click();

            await profileInput.fill(criteria.profile);

            await page.waitForTimeout(1000); // Wait for input to process

            await page.keyboard.press('Enter');

            

            // Wait for potential postback/reload

            await page.waitForTimeout(12000); 



            // Step 2b: Enter Keyword and Select from Autocomplete

            const hsInput = mainFrame.locator('#txtbxHSNumber');

            

            // Wait for input to be editable

            await expect(hsInput).toBeEditable({ timeout: 10000 });

            await hsInput.click();

            await page.waitForTimeout(2000); // Wait for focus



            // Use pressSequentially with delay

            await hsInput.pressSequentially(keyword, { delay: 300 });

            await page.waitForTimeout(3000); // Wait for autocomplete suggestions



            // Select using Keyboard: ArrowDown 3 to 5 times (randomly) to select 3rd-5th item

            const arrowDownCount = Math.floor(Math.random() * 3) + 3; // 3, 4, or 5

            console.log(`Pressing ArrowDown ${arrowDownCount} times to select autocomplete item.`);

            

            for (let k = 0; k < arrowDownCount; k++) {

                await page.keyboard.press('ArrowDown');

                await page.waitForTimeout(300);

            }

                        await page.keyboard.press('Enter');

            

                        // Wait for conversion/update and Log the Result

                        await page.waitForTimeout(3000);

                        

                        // Log the converted HS Code

                        const convertedValue = await hsInput.inputValue();

                        console.log(`Keyword converted to HS Code: '${convertedValue}'`);

            

                        // Step 2c: Verify Search Results

            const resultTree = mainFrame.locator('#ctl00_MainContent_rtvHSHierarchy');

            await expect(resultTree).toBeVisible({ timeout: 30000 });

            console.log('SUCCESS: Search results tree displayed.');



            // Wait for page to fully render

            await page.waitForTimeout(5000);



            // Take Screenshot 2

            const screenshotBufferResult = await page.screenshot({ fullPage: true });

            await test.info().attach(`global_tariffs_classic_keyword_result_${user.id}.png`, {

                body: screenshotBufferResult,

                contentType: 'image/png'

            });



            console.log('SUCCESS: Result screenshot attached. Test complete.');

            break;

        }

      });



      // Step 3: Logout

      await test.step('Logout', async () => {

        const userButtonName = `MC Test${user.id.substring(6)}`;

        await page.bringToFront();



        const userButton = page.getByRole('button', { name: new RegExp(`^${userButtonName}`)

        });

        await userButton.hover();

        await page.waitForTimeout(500); // Wait for menu to start opening

        await userButton.click();



        const logoutButton = page.locator('#logoutButton');

        await expect(logoutButton).toBeVisible({ timeout: 10000 });

        await logoutButton.click({ force: true });



        await expect(page.getByRole('textbox', { name: 'Company' })).toBeVisible({ timeout: 30000 });

      });

    });

  });

});

