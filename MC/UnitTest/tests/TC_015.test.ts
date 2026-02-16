import { test, expect } from '@playwright/test';
import { BASE_URL } from './config';
import * as fs from 'fs';
import * as path from 'path';

// ベースのユーザー定義はそのまま流用
const users = JSON.parse(fs.readFileSync(path.join(__dirname, 'account.dat'), 'utf8'));

// シナリオに応じた設定
const scenarios = [
    {
        regulation: "CN - CHINA - China Export Control List of Dual-Use Items",
        ecnCodes: ['1C902'],
        hsCodes: ['2805', '2846']
    },
    {
        regulation: "CN - CHINA - China Export License Management Catalogue of Dual-use Items and Technologies",
        ecnCodes: ['1.1.2', '1.1.3'],
        hsCodes: ['2801', '2804']
    },
    {
        regulation: "JP - JAPAN - Japan Export Control Order",
        ecnCodes: ['1.10', '1.11', '1.12', '1.13', '1.14'],
        hsCodes: [] // HSコードは使用しない
    },
    {
        regulation: "JP - JAPAN - Japan Export Control Order: Appended Table 2-3 (Prohibited Items for Export to Russia etc.)",
        ecnCodes: ['1.2'],
        hsCodes: [] // HSコードは使用しない
    },
    {
        regulation: "JP - JAPAN - Japan Foreign Exchange Order",
        ecnCodes: ['10.1', '10.2', '10.3', '10.4'],
        hsCodes: [] // HSコードは使用しない
    },
    {
        regulation: "US - UNITED STATES OF AMERICA - EAR/Commerce Control List",
        ecnCodes: ['2', '3', '4'],
        hsCodes: [] // HSコードは使用しない
    },
    {
        regulation: "US - UNITED STATES OF AMERICA - International Traffic in Arms Regulations (ITAR)",
        ecnCodes: ['4', '5', '6', '7'],
        hsCodes: [] // HSコードは使用しない
    },
    {
        regulation: "US - UNITED STATES OF AMERICA - Supplement No. 6 to Part 746 - Items That Require a License For Russia and Belarus",
        ecnCodes: ['7', '8', '9'],
        hsCodes: [] // HSコードは使用しない
    }
];

test.describe('Role Based Access Control - Specific Regulation Search', () => {
  
  for (const user of users) {
    
    test(`Perform search with specific regulation for ${user.id}`, async ({ page }) => {
      test.setTimeout(180000); // タイムアウトを180秒に設定

      console.log(`--- Specific Regulation Search Test Start: ${user.id} ---`);

      // ---------------------------------------------------
      // Step 1: ログイン処理
      // ---------------------------------------------------
      await test.step('Login', async () => {
        await page.goto(BASE_URL + '/Logon.aspx?ReturnUrl=%2fgtm%2fhome');
        await page.getByRole('textbox', { name: 'Company' }).fill(user.company);
        await page.getByRole('textbox', { name: 'Username' }).fill(user.id);
        await page.getByRole('textbox', { name: 'Password' }).fill(user.pass);
        await page.getByRole('link', { name: 'Login' }).click();
      });

      // ---------------------------------------------------
      // Step 2: 特定のRegulationで検索を実行する (3回繰り返し)
      // ---------------------------------------------------
      await test.step('Perform Search with Specific Regulation 3 times', async () => {
        for (let i = 1; i <= 3; i++) {
            console.log(`\n--- Repetition ${i}/3 for user ${user.id} ---`);
            
            // ECN/Dual Use List ページに移動
            await page.getByRole('button', { name: 'Content' }).click();
            await page.getByRole('menuitem', { name: 'ECN/Dual Use List Query' }).click();

            let mainFrameElement = page.locator('iframe[name="legacy-outlet"]');
            await mainFrameElement.waitFor({ state: 'attached', timeout: 30000 });
            let mainFrame = await mainFrameElement.contentFrame();
            if (!mainFrame) throw new Error('メインのiframeコンテンツにアクセスできませんでした。');

            // ランダムにシナリオを選択
            const scenario = scenarios[Math.floor(Math.random() * scenarios.length)];
            console.log(`ℹ️ Test Scenario: Regulation "${scenario.regulation}"`);

            // シナリオに基づきRegulation Listを選択
            console.log(`ℹ️ Regulation List [${scenario.regulation}] を選択します...`);
            const regulationInput = mainFrame.locator('#ctl00_MainContent_cmbxRegulationList_Input');
            await regulationInput.click();
            const regulationDropdown = mainFrame.locator('#ctl00_MainContent_cmbxRegulationList_DropDown');
            await regulationDropdown.getByText(scenario.regulation, { exact: true }).click();
            await expect(regulationInput).toHaveValue(scenario.regulation);
            console.log('✅ Regulation Listが選択されました。');

            // 画面遷移を待つ
            console.log('ℹ️ 画面遷移を3秒間待ちます...');
            await page.waitForTimeout(3000);

            mainFrameElement = page.locator('iframe[name="legacy-outlet"]');
            await mainFrameElement.waitFor({ state: 'attached', timeout: 30000 });
            mainFrame = await mainFrameElement.contentFrame();
            if (!mainFrame) throw new Error('リロード後のiframeコンテンツにアクセスできませんでした。');
            
            // ECNかHSかを、シナリオで利用可能なコードに基づいて決定
            const canSearchEcn = scenario.ecnCodes.length > 0;
            const canSearchHs = scenario.hsCodes.length > 0;
            let searchType = '';

            if (canSearchEcn && canSearchHs) {
                searchType = Math.random() < 0.5 ? 'ECN' : 'HS';
            } else if (canSearchEcn) {
                searchType = 'ECN';
            } else if (canSearchHs) {
                searchType = 'HS';
            } else {
                throw new Error(`No ECN or HS codes defined for regulation: ${scenario.regulation}`);
            }
            
            // 決定した検索タイプでコードを入力
            if (searchType === 'ECN') {
                const randomEcn = scenario.ecnCodes[Math.floor(Math.random() * scenario.ecnCodes.length)];
                console.log(`ℹ️ ECNコードで検索します: ${randomEcn}`);
                const ecnInput = mainFrame.locator('#ctl00_MainContent_cmbxStatusBarECN_Input');
                await ecnInput.click();
                await ecnInput.type(randomEcn, { delay: 200 });
                const dropdown = mainFrame.locator('#ctl00_MainContent_cmbxStatusBarECN_DropDown');
                await expect(dropdown.locator('li.rcbItem').first()).toBeVisible({timeout: 15000});
                await dropdown.locator('li.rcbItem').first().click();

            } else { // searchType === 'HS'
                const randomHs = scenario.hsCodes[Math.floor(Math.random() * scenario.hsCodes.length)];
                console.log(`ℹ️ HSコードで検索します: ${randomHs}`);
                const hsInput = mainFrame.locator('#ctl00_MainContent_cmbxStatusBarHSNumber_Input');
                await hsInput.click();
                await hsInput.type(randomHs, { delay: 200 });
                const dropdown = mainFrame.locator('#ctl00_MainContent_cmbxStatusBarHSNumber_DropDown');
                await expect(dropdown.locator('li.rcbItem').first()).toBeVisible({timeout: 15000});
                await dropdown.locator('li.rcbItem').first().click();
            }

            // 宛先国を選択
            const countries = [
                { search: 'Iran', expect: 'IRAN, ISLAMIC REPUBLIC OF' },
                { search: 'KOREA, DEMOCRATIC', expect: 'KOREA, DEMOCRATIC PEOPLE\'S REPUBLIC OF' },
                { search: 'Syria', expect: 'SYRIAN ARAB REPUBLIC' },
                { search: 'Cuba', expect: 'CUBA' },
                { search: 'Russia', expect: 'RUSSIAN FEDERATION' },
            ];
            const randomCountry = countries[Math.floor(Math.random() * countries.length)];
            console.log(`ℹ️ 宛先国として "${randomCountry.search}" を選択します...`);

            const countryInput = mainFrame.locator('#ctl00_MainContent_cmbxDestinationCountry_Input');
            await countryInput.click(); // Click to open

            const countryDropdown = mainFrame.locator('#ctl00_MainContent_cmbxDestinationCountry_DropDown');
            // Wait for the dropdown to be visible before trying to click an item
            await expect(countryDropdown).toBeVisible({ timeout: 10000 });
            await countryDropdown.locator('li', { hasText: randomCountry.search }).first().click();
            
            // Verify selection
            await expect(countryInput).toHaveValue(randomCountry.expect);
            console.log('✅ 宛先国が選択されました。');

            // Searchボタンをクリックして結果を待つ
            console.log('ℹ️ Searchボタンをクリックします...');
            await mainFrame.locator('#ctl00_MainContent_lnxbtnSearch').click();
            
            console.log('ℹ️ 検索結果が表示されるのを待ちます...');
            const resultsGrid = mainFrame.locator('#ctl00_MainContent_rgdMain');
            await expect(resultsGrid).toBeVisible({ timeout: 30000 });
    
            // グリッドの中に最初のデータ行が表示されるまで待つことで、データ描画を確実にする
            const firstRow = resultsGrid.locator('table.rgMasterTable > tbody > tr').first();
            await expect(firstRow).toBeVisible({ timeout: 20000 });
    
            // 最終的な描画の安定性を確保するために追加の待機
            console.log('ℹ️ 最終描画のため3秒間待機します...');
            await page.waitForTimeout(3000);
            
            console.log('✅ 検索結果が完全に表示されました。');

            // ---------------------------------------------------
            // Step 4: 検索結果の展開とタブのスクリーンショット
            // ---------------------------------------------------
            console.log('ℹ️ 検索結果の最初のノードを展開します...');
            const expandButton = mainFrame.locator('[id$="_GECBtnExpandColumn"]').first();
            await expandButton.click();

            // 展開後のタブが表示されるまで待機
            const tabStrip = mainFrame.locator('.RadTabStrip');
            await expect(tabStrip).toBeVisible({ timeout: 15000 });
            console.log('✅ ノードが展開されました。');

            console.log('ℹ️ 表示されているタブを動的に取得します...');
            const tabContainer = mainFrame.locator('div.rtsLevel.rtsLevel1');
            const tabNames = await tabContainer.getByRole('link').evaluateAll(links => 
                links.map(link => (link.textContent || '').trim()).filter(name => name)
            );
            console.log(`✅ 取得したタブ: ${tabNames.join(', ')}`);


            // "CONTROL REASONS" タブの検証
            const controlReasonsTabName = tabNames.find(name => name.toUpperCase().startsWith('CONTROL REASONS'));

            if (!controlReasonsTabName) {
                console.log('⚠️ "CONTROL REASONS" タブが見つからなかったため、確認をスキップします。');
            } else {
                console.log(`ℹ️ タブ "${controlReasonsTabName}" を選択しています...`);
                await tabStrip.getByRole('link', { name: controlReasonsTabName }).click();

                const controlReasonsGrid = mainFrame.locator('div.RadGrid[id*="rgdControlReasons"]');
                await expect(controlReasonsGrid).toBeVisible({ timeout: 15000 });
                console.log('✅ "CONTROL REASONS" タブのコンテンツが表示されました。');

                const rowsWithExpandButtons = controlReasonsGrid.locator('tr.rgRow:has(input[title="Expand"])');
                const count = await rowsWithExpandButtons.count();

                if (count > 0) {
                    console.log(`ℹ️ ${count}個の展開可能なノードが見つかりました。ランダムに1つを選択して展開します...`);
                    const randomIndex = Math.floor(Math.random() * count);
                    const randomRowToExpand = rowsWithExpandButtons.nth(randomIndex);
                    
                    const rowId = await randomRowToExpand.getAttribute('id');
                    if (!rowId) {
                        throw new Error('ランダムに選択された行のIDが取得できませんでした。テストを続行できません。');
                    }

                    const stableRowLocator = mainFrame.locator(`tr[id="${rowId}"]`);
                    await stableRowLocator.locator('input[title="Expand"]').click();

                    await expect(stableRowLocator.locator('input[title="Collapse"]')).toBeVisible({ timeout: 10000 });
                    console.log('✅ サブノードがランダムに展開され、詳細情報が表示されました。');

                    // 展開された詳細内容（隣接する行）を取得
                    const detailRow = stableRowLocator.locator('+ tr');
                    
                    // 詳細内容に "ONLY Applies" が含まれるかで、許可リストか禁止リストか判断
                    console.log('ℹ️ 詳細内容を確認し、許可/禁止リストを判断します...');
                    const isPermittedList = (await detailRow.getByText('ONLY Applies').count()) > 0;
                    
                    // "Country list" に指定した国が表示されているか、または表示されていないかを検証
                    const countryCell = detailRow.locator('td', { hasText: randomCountry.expect });

                    if (isPermittedList) {
                        console.log('ℹ️ 許可リストのため、国が表示されていないことを確認します...');
                        await expect(countryCell.first()).toBeHidden({ timeout: 10000 });
                        console.log(`✅ Country listに国「${randomCountry.expect}」が正しく表示されていません。`);
                    } else {
                        console.log('ℹ️ 禁止リストのため、国が表示されていることを確認します...');
                        await expect(countryCell.first()).toBeVisible({ timeout: 10000 });
                        console.log(`✅ Country listに国「${randomCountry.expect}」が正しく表示されています。`);
                    }

                    // Take a full page screenshot
                    console.log('📸 詳細情報のスクリーンショットを撮影します...');
                    const screenshotBuffer = await page.screenshot({ fullPage: true });
                    await test.info().attach(`${user.id}-${scenario.regulation}-${i}-CONTROL_REASONS_Expanded.png`, {
                        body: screenshotBuffer,
                        contentType: 'image/png',
                    });

                } else {
                    console.log('⚠️ "CONTROL REASONS" 内に展開可能なノードが見つかりませんでした。');
                }
            }

            // ---------------------------------------------------
            // Step 5: クリーンアップ
            // ---------------------------------------------------
            console.log('ℹ️ "New Search" をクリックして次回のループのためにページをリセットします。');
            await mainFrame.locator('#ctl00_MainContent_hyxlnkNewSearch').click();

            // ページがリセットされたことを確認
            await expect(mainFrame.locator('#ctl00_MainContent_rgdMain')).toBeHidden();
            console.log('✅ ページがリセットされました。');
        }
      });

      // ---------------------------------------------------
      // Step 3: ログアウト処理
      // ---------------------------------------------------
      await test.step('Logout', async () => {
        const userButtonName = `MC Test${user.id.substring(6)}`;
        await page.bringToFront();

        const userMenuButton = page.getByRole('button', { name: new RegExp(`^${userButtonName}`) });
        await userMenuButton.waitFor({ state: 'visible', timeout: 10000 });
        await userMenuButton.click();

        const signOutButton = page.getByRole('button', { name: ' Sign out' });
        await signOutButton.waitFor({ state: 'visible', timeout: 5000 });
        
        await signOutButton.click({ force: true }); 
        
        await expect(page.getByRole('textbox', { name: 'Company' })).toBeVisible({ timeout: 30000 });
      });
      
    });
  }
});

