import { test, expect } from '@playwright/test';

test('E2E_001: Login and Navigate to GC Search/Reporting Lookup', async ({ page }) => {
  let responseCode: string | null = null;

  // Helper function to get a random integer between min and max (inclusive)
  const getRandomInt = (min, max) => {
    min = Math.ceil(min);
    max = Math.floor(max);
    return Math.floor(Math.random() * (max - min + 1)) + min;
  };

  // Read credentials from account.dat
  const fs = require('fs');
  const path = require('path');
  const accountsData = fs.readFileSync(path.resolve(__dirname, '../account.dat'), 'utf-8');
  const accounts = JSON.parse(accountsData);
  const user = accounts.find(acc => acc.id === 'MCTest1');

  if (!user) {
    throw new Error('User MCTest1 not found in account.dat');
  }

  // 1. Login
  await page.goto('https://ogt-gtm-web-eu-imp.7166.aws.thomsonreuters.com/gtm/home');
  await page.locator('#ctl00_RightSide_txtUserName').fill(user.id);
  await page.locator('#ctl00_RightSide_txtPassword').fill(user.pass);
  await page.locator('#ctl00_RightSide_txtCompany').fill(user.company);
  await page.locator('#btxLoginGroup').click();
  
  // Verify login success
  await expect(page).toHaveURL(/.*home/);

  // 2. Navigate to GC Search/Reporting Lookup
  await page.locator('#button-menu-0').click();
  const gcSearchLink = page.locator('a:has-text("GC Search/Reporting Lookup")');
  await gcSearchLink.waitFor();
  await gcSearchLink.click();

  // Wait for navigation to complete (using URL check for safety)
  // Adjust the regex based on the actual URL of the lookup page if known, otherwise generic check
  // Based on context, it seems to be a lookup page.
  await page.waitForLoadState('domcontentloaded');
  
  console.log('Navigated to GC Search/Reporting Lookup.');
  
  // 3. Select "Approved-承認済み" in the status filter
  console.log('Searching for Status filter...');

  // Wait a moment for frames to fully settle
  await page.waitForTimeout(5000);

  let frame = page.mainFrame();
  let filterContainer = null;
  
  // Retry finding the frame a few times
  for (let attempt = 1; attempt <= 3; attempt++) {
      const frames = page.frames();
      console.log(`Attempt ${attempt}: Checking ${frames.length} frames...`);
      
      for (const f of frames) {
        console.log(`Checking frame: ${f.url()}`);
        // Try partial ID match to be safer
        const found = await f.$('[id*="drpMainDropdown"]');
        if (found) {
          frame = f;
          // Locate the Select2 container specifically (starts with s2id_)
          // We use the partial match to find the select, then look for its sibling/container if needed, 
          // or directly look for the s2id container with partial match.
          filterContainer = frame.locator('div[id*="s2id_"][id*="drpMainDropdown"]');
          console.log(`Found filter in frame: ${f.url()}`);
          break;
        }
      }
      if (filterContainer) break;
      await page.waitForTimeout(2000);
  }
  
  if (!filterContainer) {
      console.log('Filter not found in frames. Taking debug screenshot.');
      await page.screenshot({ path: 'debug_frames_failed.png', fullPage: true });
      // Also fallback to checking main page with the loose selector
      filterContainer = page.locator('div[id*="s2id_"][id*="drpMainDropdown"]');
  }

  // Ensure visible
  await filterContainer.waitFor({ state: 'visible', timeout: 10000 });
  
  // Check current text (Expect "All")
  const currentText = await filterContainer.locator('.select2-chosen').textContent();
  console.log(`Current filter selection: ${currentText}`);

  console.log('Clicking "All" container to open dropdown...');
  await filterContainer.click();

  // Wait for the option to appear. 
  const option = frame.locator('div.select2-drop ul.select2-results li').filter({ hasText: 'Approved-承認済み' });
  
  // Wait for it to be visible before clicking
  await option.waitFor({ state: 'visible', timeout: 5000 }).catch(async () => {
      console.log('Option not found. Dumping page content.');
      const content = await page.content();
      const fs = require('fs');
      fs.writeFileSync('E2E_001_debug_option.html', content);
      throw new Error('Option "Approved-承認済み" not found.');
  });
  
  await option.click();

  // The selection triggers a postback
  await page.waitForLoadState('domcontentloaded');
  
  console.log('Selected "Approved-承認済み". Waiting for grid to update...');

  // Wait for potential Telerik Ajax loading panel to disappear
  // Telerik grids often use .raDiv for loading masks
  const loadingPanel = frame.locator('.raDiv, .rgActionMode');
  if (await loadingPanel.count() > 0) {
      console.log('Waiting for loading panel to disappear...');
      await loadingPanel.first().waitFor({ state: 'hidden', timeout: 15000 });
  }

  // Add a hard wait to be safe, as Telerik updates can be tricky
  await page.waitForTimeout(5000);
  
  // Verify the selection is applied
  await expect(filterContainer.locator('.select2-chosen')).toHaveText('Approved-承認済み');

  console.log('Extracting Product Numbers from the updated table...');

  // 4. Extract Product Numbers
  // User confirmed ProductNum is the first column (index 0).
  const productNumIndex = 0;
  console.log(`Extracting Product Numbers from column index: ${productNumIndex}`);

  // Get data rows (Telerik uses rgRow and rgAltRow classes)
  // Ensure we are looking inside the correct frame
  const rows = await frame.locator('tr.rgRow, tr.rgAltRow').all();
  
  const extractedProductNums: string[] = [];
  const numProductsToExtract = getRandomInt(1, 3);
  const countToExtract = Math.min(rows.length, numProductsToExtract);

  console.log(`Found ${rows.length} rows. Will attempt to extract ${countToExtract} random product(s)...`);

  for (let i = 0; i < countToExtract; i++) {
      const cells = await rows[i].locator('td').all();
      if (cells.length > productNumIndex) {
          const text = await cells[productNumIndex].textContent();
          if (text) {
              extractedProductNums.push(text.trim());
          }
      }
  }

  console.log('Extracted Product Numbers:', extractedProductNums);

  // --- Part 2: DPS Search/Reporting Lookup ---
  console.log('--- Starting Part 2: DPS Search/Reporting Lookup ---');

  // 5. Navigate to DPS Search/Reporting Lookup
  // Click the DPS menu button. ID suggests it is button-menu-1, but text is safer.
  const dpsMenuButton = page.locator('button').filter({ hasText: 'DPS' }).first();
  console.log('Clicking DPS menu button...');
  await dpsMenuButton.click();
  
  const dpsSearchLink = page.locator('a').filter({ hasText: 'DPS Search/Reporting Lookup' }).first();
  await dpsSearchLink.waitFor();
  await dpsSearchLink.click();

  await page.waitForLoadState('domcontentloaded');
  console.log('Navigated to DPS Search/Reporting Lookup.');

  // 6. Find table and extract CompanyIDs
  console.log('Searching for DPS data table...');
  await page.waitForTimeout(5000); // Allow frames to load

  let dpsFrame = page.mainFrame();
  let dpsTableFound = false;
  
  // Reuse frame scanning logic
  for (let attempt = 1; attempt <= 3; attempt++) {
      const frames = page.frames();
      for (const f of frames) {
        // Look for Telerik grid rows
        const rowCount = await f.locator('tr.rgRow').count();
        if (rowCount > 0) {
          dpsFrame = f;
          dpsTableFound = true;
          console.log(`Found DPS data table in frame: ${f.url()}`);
          break;
        }
      }
      if (dpsTableFound) break;
      await page.waitForTimeout(2000);
  }

  if (!dpsTableFound) {
      console.log('DPS Table not found in frames. Taking screenshot.');
      await page.screenshot({ path: 'E2E_001_DPS_table_not_found.png', fullPage: true });
      // Proceeding might fail, but let's try on main page just in case
  }

  // Column indices
  const companyIdIndex = 3;
  const internalJudgementIndex = 10;
  
  const extractedCompanyIds: string[] = [];
  // Use dpsFrame for locating rows
  const dpsRows = await dpsFrame.locator('tr.rgRow, tr.rgAltRow').all();
  
  const numCompaniesToExtract = getRandomInt(2, 4);
  console.log(`Scanning ${dpsRows.length} rows for up to ${numCompaniesToExtract} random CompanyID(s)...`);

  for (const row of dpsRows) {
      if (extractedCompanyIds.length >= numCompaniesToExtract) break;

      const cells = await row.locator('td').all();
      
      if (cells.length > internalJudgementIndex) {
          let internalJudgement = await cells[internalJudgementIndex].textContent();
          
          if (internalJudgement) {
              // Normalize whitespace: replace non-breaking space with normal space and trim
              const cleanedJudgement = internalJudgement.replace(/\u00A0/g, ' ').trim();
              
              // Debug log to see what we are checking (limiting length to avoid spam)
              // console.log(`Row check: InternalJudgement="${cleanedJudgement}"`);

              if (cleanedJudgement.length > 0) {
                  const companyId = await cells[companyIdIndex].textContent();
                  if (companyId) {
                      const id = companyId.trim();
                      extractedCompanyIds.push(id);
                      console.log(`Found match: CompanyID=${id}, InternalJudgement="${cleanedJudgement}"`);
                  }
              }
          }
      }
  }

  if (extractedCompanyIds.length < 2) {
      console.warn(`Warning: Only found ${extractedCompanyIds.length} matching records (Minimum 2 required).`);
  }

  console.log('Extracted CompanyIDs:', extractedCompanyIds);

  // --- Part 3: Create New Sales Order ---
  console.log('--- Starting Part 3: Create New Sales Order ---');

  // 7. Navigate to Export -> Create New Sales Order
  const exportMenuButton = page.locator('button').filter({ hasText: 'Export' }).first();
  console.log('Clicking Export menu button...');
  await exportMenuButton.click();
  
  const createSalesOrderLink = page.locator('a').filter({ hasText: 'Create New Sales Order' }).first();
  await createSalesOrderLink.waitFor();
  await createSalesOrderLink.click();

  await page.waitForLoadState('domcontentloaded');
  console.log('Navigated to Create New Sales Order.');

  // 8. Interact with Create New Sales Order form
  console.log('Searching for Sales Order form...');
  await page.waitForTimeout(5000);

  let soFrame = page.mainFrame();
  let soFormFound = false;

  // Find frame with the Submit button
  for (let attempt = 1; attempt <= 3; attempt++) {
      const frames = page.frames();
      for (const f of frames) {
          const btn = await f.$('#ctl00_MainContent_lnxbtnSingleSubmit');
          if (btn) {
              soFrame = f;
              soFormFound = true;
              console.log(`Found Sales Order form in frame: ${f.url()}`);
              break;
          }
      }
      if (soFormFound) break;
      await page.waitForTimeout(2000);
  }

  if (!soFormFound) {
      console.log('Sales Order form not found. Taking screenshot.');
      await page.screenshot({ path: 'E2E_001_SO_form_not_found.png', fullPage: true });
      // Try main page as fallback
  }

  // Select "01-個別" from the dropdown
  // Since we don't have the ID, we look for the select2-choice inside the frame
  const dropdown = soFrame.locator('.select2-choice').first();
  console.log('Clicking Sales Order dropdown...');
  await dropdown.click();

  // Wait for option "01-個別" and click
  // Options are usually attached to the body of the frame or main page
  // We check the frame first
  const soOption = soFrame.locator('div.select2-drop ul.select2-results li').filter({ hasText: '01-個別' });
  
  await soOption.waitFor({ state: 'visible', timeout: 5000 }).catch(async () => {
      console.log('Option "01-個別" not found in frame. Checking main page...');
      // Sometimes Select2 attaches to the main document body even if inside an iframe (rare but possible)
      // Or maybe the selector needs to be broader
  });
  
  if (await soOption.isVisible()) {
      await soOption.click();
  } else {
      // Fallback search in page body just in case
       const pageOption = page.locator('div.select2-drop ul.select2-results li').filter({ hasText: '01-個別' });
       await pageOption.waitFor();
       await pageOption.click();
  }
  console.log('Selected "01-個別".');

  // Click Submit
  console.log('Clicking Submit...');
  await soFrame.locator('#ctl00_MainContent_lnxbtnSingleSubmit').click();
  
  await page.waitForLoadState('domcontentloaded');
  console.log('Submitted Sales Order form. Waiting for update...');

  // 9. Second Selection: "01-仲介貿易"
  console.log('--- Starting Second Selection: 01-仲介貿易 ---');
  await page.waitForTimeout(5000); // Increased wait for stability

  // Re-acquire frame as the page likely reloaded/navigated
  let step2Frame = null;
  const frames2 = page.frames();
  for (const f of frames2) {
      // Look for the Submit button again to confirm we are in the right frame
      const btn = await f.$('#ctl00_MainContent_lnxbtnSingleSubmit');
      if (btn) {
          step2Frame = f;
          console.log(`Found Sales Order form (Step 2) in frame: ${f.url()}`);
          break;
      }
  }

  if (!step2Frame) {
      console.log('Step 2 Frame not found. Taking screenshot.');
      await page.screenshot({ path: 'E2E_001_Step2_frame_missing.png', fullPage: true });
      throw new Error('Could not find frame for Step 2.');
  }

  const dropdown2 = step2Frame.locator('.select2-choice').nth(1);
  // Better strategy: locate by proximity to the label or ID pattern 'ctrl1' derived from elements.txt
  // The provided HTML shows hidden inputs with 'ctrl1'. Let's try to find a container with 'ctrl1' or rely on the label.
  
  const questionLabel = step2Frame.locator('div.paragraph').filter({ hasText: '2. 取引の種類を選んでください' });
  await questionLabel.waitFor({ state: 'visible' });
  console.log('Found label: 2. 取引の種類を選んでください');

  // Try to find the dropdown associated with this question.
  // Assuming they are siblings or in the same parent.
  // Let's try to find a .select2-choice that is "below" this label in the DOM order.
  // Or simpler: The inputs have 'ctrl1'. Let's look for a select2 container with 'ctrl1' in its ID or inside a 'ctrl1' container.
  
  // Attempt 1: Look for select2 choice inside a container that also has the label, or just the next one.
  // Let's count visible dropdowns.
  const visibleDropdowns = await step2Frame.locator('.select2-choice:visible').count();
  console.log(`Visible dropdowns found: ${visibleDropdowns}`);
  
  let targetDropdown;
  if (visibleDropdowns >= 2) {
      console.log('Selecting the second visible dropdown (assuming order 1 -> 2)');
      targetDropdown = step2Frame.locator('.select2-choice:visible').nth(1);
  } else {
      console.log('Only 1 or 0 visible dropdowns. Trying to find one near the label.');
      // Fallback: assume the label's parent contains the dropdown
      targetDropdown = step2Frame.locator('div.paragraph').filter({ hasText: '2. 取引の種類を選んでください' }).locator('xpath=..').locator('.select2-choice');
  }

  // Debug log current text
  const text2 = await targetDropdown.textContent();
  console.log(`Target dropdown text: "${text2?.trim()}"`);

  await targetDropdown.click();

  const option2 = step2Frame.locator('div.select2-drop ul.select2-results li').filter({ hasText: '01-仲介貿易' });
  
  // Handling potential visibility issues same as before
  await option2.waitFor({ state: 'visible', timeout: 5000 }).catch(async () => {
      console.log('Option "01-仲介貿易" not found in frame. Checking main page...');
  });

  if (await option2.isVisible()) {
      await option2.click();
  } else {
       const pageOption2 = page.locator('div.select2-drop ul.select2-results li').filter({ hasText: '01-仲介貿易' });
       await pageOption2.waitFor();
       await pageOption2.click();
  }
  console.log('Selected "01-仲介貿易".');

  // Click Submit again
  console.log('Clicking Submit (2nd time)...');
  await step2Frame.locator('#ctl00_MainContent_lnxbtnSingleSubmit').click();
  
  await page.waitForLoadState('domcontentloaded');
  console.log('Submitted second form.');

  // 10. Third Selection: "01-配慮国向け取引"
  console.log('--- Starting Third Selection: 01-配慮国向け取引 ---');
  await page.waitForTimeout(5000);

  // Re-verify frame for step 3 (likely same as step 2, but let's be safe)
  let step3Frame = step2Frame; 
  if (await step3Frame.isDetached()) {
       console.log('Frame detached, re-scanning for Step 3...');
       const frames3 = page.frames();
       for (const f of frames3) {
          const btn = await f.$('#ctl00_MainContent_lnxbtnSingleSubmit');
          if (btn) {
              step3Frame = f;
              break;
          }
       }
  }

  const label3Text = '3.配慮国/軍・国防省向けの取引ですか？';
  const label3 = step3Frame.locator('div.paragraph').filter({ hasText: label3Text });
  await label3.waitFor({ state: 'visible', timeout: 10000 });
  console.log(`Found label: ${label3Text}`);

  // Find 3rd dropdown
  const visibleDropdowns3 = await step3Frame.locator('.select2-choice:visible').count();
  console.log(`Visible dropdowns found (Step 3): ${visibleDropdowns3}`);
  
  let dropdown3;
  if (visibleDropdowns3 >= 3) {
      console.log('Selecting the third visible dropdown');
      dropdown3 = step3Frame.locator('.select2-choice:visible').nth(2);
  } else {
      console.log('Less than 3 visible dropdowns. Trying context-based search.');
      // Fallback: look near the label
      dropdown3 = label3.locator('xpath=..').locator('.select2-choice');
  }
  
  // Debug log
  console.log(`Target dropdown 3 text: "${(await dropdown3.textContent())?.trim()}"`);

  await dropdown3.click();

  const option3 = step3Frame.locator('div.select2-drop ul.select2-results li').filter({ hasText: '01-配慮国向け取引' });
  
  await option3.waitFor({ state: 'visible', timeout: 5000 }).catch(async () => {
       console.log('Option "01-配慮国向け取引" not found. Checking main page...');
  });

  if (await option3.isVisible()) {
      await option3.click();
  } else {
       const pageOption3 = page.locator('div.select2-drop ul.select2-results li').filter({ hasText: '01-配慮国向け取引' });
       await pageOption3.waitFor();
       await pageOption3.click();
  }
  console.log('Selected "01-配慮国向け取引".');

  // Click Submit again
  console.log('Clicking Submit (3rd time)...');
  await step3Frame.locator('#ctl00_MainContent_lnxbtnSingleSubmit').click();
  
  await page.waitForLoadState('domcontentloaded');
  console.log('Submitted third form.');

  // 11. Fourth, Fifth, Sixth Selections (Simultaneous)
  console.log('--- Starting Simultaneous Selections: X, X1, 001 ---');
  await page.waitForTimeout(5000);

  // Re-verify frame
  let step4Frame = step3Frame;
  if (await step4Frame.isDetached()) {
       console.log('Frame detached, re-scanning for Step 4...');
       const frames4 = page.frames();
       for (const f of frames4) {
          const btn = await f.$('#ctl00_MainContent_lnxbtnSingleSubmit');
          if (btn) {
              step4Frame = f;
              break;
          }
       }
  }

  // --- Selection 4: Group Org Code (X) ---
  const label4Text = 'グループ組織コードを入力してください';
  console.log(`Processing Q4: ${label4Text}`);
  
  // Try finding by ID first (ctrl3)
  let dropdown4 = step4Frame.locator('[id*="ctrl3"] .select2-choice').first();
  if (await dropdown4.count() === 0) {
       // Try direct ID on the choice or container
       dropdown4 = step4Frame.locator('div[id*="ctrl3"]').locator('.select2-choice');
  }
  // Fallback to label proximity if ID fails
  if (await dropdown4.count() === 0) {
      console.log('ID search failed for Q4, using label proximity.');
      const label4 = step4Frame.locator('div.paragraph').filter({ hasText: label4Text });
      dropdown4 = label4.locator('xpath=..').locator('.select2-choice');
  }

  console.log(`Dropdown 4 found: ${await dropdown4.count()}`);
  await dropdown4.click();

  const option4 = step4Frame.locator('div.select2-drop ul.select2-results li').filter({ hasText: 'X' });
  // Use a stricter filter for "X" to avoid matching "X1" or others, but allow whitespace
  // Loop through to find exact match if needed, or assume 'X' is unique enough or listed first.
  // Actually, 'X' might match 'X1' with hasText. Let's use exact text match if possible.
  const option4Exact = step4Frame.locator('div.select2-drop ul.select2-results li').filter({ hasText: /^X$/ });
  
  if (!await option4Exact.isVisible()) {
       // Try main page
       const pageOption4 = page.locator('div.select2-drop ul.select2-results li').filter({ hasText: /^X$/ });
       if (await pageOption4.isVisible()) {
           await pageOption4.click();
       } else {
           // Fallback to loose match if exact fail
           await option4.first().click();
       }
  } else {
      await option4Exact.click();
  }
  console.log('Selected "X" for Group Org Code.');
  await page.waitForTimeout(1000);

  // --- Selection 5: Division Org Code (X1) ---
  const label5Text = '本部組織コードを入力してください';
  console.log(`Processing Q5: ${label5Text}`);
  
  let dropdown5 = step4Frame.locator('[id*="ctrl4"] .select2-choice').first();
  if (await dropdown5.count() === 0) {
       dropdown5 = step4Frame.locator('div[id*="ctrl4"]').locator('.select2-choice');
  }
  if (await dropdown5.count() === 0) {
      console.log('ID search failed for Q5, using label proximity.');
      const label5 = step4Frame.locator('div.paragraph').filter({ hasText: label5Text });
      dropdown5 = label5.locator('xpath=..').locator('.select2-choice');
  }

  await dropdown5.click();

  const option5 = step4Frame.locator('div.select2-drop ul.select2-results li').filter({ hasText: 'X1' });
  
  if (!await option5.isVisible()) {
       const pageOption5 = page.locator('div.select2-drop ul.select2-results li').filter({ hasText: 'X1' });
       await pageOption5.click();
  } else {
      await option5.click();
  }
  console.log('Selected "X1" for Division Org Code.');
  await page.waitForTimeout(1000);

  // --- Selection 6: BU Code (001) ---
  const label6Text = 'BUコードを入力してください';
  console.log(`Processing Q6: ${label6Text}`);
  
  let dropdown6 = step4Frame.locator('[id*="ctrl5"] .select2-choice').first();
  if (await dropdown6.count() === 0) {
       dropdown6 = step4Frame.locator('div[id*="ctrl5"]').locator('.select2-choice');
  }
  if (await dropdown6.count() === 0) {
       console.log('ID search failed for Q6, using label proximity.');
       const label6 = step4Frame.locator('div.paragraph').filter({ hasText: label6Text });
       dropdown6 = label6.locator('xpath=..').locator('.select2-choice');
  }

  await dropdown6.click();

  const option6 = step4Frame.locator('div.select2-drop ul.select2-results li').filter({ hasText: '001' });
  
  if (!await option6.isVisible()) {
       const pageOption6 = page.locator('div.select2-drop ul.select2-results li').filter({ hasText: '001' });
       await pageOption6.click();
  } else {
      await option6.click();
  }
  console.log('Selected "001" for BU Code.');

  // Click Intermediate Submit (Single Submit)
  console.log('Clicking Intermediate Submit...');
  await step4Frame.locator('#ctl00_MainContent_lnxbtnSingleSubmit').click();
  
  await page.waitForLoadState('domcontentloaded');
  
  // Verify Success Message
  console.log('Waiting for completion message...');
  const successMsg = step4Frame.locator('.IPMessage-Success .MessageText');
  await successMsg.waitFor({ state: 'visible', timeout: 10000 });
  await expect(successMsg).toHaveText(/No more questions, questionnaire completed/i);
  console.log('Verified message: "No more questions, questionnaire completed."');

  // Click Final Submit
  console.log('Clicking Final Submit (Complete)...');
  const finalSubmitBtn = step4Frame.locator('#ctl00_MainContent_lnxbtnComplete');
  
  // Wait for button to be enabled (attribute disabled to be removed)
  await expect(finalSubmitBtn).toBeEnabled({ timeout: 10000 });
  await finalSubmitBtn.click();
  console.log('Clicked Final Submit. Waiting for confirmation dialog...');
  
  // Handle Confirmation Dialog
  console.log('Searching for OK button in all frames...');
  await page.waitForTimeout(3000); // Initial wait for dialog animation

  let okBtnLocator = null;
  let okFrame = null;
  
  // Retry loop to find the button
  for (let attempt = 1; attempt <= 3; attempt++) {
      const allFrames = page.frames();
      console.log(`Attempt ${attempt}: Checking ${allFrames.length} frames for OK button...`);
      
      for (const f of allFrames) {
          // Use :visible to ignore hidden templates (Alert, Prompt, Confirm templates often exist but hidden)
          const btn = f.locator('button.rwOkBtn:visible');
          if (await btn.count() > 0) {
              okFrame = f;
              okBtnLocator = btn.first(); // Use first if multiple visible (rare)
              console.log(`Found visible OK button in frame: ${f.url()}`);
              break;
          }
      }
      
      if (okBtnLocator) break;
      
      // Check main page
      const mainBtn = page.locator('button.rwOkBtn:visible');
      if (await mainBtn.count() > 0) {
           okBtnLocator = mainBtn.first();
           console.log('Found visible OK button on main page.');
           break;
      }

      await page.waitForTimeout(3000);
  }

  if (okBtnLocator) {
      await okBtnLocator.click();
      console.log('Clicked OK on confirmation dialog.');
      // Wait for the dialog to disappear or page to update
      await page.waitForTimeout(3000); 

      // Extract the dynamic code (Response Code) after handling the confirmation dialog
      const codeElement = step4Frame.locator('#ctl00_MainContent_updResults p').first();
      responseCode = await codeElement.textContent(); // Assign to the higher-scoped variable
      console.log(`Extracted Response Code: ${responseCode?.trim()}`);
  } else {
      console.log('OK button not found. Taking screenshot.');
      await page.screenshot({ path: 'E2E_001_OK_missing.png', fullPage: true });
      // Don't fail immediately, maybe it's already gone or handled? But usually we want to click it.
      throw new Error('OK button on confirmation dialog not found.');
  }

  // --- Final Step: Go To Shipment Page ---
  console.log('--- Final Step: Go To Shipment Page ---');

  let shipmentLink = null;
  
  // Retry loop to find the "Go To Shipment Page" link, as its appearance can be slow and flaky.
  console.log('Searching for "Go To Shipment Page" link...');
  for (let attempt = 1; attempt <= 10; attempt++) {
      // Check all frames on the page in each attempt
      for (const frame of page.frames()) {
          // It's possible the frame is detached, so add a guard
          if (frame.isDetached()) continue;

          const link = frame.locator('a:has-text("Go To Shipment Page")');
          if (await link.count() > 0 && await link.isVisible()) {
              console.log(`Found link in frame ${frame.url()} on attempt ${attempt}.`);
              shipmentLink = link;
              break;
          }
      }
      
      if (shipmentLink) break;

      // Also check the main page itself, outside of any frames
      const mainPageLink = page.locator('a:has-text("Go To Shipment Page")');
      if (await mainPageLink.count() > 0 && await mainPageLink.isVisible()) {
          console.log(`Found link on main page on attempt ${attempt}.`);
          shipmentLink = mainPageLink;
          break;
      }

      console.log(`Link not found on attempt ${attempt}. Retrying in 3 seconds...`);
      await page.waitForTimeout(3000);
  }

  if (!shipmentLink) {
      console.log('Failed to find "Go To Shipment Page" link after multiple attempts. Taking screenshot.');
      await page.screenshot({ path: 'E2E_001_ShipmentLink_NotFound.png', fullPage: true });
      throw new Error('Could not find the "Go To Shipment Page" link.');
  }
  
  console.log('Successfully located "Go To Shipment Page" link.');

  // Click and handle new tab
  const [newPage] = await Promise.all([
      page.context().waitForEvent('page'), // Wait for new page (popup)
      shipmentLink.click()
  ]);

  await newPage.waitForLoadState('domcontentloaded');
  console.log(`Navigated to Shipment Page: ${newPage.url()}`);
  
  // Take a screenshot of the new page to verify
  await newPage.screenshot({ path: 'E2E_001_ShipmentPage.png' });

  // --- Interact with Shipment Page ---
  console.log('--- Interacting with Shipment Page ---');
  
  let shipmentFrame = null;
  
  // Retry loop to find the frame containing the button, as it may take time to load
  for (let attempt = 1; attempt <= 5; attempt++) {
    await newPage.waitForTimeout(3000); // Wait before each attempt
    console.log(`Shipment Page frame scan attempt ${attempt}...`);
    const shipmentFrames = newPage.frames();

    for (const frame of shipmentFrames) {
        const saveButton = frame.locator('#ctl00_MainContent_lnxbtnSaveAndValidate');
        // Use isVisible() as it implies the element is in the DOM and rendered
        if (await saveButton.isVisible()) {
            console.log(`Found "Save and Validate" button in frame: ${frame.url()}`);
            shipmentFrame = frame;
            break;
        }
    }
    if (shipmentFrame) {
      break; // Exit retry loop if frame is found
    }
  }

  if (!shipmentFrame) {
      console.log('"Save and Validate" button not found in any frame after multiple attempts. Taking debug screenshot.');
      await newPage.screenshot({ path: 'E2E_001_ShipmentFrame_NotFound.png', fullPage: true });
      throw new Error('Could not find frame containing "Save and Validate" button.');
  }

  // Click "Save and Validate"
  console.log('Clicking "Save and Validate"...');
  await shipmentFrame.locator('#ctl00_MainContent_lnxbtnSaveAndValidate').click();
  console.log('"Save and Validate" clicked.');
  
  // Wait for the "System Messages" tab to appear and click it
  console.log('Waiting for "System Messages" tab...');
  const systemMessagesTab = shipmentFrame.locator('#__tab_tabExportMessages');
  await systemMessagesTab.waitFor({ state: 'visible', timeout: 15000 }); // Wait for tab to be visible
  
  console.log('Clicking "System Messages" tab...');
  await systemMessagesTab.click();
  
  // Wait for the tab panel to update, maybe a short wait is enough
  await shipmentFrame.waitForTimeout(3000); 
  console.log('"System Messages" tab clicked.');

  // Click the "Expand group" button within the System Messages tab
  console.log('Clicking "Expand group" button...');
  const expandButton = shipmentFrame.locator('.rgExpand[title="Expand group"]');
  await expandButton.waitFor({ state: 'visible', timeout: 10000 });
  await expandButton.click();
  console.log('"Expand group" button clicked.');
  await shipmentFrame.waitForTimeout(2000); // Wait for the group to expand

  // Check for and log any system messages
  console.log('Checking for system messages...');
  const messageRows = shipmentFrame.locator('#ctl00_MainContent_tabs_tabExportMessages_rgMessages_GridData tr.rgRow, #ctl00_MainContent_tabs_tabExportMessages_rgMessages_GridData tr.rgAltRow');
  const rowCount = await messageRows.count();

  if (rowCount > 0) {
    console.log(`Found ${rowCount} system message(s):`);
    const allRows = await messageRows.all();
    for (let i = 0; i < allRows.length; i++) {
      const row = allRows[i];
      const cells = row.locator('td');
      const date = await cells.nth(1).textContent();
      const message = await cells.nth(2).textContent();
      console.log(`- Date: ${date?.trim()}, Message: ${message?.trim()}`);
    }
  } else {
    console.log('No system messages found.');
  }

  // --- Add all extracted products to the shipment ---
  console.log(`--- Starting to add ${extractedProductNums.length} products to the shipment ---`);

  // First, navigate to the Details tab
  console.log('Navigating to "Details" tab...');
  const detailsTab = shipmentFrame.locator('#__tab_tabExportDetail');
  await detailsTab.waitFor({ state: 'visible', timeout: 10000 });
  await detailsTab.click();
  console.log('"Details" tab clicked.');
  await shipmentFrame.waitForTimeout(3000); // Wait for tab content to potentially load

  for (const productNum of extractedProductNums) {
    console.log(`--- Adding product: ${productNum} ---`);

    // Click "Add New Record"
    const addNewRecordButton = shipmentFrame.locator('#ctl00_MainContent_tabs_tabExportDetail_rgExportDetail_ctl00_ctl02_ctl00_lnkbtnInsert');
    await addNewRecordButton.waitFor({ state: 'visible', timeout: 10000 });
    await addNewRecordButton.click();
    console.log('"Add New Record" clicked.');

    // The "Add New Record" button opens a modal dialog, likely in a new iframe.
    console.log('Searching for the "Add Product" modal frame...');
    let addProductFrame = null;
    for (let attempt = 1; attempt <= 5; attempt++) {
        await newPage.waitForTimeout(2000); // Wait for modal to appear
        console.log(`Add Product frame scan attempt ${attempt}...`);
        const frames = newPage.frames();
        for (const frame of frames) {
            const productSearchInput = frame.locator('#ctl00_MainContent_txtProductSearch');
            if (await productSearchInput.count() > 0) {
                console.log(`Found "Add Product" modal frame: ${frame.url()}`);
                addProductFrame = frame;
                break;
            }
        }
        if (addProductFrame) break;
    }

    if (!addProductFrame) {
        console.log('Could not find "Add Product" modal frame. Taking debug screenshot.');
        await newPage.screenshot({ path: `E2E_001_AddProductFrame_NotFound_${productNum}.png`, fullPage: true });
        throw new Error(`Could not find modal frame for adding product: ${productNum}`);
    }

    // Type the product number into the search field (character by character for autocomplete)
    console.log(`Typing product number "${productNum}" character by character into the search input...`);
    const searchInput = addProductFrame.locator('#ctl00_MainContent_txtProductSearch');
    await searchInput.waitFor({ state: 'visible', timeout: 10000 });
    await searchInput.type(productNum, { delay: 100 }); // Simulate human typing
    console.log('Product number entered. Waiting for autocomplete suggestion to appear...');

    // Wait for the specific autocomplete suggestion item to be visible within the addProductFrame
    const suggestionItemLocator = addProductFrame.getByText(`SEARCH: ${productNum}`).first();
    await suggestionItemLocator.waitFor({ state: 'visible', timeout: 15000 }); // Increased timeout
    console.log('Autocomplete suggestion is visible.');

    // Use keyboard navigation to select the item - this is often more reliable than clicking.
    console.log('Pressing ArrowDown and Enter to select the the suggestion...');
    // A slight delay might be needed between typing and pressing ArrowDown, or between ArrowDown and Enter
    // Let's try pressing ArrowDown on the searchInput first, then Enter.
    await searchInput.press('ArrowDown');
    await searchInput.press('Enter');

    // The selection should close the dropdown and populate the input.
    // We expect the input's value to START WITH the product number, as it may be appended with other text.
    await expect(searchInput).toHaveValue(new RegExp(`^${productNum}`), { timeout: 5000 });
    console.log(`Successfully selected "${productNum}" via keyboard.`); 
    
    // Click "Save and Close"
    console.log('Clicking "Save and Close"...');
    const saveAndCloseButton = addProductFrame.locator('#ctl00_MainContent_lnxbtnSaveAndClose');
    await saveAndCloseButton.click();
    
    // Wait for the modal to close by waiting for one of its elements to be hidden.
    console.log('Waiting for "Add Product" modal to close...');
    await expect(saveAndCloseButton).toBeHidden({ timeout: 15000 });
    console.log(`Product ${productNum} added and modal closed.`);
    
    // Wait for the newly added product to appear in the main grid
    console.log(`Waiting for product "${productNum}" to appear in the grid...`);
    const productRowLocator = shipmentFrame.locator(`tr.rgRow:has-text("${productNum}"), tr.rgAltRow:has-text("${productNum}")`);
    await productRowLocator.waitFor({ state: 'visible', timeout: 20000 }); // Increased timeout for grid update
    console.log(`Product "${productNum}" found in the grid.`);
  }
  console.log('--- Finished adding all products ---');

  // --- Start adding all extracted company IDs (customers) ---
  console.log(`--- Starting to add ${extractedCompanyIds.length} customer(s) ---`);

  const partyTabs = [
      { tabId: '__tab_tabSeller', inputId: 'ctl00_MainContent_tabs_tabExportParties_parties_tabSeller_rcbSeller_Input', name: 'Customer (Seller)' },
      { tabId: '__tab_tabShipFrom', inputId: 'ctl00_MainContent_tabs_tabExportParties_parties_tabShipFrom_rcbShipFrom_Input', name: 'EndUser (ShipFrom)' },
      { tabId: '__tab_tabBillTo', inputId: 'ctl00_MainContent_tabs_tabExportParties_parties_tabBillTo_rcbBillTo_Input', name: 'Party1 (BillTo)' },
      { tabId: '__tab_tabShipTo', inputId: 'ctl00_MainContent_tabs_tabExportParties_parties_tabShipTo_rcbShipTo_Input', name: 'Party2 (ShipTo)' },
  ];

  // Click "Parties" tab once before the loop
  console.log('Navigating to "Parties" tab...');
  const partiesTab = shipmentFrame.locator('#__tab_tabExportParties');
  await partiesTab.waitFor({ state: 'visible', timeout: 10000 });
  await partiesTab.click();
  console.log('"Parties" tab clicked.');
  await shipmentFrame.waitForTimeout(3000); // Wait for tab content to potentially load


  for (let i = 0; i < extractedCompanyIds.length; i++) {
      if (i >= partyTabs.length) {
          console.warn(`Warning: More company IDs extracted (${extractedCompanyIds.length}) than available party tabs (${partyTabs.length}). Skipping remaining.`);
          break;
      }

      const companyIdToEnter = extractedCompanyIds[i];
      const tabInfo = partyTabs[i];

      console.log(`--- Entering customer: "${companyIdToEnter}" into tab: "${tabInfo.name}" ---`);

      // Click the relevant party sub-tab
      const partySubTab = shipmentFrame.locator(`#${tabInfo.tabId}`);
      await partySubTab.waitFor({ state: 'visible', timeout: 10000 });
      await partySubTab.click();
      console.log(`Sub-tab "${tabInfo.name}" clicked.`);
      await shipmentFrame.waitForTimeout(1000); // Small wait for tab content to become active

      // Locate the input field for the current party tab
      const companyInput = shipmentFrame.locator(`#${tabInfo.inputId}`);
      await companyInput.waitFor({ state: 'visible' });

      // Type the company ID to trigger the autocomplete
      await companyInput.type(companyIdToEnter, { delay: 100 });
      console.log(`Company ID "${companyIdToEnter}" typed. Waiting for autocomplete...`);
      
      // Wait for the autocomplete dropdown to appear (inside shipmentFrame)
      // We use the specific ID for this dropdown to avoid strict mode violations.
      const dropdownId = tabInfo.inputId.replace('_Input', '_DropDown');
      const customerAutocompleteDropdown = shipmentFrame.locator(`#${dropdownId}`);
      await customerAutocompleteDropdown.waitFor({ state: 'visible', timeout: 15000 });
      console.log('Autocomplete dropdown is visible.');

      // Find the specific item in the dropdown
      const suggestionItem = customerAutocompleteDropdown.locator(`li:has-text("${companyIdToEnter}")`).first();
      await suggestionItem.waitFor();
      console.log('Suggestion item found.');

      // Use keyboard to select it
      await companyInput.press('ArrowDown');
      await companyInput.press('Enter');
      
      // Wait for the dropdown to disappear as confirmation of selection
      await expect(customerAutocompleteDropdown).toBeHidden({ timeout: 10000 });
      console.log('Autocomplete dropdown has closed, selection is likely complete.');
      
      await shipmentFrame.waitForTimeout(2000); // Small delay between customers
  }
  console.log('--- Finished adding all customers ---');

  console.log('--- Re-validating after adding all parties ---');

  // Click "Save and Validate" again
  console.log('Clicking "Save and Validate" again...');
  await shipmentFrame.locator('#ctl00_MainContent_lnxbtnSaveAndValidate').click();
  console.log('"Save and Validate" clicked.');
  
  // The page will reload/update. We need to wait for the message tab to be ready.
  // Clicking it again is the most reliable way to ensure we are on it and it's loaded.
  console.log('Waiting for and clicking "System Messages" tab...');
  const finalSystemMessagesTab = shipmentFrame.locator('#__tab_tabExportMessages');
  await finalSystemMessagesTab.waitFor({ state: 'visible', timeout: 15000 });
  await finalSystemMessagesTab.click();
  await shipmentFrame.waitForTimeout(3000); // Wait for content to load

  // Expand the group again to see the new messages
  // There might be a new expand button or the old one might now be a collapse button
  console.log('Looking for group expand/collapse button...');
  const expandButtonAgain = shipmentFrame.locator('.rgExpand[title="Expand group"]');
  if (await expandButtonAgain.isVisible()) {
    console.log('Found "Expand group" button. Clicking it...');
    await expandButtonAgain.click();
    await shipmentFrame.waitForTimeout(2000);
  } else {
    console.log('Group seems to be already expanded (or uses a collapse button).');
  }

  // Check for and log the final system messages
  console.log('Checking for final system messages...');
  const finalMessageRows = shipmentFrame.locator('#ctl00_MainContent_tabs_tabExportMessages_rgMessages_GridData tr.rgRow, #ctl00_MainContent_tabs_tabExportMessages_rgMessages_GridData tr.rgAltRow');
  const finalRowCount = await finalMessageRows.count();

  if (finalRowCount > 0) {
        console.log(`Found ${finalRowCount} final system message(s):`);
        const allFinalRows = await finalMessageRows.all();
        let allHeaderMessages = true;
        let hasTransactionHold = false;
        let hasInformRequirements = false;
        let hasEndUse = false;
        let hasUsageCheck = false;
        let hasEndUserCheck = false;
        let hasGuidelineCheck = false;
        const allMessages: string[] = []; // Array to hold all messages
        
        for (const row of allFinalRows) {
            const cells = row.locator('td');
            const date = await cells.nth(1).textContent();
            const message = await cells.nth(2).textContent();
                  
            const cleanedMessage = message?.trim();
            if (cleanedMessage) {
                allMessages.push(cleanedMessage);
                console.log(`- 日付: ${date?.trim()}, メッセージ: ${cleanedMessage}`);
        
                // 取引審査に関するメッセージの検出
                if (cleanedMessage.includes('取引審査')) {
                    console.log('取引審査に関するメッセージを検出しました。');
                    hasTransactionHold = true;
                }
                    
                // インフォーム要件に関するメッセージの検出
                if (cleanedMessage.includes('インフォーム要件')) {
                    console.log('インフォーム要件に関するメッセージを検出しました。');
                    hasInformRequirements = true;
                }

                // 最終用途に関するメッセージの検出
                if (cleanedMessage.includes('最終用途')) {
                    console.log('最終用途に関するメッセージを検出しました。');
                    hasEndUse = true;
                }

                // 用途チェックリストに関するメッセージの検出
                if (cleanedMessage.includes('用途チェックリスト')) {
                    console.log('用途チェックリストに関するメッセージを検出しました。');
                    hasUsageCheck = true;
                }

                // 需要者チェックリストに関するメッセージの検出
                if (cleanedMessage.includes('需要者チェックリスト')) {
                    console.log('需要者チェックリストに関するメッセージを検出しました。');
                    hasEndUserCheck = true;
                }

                // 明らかガイドラインに関するメッセージの検出
                if (cleanedMessage.includes('明らかガイドライン')) {
                    console.log('明らかガイドラインに関するメッセージを検出しました。');
                    hasGuidelineCheck = true;
                }
        
                // メッセージが "Header" で始まっているかチェック
                if (!cleanedMessage.startsWith('Header')) {
                    console.error(`エラー: "Header" 以外のメッセージが見つかりました: "${cleanedMessage}"`);
                    allHeaderMessages = false;
                }
            }
        }
        
        // レスポンスコード '01010100' の検証
        if (responseCode?.trim() === '01010100') {
            console.log("レスポンスコードは '01010100' です。特定のエラーメッセージが存在するか検証します...");
            const requiredKeywords = [
                '取引審査結果',
                '最終用途',
                'インフォーム要件',
                '需要者チェックリスト',
                '明らかガイドライン',
                '用途チェックリスト'
            ];
                  
            const allMessagesString = allMessages.join(' || ');
            const missingKeywords: string[] = [];
        
            for (const keyword of requiredKeywords) {
                if (!allMessagesString.includes(keyword)) {
                    missingKeywords.push(keyword);
                }
            }
        
            if (missingKeywords.length > 0) {
                throw new Error(`検証失敗: レスポンスコード '01010100' ですが、システムメッセージに必須キーワードが見つかりません: ${missingKeywords.join(', ')}`);
            } else {
                console.log("検証成功: レスポンスコード '01010100' に対応する全ての必須キーワードが見つかりました。");
            }
        }
        
        // --- 未入力項目の修正処理 ---
        const needsCorrection = hasTransactionHold || hasInformRequirements || hasEndUse || hasUsageCheck || hasEndUserCheck || hasGuidelineCheck;
        
        if (needsCorrection) {
            console.log('未入力項目を検出しました。Headerタブへ移動して修正します...');
            // Headerタブへ移動 (共通処理)
            const headerTab = shipmentFrame.locator('#__tab_tabExportHeader');
            await headerTab.waitFor({ state: 'visible', timeout: 10000 });
            await headerTab.click();
            console.log('Headerタブをクリックしました。');
            await shipmentFrame.waitForTimeout(3000);
        }
        
        // 1. 取引審査結果の入力
        if (hasTransactionHold) {
            console.log('「取引審査結果」の入力を開始します...');
            const judgementContainer = shipmentFrame.locator('#s2id_drpJudgement_TTM');
            if (await judgementContainer.isVisible()) {
                const judgementDropdown = judgementContainer.locator('.select2-choice');
                await judgementDropdown.click();
                console.log('取引審査結果ドロップダウンをクリックしました。');

                // ランダム選択: "取引可/OK" or "取引不可/NG"
                const judgementOptions = ['取引可/OK', '取引不可/NG'];
                const randomJudgement = judgementOptions[Math.floor(Math.random() * judgementOptions.length)];
                console.log(`ランダム選択された取引審査結果: "${randomJudgement}"`);
                const optionToSelect = shipmentFrame.locator('div.select2-drop ul.select2-results li').filter({ hasText: randomJudgement }).first();
                await optionToSelect.waitFor({ state: 'visible', timeout: 5000 }).catch(() => console.log(`選択肢 "${randomJudgement}" がすぐに見つかりませんでした。`));

                if (await optionToSelect.isVisible()) {
                    await optionToSelect.click();
                    console.log(`選択肢 "${randomJudgement}" を選択しました。`);
                    await expect(judgementDropdown.locator('.select2-chosen')).toHaveText(randomJudgement, { timeout: 5000 });
                    console.log(`検証: 取引審査結果が "${randomJudgement}" に設定されました。`);
                } else {
                    throw new Error(`取引審査ドロップダウンに選択肢 "${randomJudgement}" が見つかりません。`);
                }
            } else {
                throw new Error('取引審査 (Judgement) ドロップダウンが見つかりません (#s2id_drpJudgement_TTM)。');
            }
        }
              
        // 2. インフォーム要件の入力
        if (hasInformRequirements) {
            console.log('「インフォーム要件」の入力を開始します...');
            const informContainer = shipmentFrame.locator('#s2id_drpInformRequirementDetermination_TTM');

            if (await informContainer.isVisible()) {
                const informDropdown = informContainer.locator('.select2-choice');
                await informDropdown.click();
                console.log('インフォーム要件ドロップダウンをクリックしました。');

                // ランダム選択: "Y-はい" or "N-いいえ"
                const informOptions = ['N-いいえ', 'Y-はい'];
                const randomInform = informOptions[Math.floor(Math.random() * informOptions.length)];
                console.log(`ランダム選択されたインフォーム要件: "${randomInform}"`);
                const optionToSelect = shipmentFrame.locator('div.select2-drop ul.select2-results li').filter({ hasText: randomInform }).first();
                await optionToSelect.waitFor({ state: 'visible', timeout: 5000 }).catch(() => console.log(`選択肢 "${randomInform}" がすぐに見つかりませんでした。`));

                if (await optionToSelect.isVisible()) {
                    await optionToSelect.click();
                    console.log(`選択肢 "${randomInform}" を選択しました。`);
                    await expect(informDropdown.locator('.select2-chosen')).toHaveText(randomInform, { timeout: 5000 });
                    console.log(`検証: インフォーム要件が "${randomInform}" に設定されました。`);
                } else {
                    throw new Error(`インフォーム要件ドロップダウンに必須の選択肢 "N-いいえ" が見つかりません。`);
                }
            } else {
                throw new Error('インフォーム要件ドロップダウンが見つかりません (#s2id_drpInformRequirementDetermination_TTM)。');
            }

        }           
        
                
        // 3. 最終用途・使用場所の入力
        if (hasEndUse) {
            console.log('「最終用途・使用場所」の入力を開始します...');
            const endUseInput = shipmentFrame.locator('#txtEndUseDestination_TTM');
            if (await endUseInput.isVisible()) {
                const inputText = "EndUse: Test Purpose / Place: Test Place";
                await endUseInput.fill(inputText);
                console.log(`最終用途欄に入力しました: "${inputText}"`);
                await expect(endUseInput).toHaveValue(inputText);
            } else {
                throw new Error('最終用途入力フィールドが見つかりません (#txtEndUseDestination_TTM)。');
            }
        }
        
        
        
        // 4. 用途チェックリストの入力
        if (hasUsageCheck) {
            console.log('「用途チェックリスト」の入力を開始します...');
            const usageContainer = shipmentFrame.locator('#s2id_drpUsageCheckResult_TTM');
            if (await usageContainer.isVisible()) {
                const usageDropdown = usageContainer.locator('.select2-choice');
                await usageDropdown.click();
                console.log('用途チェックリストドロップダウンをクリックしました。');

                const options = ['1-未実施', '2-懸念なし', '3-懸念あり'];
                const randomOption = options[Math.floor(Math.random() * options.length)];
                console.log(`ランダム選択された用途チェックリスト: "${randomOption}"`);

                const optionToSelect = shipmentFrame.locator('div.select2-drop ul.select2-results li').filter({ hasText: randomOption }).first();
                await optionToSelect.waitFor({ state: 'visible', timeout: 5000 });

                if (await optionToSelect.isVisible()) {
                    await optionToSelect.click();
                    console.log(`選択肢 "${randomOption}" を選択しました。`);
                    await expect(usageDropdown.locator('.select2-chosen')).toHaveText(randomOption, { timeout: 5000 });
                } else {
                    throw new Error(`用途チェックリストドロップダウンに選択肢 "${randomOption}" が見つかりません。`);
                }
            } else {
                throw new Error('用途チェックリストドロップダウンが見つかりません (#s2id_drpUsageCheckResult_TTM)。');
            }
        }
        
        
        
        // 5. 需要者チェックリストの入力
        if (hasEndUserCheck) {
            console.log('「需要者チェックリスト」の入力を開始します...');
            const endUserContainer = shipmentFrame.locator('#s2id_drpEndUserCheckResult_TTM');
            if (await endUserContainer.isVisible()) {
                const endUserDropdown = endUserContainer.locator('.select2-choice');
                await endUserDropdown.click();
                console.log('需要者チェックリストドロップダウンをクリックしました。');

                const options = ['1-未実施', '2-懸念なし', '3-懸念あり'];
                const randomOption = options[Math.floor(Math.random() * options.length)];
                console.log(`ランダム選択された需要者チェックリスト: "${randomOption}"`);

                const optionToSelect = shipmentFrame.locator('div.select2-drop ul.select2-results li').filter({ hasText: randomOption }).first();
                await optionToSelect.waitFor({ state: 'visible', timeout: 5000 });

                if (await optionToSelect.isVisible()) {
                    await optionToSelect.click();
                    console.log(`選択肢 "${randomOption}" を選択しました。`);
                    await expect(endUserDropdown.locator('.select2-chosen')).toHaveText(randomOption, { timeout: 5000 });
                } else {
                    throw new Error(`需要者チェックリストドロップダウンに選択肢 "${randomOption}" が見つかりません。`);
                }
            } else {
                throw new Error('需要者チェックリストドロップダウンが見つかりません (#s2id_drpEndUserCheckResult_TTM)。');
            }
        }
        
        
        // 6. 明らかガイドラインの入力
        if (hasGuidelineCheck) {
            console.log('「明らかガイドライン」の入力を開始します...');
            const guidelineContainer = shipmentFrame.locator('#s2id_drpGuidelineResult_TTM');
            if (await guidelineContainer.isVisible()) {
                const guidelineDropdown = guidelineContainer.locator('.select2-choice');
                await guidelineDropdown.click();
                console.log('明らかガイドラインドロップダウンをクリックしました。');

                const options = ['1-未実施', '2-懸念なし', '3-懸念あり'];
                const randomOption = options[Math.floor(Math.random() * options.length)];
                console.log(`ランダム選択された明らかガイドライン: "${randomOption}"`);

                const optionToSelect = shipmentFrame.locator('div.select2-drop ul.select2-results li').filter({ hasText: randomOption }).first();
                await optionToSelect.waitFor({ state: 'visible', timeout: 5000 });

                if (await optionToSelect.isVisible()) {
                    await optionToSelect.click();
                    console.log(`選択肢 "${randomOption}" を選択しました。`);
                    await expect(guidelineDropdown.locator('.select2-chosen')).toHaveText(randomOption, { timeout: 5000 });
                } else {
                    throw new Error(`明らかガイドラインドロップダウンに選択肢 "${randomOption}" が見つかりません。`);
                }
            } else {
                throw new Error('明らかガイドラインドロップダウンが見つかりません (#s2id_drpGuidelineResult_TTM)。');
            }
        }
        
                
        
        // 7. まとめて保存と検証
        if (needsCorrection) {
            console.log('--- 修正完了後、再度保存して検証します ---');
            const saveButton = shipmentFrame.locator('#ctl00_MainContent_lnxbtnSaveAndValidate');
            await saveButton.click();
            console.log('「保存と検証」ボタンをクリックしました。');

            // システムメッセージタブが表示されるのを待機
            const sysMsgTab = shipmentFrame.locator('#__tab_tabExportMessages');
            await sysMsgTab.waitFor({ state: 'visible', timeout: 20000 });
            await sysMsgTab.click();
            await shipmentFrame.waitForTimeout(3000);

            // メッセージを再確認
            const finalRows = shipmentFrame.locator('#ctl00_MainContent_tabs_tabExportMessages_rgMessages_GridData tr.rgRow, #ctl00_MainContent_tabs_tabExportMessages_rgMessages_GridData tr.rgAltRow');
            const finalCount = await finalRows.count();
            console.log(`再検証時のメッセージ数: ${finalCount}`);

            if (finalCount > 0) {
                const finalAllRows = await finalRows.all();
                let remainingErrors = false;
                for (const row of finalAllRows) {
                    const msg = await row.locator('td').nth(2).textContent();
                    console.log(`再検証メッセージ: ${msg?.trim()}`);

                    // エラーメッセージが残っていないかチェック
                    if (hasTransactionHold && msg?.includes('取引審査')) remainingErrors = true;
                    if (hasInformRequirements && msg?.includes('インフォーム要件')) remainingErrors = true;
                    if (hasEndUse && msg?.includes('最終用途')) remainingErrors = true;
                    if (hasUsageCheck && msg?.includes('用途チェックリスト')) remainingErrors = true;
                    if (hasEndUserCheck && msg?.includes('需要者チェックリスト')) remainingErrors = true;
                    if (hasGuidelineCheck && msg?.includes('明らかガイドライン')) remainingErrors = true;
                }

                if (remainingErrors) {
                    throw new Error('検証失敗: 入力後もエラーメッセージが残っています。');
                } else {
                    console.log('検証成功: 対象のエラーメッセージは解消されました。');
                }
            } else {
                console.log('検証成功: システムメッセージはありません。');
            }

        }  

        // 最終的な "Header" チェックのアサーション
        if (!allHeaderMessages) {
        throw new Error('Not all final system messages were "Header" related.');
        } else {
        console.log('Validation successful: All final system messages are "Header" related.');
        }

    // Conditional check for responseCode '01010100'
    if (responseCode?.trim() === '01010100') {
        console.log("Response code is '01010100'. Verifying specific error messages...");
        const requiredKeywords = [
            '取引審査結果',
            '最終用途',
            'インフォーム要件',
            '需要者チェックリスト',
            '明らかガイドライン',
            '用途チェックリスト'
        ];
      
        const allMessagesString = allMessages.join(' || '); // Join all messages into one string for easy searching
        const missingKeywords: string[] = [];

        for (const keyword of requiredKeywords) {
            if (!allMessagesString.includes(keyword)) {
            missingKeywords.push(keyword);
            }
        }

        if (missingKeywords.length > 0) {
            throw new Error(`Validation failed for response code '01010100'. Missing required keywords in system messages: ${missingKeywords.join(', ')}`);
        } else {
            console.log("Validation successful: All required keywords found for response code '01010100'.");
        }                                         
    }
} else {
    console.log('No final system messages found. Validation successful!');
}

});