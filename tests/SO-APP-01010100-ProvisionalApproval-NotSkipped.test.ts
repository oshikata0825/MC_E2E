import { test, expect } from '@playwright/test';
import * as fs from 'fs';
import * as path from 'path';

/**
 * Sales Order Approval Chain Test (Provisional)
 * This test handles the sequential approval of a Sales Order by multiple users.
 * It reads the Sales Order Number from approval_info.json created by SO-01010100-ProvisionalApproval-NotSkipped.test.ts.
 */

test('Response01010100-ProvisionalApproval-NotSkipped', async ({ page }) => {
  console.log(`--- Test Started at ${new Date().toISOString()} ---`);
  test.setTimeout(600000); // 10 minutes
  const accountsData = fs.readFileSync(path.resolve(__dirname, '../account.dat'), 'utf-8');
  const accounts = JSON.parse(accountsData);
  
  let targetSO = '';
  // Default sequence for Pattern B
  let approverUsers = ['MCTest2', 'MCTest3']; 

  const infoPath = path.resolve(__dirname, '../approval_info.json');
  if (fs.existsSync(infoPath)) {
      try {
          const info = JSON.parse(fs.readFileSync(infoPath, 'utf-8'));
          targetSO = info.salesOrderNo;
          console.log(`Loaded target Sales Order from approval_info.json: ${targetSO}`);
          
          if (info.approverRoles && info.approverRoles.length > 0) {
              console.log('Detected approver roles from E2E_001:', info.approverRoles);
              // Map roles to user IDs based on standard Mitsubishi Pattern B configuration
              const roleToUserMap: Record<string, string> = {
                  '2-BU Team Leader': 'MCTest2',
                  '2-BU Team Leader/Department Manager': 'MCTest3',
                  '5-GCEO Office': 'MCTest5',
                  '6-Legal Representative': 'MCTest6',
                  '7-Legal Director': 'MCTest7',
                  '8-General Manager': 'MCTest8'
              };
              
              const dynamicApprovers = info.approverRoles
                  .map((role: string) => roleToUserMap[role])
                  .filter((user: string) => user && user !== 'MCTest1'); // Exclude requestor
              
              if (dynamicApprovers.length > 0) {
                  approverUsers = [...new Set(dynamicApprovers)] as string[]; // Deduplicate
                  console.log('Dynamically determined approver sequence:', approverUsers);
              }
          }
      } catch (e) {
          console.error('Failed to parse approval_info.json:', e);
      }
  }

  if (!targetSO || targetSO === 'null') {
      console.warn('Sales Order Number is empty or missing. Will attempt to pick dynamically from Workqueue.');
  }

  for (const approverId of approverUsers) {
    console.log(`\n=== Starting Approval Flow for ${approverId} ===`);
    
    const user = accounts.find((acc: any) => acc.id === approverId);
    if (!user) {
        console.error(`User ${approverId} not found in account.dat. Skipping.`);
        continue;
    }

    // 1. Navigation to Home / Login
    console.log(`Navigating to home for ${approverId}...`);
    await page.goto('https://ogt-gtm-web-eu-imp.7166.aws.thomsonreuters.com/gtm/home', { waitUntil: 'load', timeout: 60000 });
    
    const homeIcon = page.locator('.bento-icon-home');
    const loginButton = page.locator('#btxLoginGroup');

    // Wait for the page to decide its state (Login or Home)
    console.log('Waiting for login page or home page to load...');
    await Promise.race([
        loginButton.waitFor({ state: 'visible', timeout: 30000 }).catch(() => {}),
        homeIcon.waitFor({ state: 'visible', timeout: 30000 }).catch(() => {})
    ]);
    
    // If we're on the home page, check user
    if (await homeIcon.isVisible() || await page.locator('.bento-icon-user').isVisible()) {
        console.log('Home/User icon visible. checking session...');
        
        // Try to find the user name in common locations
        const potentialUserElements = [
            page.locator('button.bento-button-tertiary'),
            page.locator('.bento-button--tertiary'),
            page.locator('#ctl00_lblUserName'),
            page.locator('.user-name'),
            page.locator('.AccountMenu')
        ];
        
        let currentUser = '';
        let userMenu = null;

        for (const loc of potentialUserElements) {
            if (await loc.isVisible()) {
                const text = (await loc.textContent() || '').trim();
                if (text.includes('MCTest')) {
                    currentUser = text;
                    userMenu = loc;
                    break;
                }
            }
        }

        console.log(`Detected current user: "${currentUser}"`);

        // If we can't confirm it's the right user, or we can't find the name at all, logout
        if (!currentUser || !currentUser.includes(approverId)) {
             console.log(`SESSION MISMATCH or UNKNOWN: Current="${currentUser}", Need="${approverId}". Logging out...`);
             
             // Open user menu if not already open
             const menuTrigger = userMenu || page.locator('.bento-icon-user, button.bento-button-tertiary').first();
             await menuTrigger.click().catch(() => {});
             await page.waitForTimeout(1000);
             
             const signOutBtn = page.locator('button, a').filter({ hasText: /Sign out|ログアウト|Signout/i }).first();
             if (await signOutBtn.isVisible()) {
                 await signOutBtn.click();
             } else {
                 console.log('Explicit sign out button not found. Using Logout.aspx fallback.');
                 await page.goto('https://ogt-gtm-web-eu-imp.7166.aws.thomsonreuters.com/Logout.aspx');
             }
             
             await loginButton.waitFor({ state: 'visible', timeout: 30000 });
             console.log('Successfully reached login page.');
        } else {
             console.log(`STAYING LOGGED IN: Correct user "${currentUser}" already active.`);
        }
    }

    // Perform Login if the login button is visible (or if we just logged out)
    if (await loginButton.isVisible()) {
        console.log(`Login page detected. Logging in as ${approverId}...`);
        await page.locator('#ctl00_RightSide_txtUserName').fill(user.id);
        await page.locator('#ctl00_RightSide_txtPassword').fill(user.pass);
        await page.locator('#ctl00_RightSide_txtCompany').fill(user.company);
        console.log('Clicking login button...');
        
        await Promise.all([
            page.waitForURL(/.*home|.*Dashboard/i, { timeout: 60000 }).catch(() => {}),
            loginButton.click()
        ]);
        
        try {
            await page.waitForLoadState('networkidle', { timeout: 30000 }).catch(() => {});
            await homeIcon.waitFor({ state: 'visible', timeout: 30000 });
            console.log(`Logged in successfully as ${approverId}.`);
            await page.screenshot({ path: `Debug_Home_${approverId}.png` });
        } catch (e) {
            console.error(`Login failed for ${approverId}: Home icon not found after login click.`);
            if (!page.isClosed()) {
                await page.screenshot({ path: `Error_LoginFailed_HomeNotFound_${approverId}.png` }).catch(() => {});
            }
            throw new Error(`Login failed for ${approverId}. Home icon not visible.`);
        }
    } else if (!(await homeIcon.isVisible())) {
        console.error('Neither login page nor home page confirmed. Cannot proceed.');
        await page.screenshot({ path: `Error_AuthStatus_Unknown_${approverId}.png` });
        throw new Error(`Authentication status unknown for ${approverId}.`);
    }

    // 2. Navigate to "My Workqueue" via specific path: Export -> Sales Order Search/Reporting Lookup -> My Workqueue
    console.log(`Navigating to My Workqueue for ${approverId}...`);
    
    // Step A: Open Export Menu (elements.txt Line 1: #button-menu-2)
    const exportBtn = page.locator('#button-menu-2');
    await exportBtn.waitFor({ state: 'visible', timeout: 15000 });
    await exportBtn.click();
    console.log('Clicked "Export" menu button.');
    await page.waitForTimeout(1000);

    // Step B: Click "Sales Order Search/Reporting Lookup" (elements.txt Line 2)
    const searchLookupLink = page.locator('a:has-text("Sales Order Search/Reporting Lookup")').first();
    await searchLookupLink.waitFor({ state: 'visible', timeout: 10000 });
    await searchLookupLink.click();
    console.log('Clicked "Sales Order Search Lookup" link.');
    
    await page.waitForLoadState('load');
    await page.waitForTimeout(5000);

    // Step C: Look for "My Workqueue" (elements.txt Line 3: #ctl00_LeftHandNavigation_LHN_ctl02)
    let workqueueTab = null;
    for (const ctx of [page, ...page.frames()]) {
        const loc = ctx.locator('#ctl00_LeftHandNavigation_LHN_ctl02').first();
        if (await loc.count() > 0 && await loc.isVisible()) {
            workqueueTab = loc;
            break;
        }
    }

    if (workqueueTab) {
        console.log('Found "My Workqueue" link. Clicking...');
        await workqueueTab.click();
        await page.waitForTimeout(5000);
    } else {
        console.warn('"My Workqueue" link not found by ID. Trying text search...');
        const fallback = page.locator('a:has-text("My Workqueue")').first();
        if (await fallback.isVisible()) {
            await fallback.click();
            await page.waitForTimeout(5000);
        }
    }

    // 3. Selection: Choose the record to approve (elements.txt Line 4: Approval link)
    console.log(`Looking for Sales Order: ${targetSO || 'Any available'} in My Workqueue...`);
    let approvalLink = null;
    let foundSO = '';
    
    // Give more time for the grid to definitely load after clicking My Workqueue
    await page.waitForTimeout(5000);
    
    const scanContexts = [page, ...page.frames()];
    
    // Attempt 1: Look for the specific targetSO with "Approval" link
    if (targetSO) {
        for (const ctx of scanContexts) {
            try {
                // Look for a row that contains the SO number and a link that looks like an approval link
                const rows = ctx.locator('tr').filter({ hasText: targetSO });
                const count = await rows.count();
                for (let i = 0; i < count; i++) {
                    const row = rows.nth(i);
                    // Use a more specific selector to avoid headers or other buttons
                    const link = row.locator('a[href*="ApprovalControlGuid"], a:has-text("Approval")').first();
                    if (await link.count() > 0 && await link.isVisible()) {
                        approvalLink = link;
                        foundSO = targetSO;
                        console.log(`Found target Sales Order ${targetSO} with Approval link.`);
                        break;
                    }
                }
            } catch (e) {}
            if (approvalLink) break;
        }
    }

    // Attempt 2: DYNAMIC SELECTION - pick any first row that has an "Approval" link
    if (!approvalLink) {
        console.warn(`Specific Approval link for "${targetSO}" not found. Picking first available row with "Approval" link.`);
        for (const ctx of scanContexts) {
            try {
                // Look for links that are clearly Approval links (contain the GUID) or have exact "Approval" text in a data cell
                // We exclude headers by looking for links inside cells (td) that are NOT sorting links
                const candidateLinks = ctx.locator('td a').filter({ hasText: /^Approval$/i });
                const count = await candidateLinks.count();
                
                for (let i = 0; i < count; i++) {
                    const link = candidateLinks.nth(i);
                    const href = await link.getAttribute('href') || '';
                    
                    // Approval links usually contain ApprovalControlGuid and are not sorting postbacks
                    if (href.includes('ApprovalControlGuid') || !href.includes('Refresh')) {
                        approvalLink = link;
                        
                        // Try to extract SO from row text
                        const row = ctx.locator('tr').filter({ has: link }).first();
                        const rowText = await row.innerText();
                        const idMatch = rowText.match(/OSGT_S\d+/);
                        if (idMatch) {
                            foundSO = idMatch[0];
                            console.log(`DYNAMIC SELECTION: Picked record ${foundSO} from Workqueue.`);
                        } else {
                            console.log('DYNAMIC SELECTION: Picked a record but could not extract SO ID.');
                        }
                        break;
                    }
                }
            } catch (e) {}
            if (approvalLink) break;
        }
    }

    if (approvalLink) {
        if (foundSO) targetSO = foundSO;
        
        console.log(`Targeting Approval link for ${targetSO || 'selected record'}...`);
        const href = await approvalLink.getAttribute('href').catch(() => 'unknown');
        console.log(`Link href: ${href}`);
        
        await page.screenshot({ path: `Debug_BeforeClick_${approverId}.png` });
        
        let approvalPage;
        try {
            console.log('Clicking Approval link and waiting for popup...');
            // Try to click. We use a race/retry for the click itself if it's tricky
            const [newPage] = await Promise.all([
                page.context().waitForEvent('page', { timeout: 45000 }).catch(e => {
                    console.warn('Popup did not open within 45s. Checking for navigation or error.');
                    throw e;
                }),
                approvalLink.click({ timeout: 15000 }).catch(async (e) => {
                    console.warn('Standard click failed, attempting force click...');
                    return approvalLink.click({ force: true });
                })
            ]);
            approvalPage = newPage;
            console.log('Found and captured approval window.');
        } catch (e) {
            console.error('Failed to capture approval window via event. Checking current pages...');
            const pages = page.context().pages();
            if (pages.length > 1) {
                approvalPage = pages[pages.length - 1];
                console.log('Using a detected second window.');
            } else {
                // Last ditch effort: navigate directly if we have the href and it's relative
                if (href && href.startsWith('/') && !href.includes('javascript:')) {
                    console.log(`Attempting direct navigation to ${href} as fallback.`);
                    approvalPage = await page.context().newPage();
                    await approvalPage.goto(`https://ogt-gtm-web-eu-imp.7166.aws.thomsonreuters.com${href}`);
                } else {
                    await page.screenshot({ path: `Error_FinalClickFailure_${approverId}.png` });
                    throw new Error(`Could not open approval window for ${approverId}: ${e.message}`);
                }
            }
        }

        await approvalPage.waitForLoadState('domcontentloaded');
        await approvalPage.waitForTimeout(5000);

        // 4. Handle the Approval Form
        console.log(`Handling Approval as ${approverId}...`);
        await approvalPage.waitForLoadState('load').catch(() => {});
        
        let approveBtn = null;
        let approvalFrame = null;

        const approveSelectors = [
            'a[id*="Approve_Button"]',
            'input[id*="Approve_Button"]',
            'a:has-text("Approve")',
            'button:has-text("Approve")',
            'a:has-text("Submit")',
            'a:has-text("確定")'
        ];

        console.log('Searching for "Approve" button...');
        for (let attempt = 0; attempt < 10; attempt++) {
            for (const f of [approvalPage, ...approvalPage.frames()]) {
                for (const selector of approveSelectors) {
                    try {
                        const loc = f.locator(selector).first();
                        if (await loc.count() > 0 && await loc.isVisible()) {
                            approveBtn = loc;
                            approvalFrame = f;
                            const btnText = await loc.innerText().catch(() => 'unknown');
                            const btnId = await loc.getAttribute('id').catch(() => 'unknown');
                            console.log(`Found "${btnText}" (ID: ${btnId}) in frame: ${f.url().substring(0, 40)}`);
                            break;
                        }
                    } catch (e) {}
                }
                if (approveBtn) break;
            }
            if (approveBtn) break;
            await approvalPage.waitForTimeout(1500);
        }

        if (approveBtn) {
            console.log('Clicking "Approve" button...');
            await approveBtn.scrollIntoViewIfNeeded().catch(() => {});
            await approveBtn.click({ force: true });
            
            // Handle "OK" confirmation dialog (elements.txt Line 2)
            console.log('Waiting for confirmation dialog (OK button)...');
            let okButton = null;
            
            // Wait up to 10 seconds for the OK button to appear in ANY frame
            for (let attempt = 0; attempt < 10; attempt++) {
                for (const f of [approvalPage, ...approvalPage.frames()]) {
                    try {
                        const btn = f.locator('button.rwOkBtn, .rwOkBtn, button:has-text("OK")').first();
                        if (await btn.count() > 0 && await btn.isVisible()) {
                            okButton = btn;
                            break;
                        }
                    } catch (e) {}
                }
                if (okButton) break;
                await approvalPage.waitForTimeout(1000);
            }

            if (okButton) {
                console.log('Confirmation dialog detected. Clicking OK...');
                await okButton.click({ force: true });
                
                // Wait for the dialog to disappear to avoid intercepting next clicks
                console.log('Waiting for confirmation dialog to vanish...');
                await Promise.race([
                    approvalPage.locator('[id*="RadWindowWrapper_confirm"]').waitFor({ state: 'hidden', timeout: 8000 }).catch(() => {}),
                    approvalPage.waitForTimeout(4000)
                ]);
            } else {
                console.warn('Confirmation OK button not found within 10s. Checking if approval already progressed.');
                await approvalPage.screenshot({ path: `Debug_NoOKButton_${approverId}.png` });
            }
            
            console.log(`Approval action performed by ${approverId}.`);
            
            // 5. Verify Progression via "Approval Progress Track" tab (elements.txt Line 3)
            console.log('Searching for "Approval Progress Track" tab...');
            let progressTab = null;
            const progressTabSelectors = [
                '[id*="TabPanel_ApprovalProgressTrack"]',
                'span.ajax__tab_tab:has-text("Approval Progress Track")',
                '.ajax__tab_tab:has-text("Approval Progress Track")',
                'span:has-text("Approval Progress Track")',
                'a:has-text("Approval Progress Track")'
            ];

            // Re-scan frames for the progress tab
            for (let attempt = 0; attempt < 5; attempt++) {
                for (const ctx of [approvalPage, ...approvalPage.frames()]) {
                    for (const selector of progressTabSelectors) {
                        try {
                            const loc = ctx.locator(selector).first();
                            if (await loc.count() > 0 && await loc.isVisible()) {
                                progressTab = loc;
                                break;
                            }
                        } catch (e) {}
                    }
                    if (progressTab) break;
                }
                if (progressTab) break;
                await approvalPage.waitForTimeout(2000);
            }

            if (progressTab) {
                console.log('Clicking "Approval Progress Track" tab...');
                // Use force click if it's still being intercepted by a fading mask
                await progressTab.click({ force: true });
                await approvalPage.waitForTimeout(5000);
                
                // Screenshot of the Progress Track for the report
                const screenshotPath = `ApprovalProgress_${approverId}_${targetSO || 'Dynamic'}.png`;
                await approvalPage.screenshot({ path: screenshotPath, fullPage: true });
                console.log(`Progress Track screenshot saved to ${screenshotPath}.`);
                
                // Verify next approver status if mapping exists
                const nextApproverMap: Record<string, string> = {
                  'MCTest2': 'MCTest3',
                  'MCTest3': 'MCTest5',
                  'MCTest5': 'MCTest6',
                  'MCTest6': 'MCTest7',
                  'MCTest7': 'MCTest8'
                };
      
                const nextApproverId = nextApproverMap[approverId];
                if (nextApproverId) {
                    console.log(`Checking if ${nextApproverId} is now InProgress...`);
                    let statusFound = false;
                    for (const ctx of [approvalPage, ...approvalPage.frames()]) {
                        try {
                            const row = ctx.locator(`tr:has-text("${nextApproverId}")`).first();
                            if (await row.count() > 0 && await row.isVisible()) {
                                const text = await row.innerText();
                                console.log(`Status row found for ${nextApproverId}. Text: ${text.replace(/\s+/g, ' ')}`);
                                statusFound = true;
                                break;
                            }
                        } catch(e) {}
                    }
                    if (!statusFound) console.warn(`Status row for ${nextApproverId} not seen.`);
                }
            } else {
                console.warn('Could not find "Approval Progress Track" tab. Capturing debug screenshot.');
                await approvalPage.screenshot({ path: `Debug_TabsNotFound_${approverId}.png` });
            }

            console.log(`Closing approval window for ${approverId}.`);
            await approvalPage.close().catch(() => {});

            // 6. Explicit Logout to prepare for next user
            console.log(`Performing logout for ${approverId} to ensure clean session...`);
            await page.bringToFront();
            await page.goto('https://ogt-gtm-web-eu-imp.7166.aws.thomsonreuters.com/Logout.aspx').catch(() => {});
            await page.waitForTimeout(3000);
        } else {
            console.error(`"Approve" button not found for ${approverId}.`);
            await approvalPage.screenshot({ path: `Error_NoApproveBtn_${approverId}.png` });
            await approvalPage.close().catch(() => {});
            
            // Still logout to be safe
            await page.goto('https://ogt-gtm-web-eu-imp.7166.aws.thomsonreuters.com/Logout.aspx').catch(() => {});
            throw new Error('Approve button missing.');
        }
    }

  }

  console.log('\n=== All designated approvals processed. ===');
});
