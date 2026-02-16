import { test, expect } from '@playwright/test';
import fs from 'fs';

test('GC_data_creation', async ({ page }) => {
  test.setTimeout(100000);
  await page.goto('https://ogt-gtm-web-eu-imp.7166.aws.thomsonreuters.com/Logon.aspx?ReturnUrl=%2fgtm%2fhome');
  await page.getByRole('textbox', { name: 'Company' }).click();
  await page.getByRole('textbox', { name: 'Company' }).fill('Mitsubishi Corporation');
  await page.getByRole('textbox', { name: 'Username' }).click();
  await page.getByRole('textbox', { name: 'Username' }).fill('TOshikata');
  await page.getByRole('textbox', { name: 'Password' }).click();
  await page.getByRole('textbox', { name: 'Password' }).fill('5D9dtEi*Qaivft');
  await page.getByRole('link', { name: 'Login' }).click();
  // await page.getByRole('button', { name: 'Tomohiro Oshikata ' }).click();
  // await page.getByLabel('言語').selectOption('0: Object');
  await page.getByRole('button', { name: 'Classification' }).click();
  await page.getByRole('menuitem', { name: 'New Classifications' }).click();
  await page.locator('#JudgmentCategory > .root > slot > .indicator').click();
  await page.getByRole('option', { name: '貨物' }).click();
  const productName = `Test product ${Date.now()}`;
  await page.getByRole('textbox', { name: 'ProductNameEN' }).fill(productName);
  await page.getByRole('textbox', { name: 'ProductNameJP' }).fill(productName);
  await page.locator('#GHReportFlag > .root > slot > .indicator > slot > saf-icon > .fa-light').click();
  await page.getByRole('option', { name: 'Y-はい' }).click();
  await page.getByRole('textbox', { name: 'ManufacturerName' }).fill('Manufacturer B002');
  await page.locator('#GroupOrgCode > .root > slot > .indicator').click();
  await page.getByRole('combobox', { name: 'GroupOrgCode' }).fill('x');
  await page.getByRole('option', { name: 'X:汤森路透' }).click();
  await page.locator('#DivisionOrgCode > .root > slot > .indicator > slot > saf-icon > .fa-light').click();
  await page.getByRole('combobox', { name: 'DivisionOrgCode' }).fill('x1');
  await page.getByRole('option', { name: 'X1:实施' }).click();
  await page.locator('#BUCode > .root > slot > .indicator > slot > saf-icon > .fa-light').click();
  await page.getByRole('combobox', { name: 'BUCode' }).fill('001');
  await page.getByRole('option', { name: ':实施事業部' }).click();
  await page.getByRole('button', { name: 'Save' }).first().click();

  // elements.txt に依存せず、直接リクエスト番号を取得する
  // 一般的な valuedisplay クラス、または ID 指定を試みる
  const requestNumberLocator = page.locator('span.valuedisplay, #lblRequestNumber, .RequestNumberValue').first();
  await requestNumberLocator.waitFor({ state: 'visible', timeout: 10000 });
  const requestNumber = await requestNumberLocator.innerText();
  
  console.log(`Retrieved Request Number: ${requestNumber}`);
  await page.getByRole('button', { name: 'Classification' }).click();
  await page.getByRole('menuitem', { name: 'GC Search/Reporting Lookup' }).click();
  await page.goto('https://ogt-gtm-web-eu-imp.7166.aws.thomsonreuters.com/gtm/aspx?href=%2FMaintenance%2FfmgProductLookup.aspx%3FProduct%3DGlobal%2BClassification%26Category%3DSearch%252fReporting');
  await page.locator('iframe[name="legacy-outlet"]').contentFrame().getByRole('link', { name: 'List for Not Requested Yet' }).click();
  const page1Promise = page.waitForEvent('popup');
  await page.locator('iframe[name="legacy-outlet"]').contentFrame().locator('tr', { hasText: requestNumber }).getByRole('link', { name: 'Approval' }).click();
  const page1 = await page1Promise;
  await page1.locator('iframe[name="legacy-outlet"]').contentFrame().getByRole('link', { name: 'Save' }).click();
  await page1.locator('iframe[name="legacy-outlet"]').contentFrame().getByText('Approval Validation Messages').click();
  await page1.locator('iframe[name="legacy-outlet"]').contentFrame().getByText('Item Gaihihantei').click();
  await page1.locator('iframe[name="legacy-outlet"]').contentFrame().locator('[id="ctl00_MainContent__GC Approval_ApprovalHeaderpnlDeclaration_TabContainer__GC Approval_ApprovalHeaderpnlDeclaration_TabContainer_TabPanel_ItemGaihihantei_drpKouban1Result"]').selectOption('1');
  await page1.locator('iframe[name="legacy-outlet"]').contentFrame().locator('[id="ctl00_MainContent__GC Approval_ApprovalHeaderpnlDeclaration_TabContainer__GC Approval_ApprovalHeaderpnlDeclaration_TabContainer_TabPanel_ItemGaihihantei_drpKouban1"]').selectOption('1(3)');
  await page1.locator('iframe[name="legacy-outlet"]').contentFrame().getByRole('link', { name: 'Save', exact: true }).click();
  await page1.locator('iframe[name="legacy-outlet"]').contentFrame().getByText('Item Gaihihantei').click();
  await page1.locator('iframe[name="legacy-outlet"]').contentFrame().locator('div').filter({ hasText: 'SaveSaveApproversTemplate' }).nth(3).click();
  await page1.locator('iframe[name="legacy-outlet"]').contentFrame().locator('[id="ctl00_MainContent__GC Approval_ApprovalHeaderpnlDeclaration_TabContainer__GC Approval_ApprovalHeaderpnlDeclaration_TabContainer_TabPanel_ItemGaihihantei_drpAppendix2Kouban1Result"]').selectOption('1');
  await page1.locator('iframe[name="legacy-outlet"]').contentFrame().locator('[id="ctl00_MainContent__GC Approval_ApprovalHeaderpnlDeclaration_TabContainer__GC Approval_ApprovalHeaderpnlDeclaration_TabContainer_TabPanel_ItemGaihihantei_drpAppendix2Kouban1"]').selectOption('35の4');
  await page1.locator('iframe[name="legacy-outlet"]').contentFrame().getByRole('link', { name: 'Save', exact: true }).click();
  await page1.locator('iframe[name="legacy-outlet"]').contentFrame().getByText('Item Gaihihantei').click();
  await page1.locator('iframe[name="legacy-outlet"]').contentFrame().locator('[id="ctl00_MainContent__GC Approval_ApprovalHeaderpnlDeclaration_TabContainer__GC Approval_ApprovalHeaderpnlDeclaration_TabContainer_TabPanel_ItemGaihihantei_drpApdx23Kouban"]').selectOption('1');
  await page1.locator('iframe[name="legacy-outlet"]').contentFrame().getByRole('link', { name: 'Save', exact: true }).click();

  // ログアウト処理
  await page.bringToFront();
  await page.getByRole('button', { name: /Tomohiro Oshikata/ }).click();
  await page.getByRole('button', { name: ' Sign out' }).click();
  await page.waitForURL(/Logon|default/i);
});