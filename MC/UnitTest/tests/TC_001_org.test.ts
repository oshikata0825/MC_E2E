import { test, expect } from '@playwright/test';
import { BASE_URL } from './config';

test('TC_001', async ({ page }) => {
  // ECN Quick Lookup Screen Access
  // Verify that the ECN/Dual Use List Quick Lookup screen loads without errors
  await page.goto(BASE_URL + '/Logon.aspx?ReturnUrl=%2fgtm%2fhome');
  await page.getByRole('textbox', { name: 'Company' }).click();
  await page.getByRole('textbox', { name: 'Company' }).fill('Mitsubishi Corporation');
  await page.getByRole('textbox', { name: 'Company' }).press('Tab');
  await page.getByRole('textbox', { name: 'Username' }).click();
  await page.getByRole('textbox', { name: 'Username' }).fill('TOshikata');
  await page.getByRole('textbox', { name: 'Password' }).click();
  await page.getByRole('textbox', { name: 'Password' }).fill('5D9dtEi*Qaivft');
  await page.getByRole('link', { name: 'Login' }).click();
  await page.getByRole('button', { name: 'Tomohiro Oshikata ' }).click();
  await page.getByLabel('言語').selectOption('0: Object');
  await page.getByRole('button', { name: 'Content' }).click();
  await page.getByRole('menuitem', { name: 'ECN/Dual Use List (Quick' }).click();
  await page.goto(BASE_URL + '/gtm/aspx?href=%2FContent%2FfugECCNDetail.aspx');

  console.log('--- 最終解決フェーズ ---');

  // 1. "View As" パネルなどのオーバーレイ対策（念のため）
  const overlay = page.locator('text=Select group(s) to view as');
  if (await overlay.isVisible()) {
      console.log('⚠️ "View As" パネルを検出。必要に応じて閉じる処理を追加してください。');
  }

  // 2. 【修正点】名前(name属性)でピンポイントに指名します
  // これなら "legacy-bridge" と迷うことはありません
  console.log('★ メイン画面(legacy-outlet)の生成を待機中...');
  
  const iframeElement = page.locator('iframe[name="legacy-outlet"]');
  
  // 念のため存在確認（生成待ち）
  await iframeElement.waitFor({ state: 'attached', timeout: 30000 });
  console.log('★ Iframeを特定しました！');

  // 3. 中身の操作権を取得
  const contentFrame = iframeElement.contentFrame();
  if (!contentFrame) throw new Error('フレームの中身にアクセスできません');

  // 4. ステータスバー(pnlStatusBarItems)の描画を待機
  // パネル自体(#ctl00_MainContent_pnlStatusBarItems)を待ちます
  console.log('★ ステータスバーの描画を待機中...');
  
  const statusBar = contentFrame.locator('#ctl00_MainContent_pnlStatusBarItems');
  
  // Ajaxロード待ち（最長60秒）
  await expect(statusBar).toBeVisible({ timeout: 60000 });
  
  console.log('★ 大成功：ステータスバーが表示されました！テスト通過です。');

  // 5. (検証) 念のためSearchボタンの有無も確認
  const searchLink = statusBar.locator('a[href*="lnxbtnStatusBarSearch"]');
  if (await searchLink.count() > 0) {
      console.log('   -> Searchボタンリンクも確認済み');
  }

  // href属性の中に "lnxbtnStatusBarSearch" という文字列が含まれるリンクを探す
  // これなら # の有無や、見た目のテキスト(Search)に依存せずに特定できます
  // await page.locator('a[href*="lnxbtnStatusBarSearch"]').waitFor({ state: 'visible' });

  await page.getByRole('button', { name: 'Tomohiro Oshikata ' }).click();
  await page.getByRole('button', { name: ' Sign off' }).click();
});
