#!/usr/bin/env node
/**
 * 佐川 e飛伝III の現在 DOM が selectors.json と合うか確認する診断スクリプト。
 * CSV 取込・PDF までは行かず、ログイン〜入口ステップ後の URL とセレクタ件数を出力する。
 *
 * 使い方（worker ディレクトリで）:
 *   npm run probe:sagawa
 *
 * 必要: .env に SAGAWA_LOGIN_ID / SAGAWA_PASSWORD（任意で SAGAWA_LOGIN_URL, BROWSER_USER_AGENT）
 */

require('dotenv').config({ path: require('path').join(__dirname, '..', '.env') });

const fs = require('fs');
const path = require('path');
const { getSelectors } = require('../src/load-selectors');
const {
  SAGAWA_INFO_POPUP_URL_RE,
  launchBrowser,
  removeWalkMe,
  dismissMessageBox,
  dismissPortalNotice,
} = require('../src/carriers/sagawa-helpers');

const LOGIN_URL = process.env.SAGAWA_LOGIN_URL || 'https://www.e-service.sagawa-exp.co.jp/';
const LOGIN_ID = process.env.SAGAWA_LOGIN_ID;
const LOGIN_PASSWORD = process.env.SAGAWA_PASSWORD;
const TIMEOUT = 60_000;

async function countVisible(page, selector) {
  try {
    if (!page || page.isClosed()) return -1;
    return await page.locator(selector).count();
  } catch {
    return -1;
  }
}

async function safePageOp(page, label, fn) {
  if (!page || page.isClosed()) {
    console.warn(`${label}: ページが閉じているためスキップ`);
    return;
  }
  await fn();
}

async function main() {
  if (!LOGIN_ID || !LOGIN_PASSWORD) {
    console.error('SAGAWA_LOGIN_ID / SAGAWA_PASSWORD を .env に設定してください。');
    process.exit(1);
  }

  const sel = getSelectors('sagawa');
  const outDir = path.join(__dirname, '..', 'logs');
  fs.mkdirSync(outDir, { recursive: true });
  const shotPath = path.join(outDir, `sagawa-probe-${Date.now()}.png`);

  let context;
  let browser;
  let workPage;

  try {
    ({ context, browser } = await launchBrowser());
    const page = await context.newPage();
    page.setDefaultTimeout(TIMEOUT);

    console.log('\n=== 佐川セレクタ診断 ===\n');
    console.log('LOGIN_URL:', LOGIN_URL);

    await page.goto(LOGIN_URL, { waitUntil: 'load' });

    if (sel.businessTab) {
      const bizTab = page.locator(sel.businessTab).first();
      try {
        await bizTab.waitFor({ state: 'visible', timeout: 5000 });
        await bizTab.click();
        await page.locator(sel.loginId).first().waitFor({ state: 'visible', timeout: 5000 }).catch(() => {});
      } catch {
        console.log('法人タブ: 見つからずスキップ');
      }
    }

    await page.locator(sel.loginId).first().fill(LOGIN_ID);
    await page.locator(sel.password).first().fill(LOGIN_PASSWORD);
    await page.locator(sel.loginButton).first().click();
    await page.waitForLoadState('load');
    console.log('ログイン完了 url:', page.url());

    workPage = page;
    await page.waitForLoadState('networkidle').catch(() => {});
    await page.waitForTimeout(2000);
    await removeWalkMe(workPage);
    await dismissMessageBox(workPage);
    await dismissPortalNotice(workPage);

    const entrySteps = Array.isArray(sel.entrySteps) ? sel.entrySteps : [];
    for (let i = 0; i < entrySteps.length; i++) {
      const selector = entrySteps[i];
      await removeWalkMe(workPage);
      await dismissMessageBox(workPage);
      await dismissPortalNotice(workPage);

      const n = await countVisible(workPage, selector);
      console.log(`\n入口ステップ ${i + 1} セレクタ: ${selector}`);
      console.log(`  マッチ件数: ${n}`);

      const loc = workPage.locator(selector).first();
      try {
        await loc.waitFor({ state: 'visible', timeout: 15000 });
        console.log('  → visible OK');
      } catch {
        console.log('  → visible NG（本番と同じくスキップ）');
        continue;
      }

      const popupPromise = context.waitForEvent('page', { timeout: 15000 }).catch(() => null);
      await loc.click();
      let newPage = await popupPromise;
      if (newPage) {
        await newPage.waitForLoadState('domcontentloaded').catch(() => {});
        await newPage.waitForTimeout(3000);
        let openedUrl = '';
        try {
          openedUrl = newPage.url();
        } catch {
          /* closed */
        }
        if (i === 0 && openedUrl && SAGAWA_INFO_POPUP_URL_RE.test(openedUrl)) {
          console.warn('  → お知らせ(info)タブを検出。閉じて e飛伝(/info/除外)で再試行', { url: openedUrl });
          await newPage.close().catch(() => {});
          workPage = page;
          const retryEntrySelector = Array.isArray(sel.entrySteps) ? sel.entrySteps[0] : null;
          if (!retryEntrySelector) {
            console.warn('  → entrySteps[0] 未設定のため再試行できません');
            throw new Error('sagawa.entrySteps[0] が selectors.json にありません');
          }
          const alt = workPage.locator(retryEntrySelector).first();
          try {
            await alt.waitFor({ state: 'visible', timeout: 10000 });
            const p2 = context.waitForEvent('page', { timeout: 15000 }).catch(() => null);
            await alt.click();
            newPage = await p2;
            if (newPage) {
              await newPage.waitForLoadState('domcontentloaded').catch(() => {});
              await newPage.waitForTimeout(3000);
              workPage = newPage;
              workPage.setDefaultTimeout(TIMEOUT);
              console.log('  → 新規タブへ切替 url（再試行後）:', workPage.url());
            } else {
              await workPage.waitForLoadState('load').catch(() => {});
              await workPage.waitForTimeout(2000);
              console.log('  → 同一タブ url（再試行後）:', workPage.url());
            }
          } catch (e) {
            console.warn('  → 再試行失敗:', e.message);
          }
        } else {
          workPage = newPage;
          workPage.setDefaultTimeout(TIMEOUT);
          console.log('  → 新規タブへ切替 url:', openedUrl || workPage.url());
        }
        await removeWalkMe(workPage);
        await dismissMessageBox(workPage);
        await dismissPortalNotice(workPage);
      } else {
        await workPage.waitForLoadState('load');
        await workPage.waitForTimeout(3000);
        console.log('  → 同一タブ url:', workPage.url());
      }
    }

    if (!workPage || workPage.isClosed()) {
      console.error('\n診断: workPage が無効のため取込セレクタ確認を省略します');
      return;
    }

    await safePageOp(workPage, '後処理', async () => {
      await removeWalkMe(workPage);
      await dismissMessageBox(workPage);
      await workPage.waitForTimeout(1500);
    });

    if (workPage.isClosed()) {
      console.error('\n診断: ページが閉じたため以降を省略');
      return;
    }

    console.log('\n--- 取込画面付近セレクタ（入口後の workPage）---');
    console.log('url:', workPage.url());

    if (sel.templateDropdownTrigger) {
      const t = await countVisible(workPage, sel.templateDropdownTrigger);
      console.log(`templateDropdownTrigger (${sel.templateDropdownTrigger}): ${t} 件`);
    }
    if (sel.fileSelectButton) {
      const f = await countVisible(workPage, sel.fileSelectButton);
      console.log(`fileSelectButton (${sel.fileSelectButton}): ${f} 件`);
    }
    if (sel.directImportButton) {
      const d = await countVisible(workPage, sel.directImportButton);
      console.log(`directImportButton (${sel.directImportButton}): ${d} 件`);
    }

    if (!workPage.isClosed()) {
      await workPage.screenshot({ path: shotPath, fullPage: true });
      console.log('\nスクリーンショット:', shotPath);
    }
    console.log('\n=== 診断終了（件数が 0 のセレクタはサイト変更の疑い）===\n');
  } finally {
    if (context) await context.close().catch(() => {});
    if (browser) await browser.close().catch(() => {});
  }
}

main().catch((e) => {
  console.error(e);
  process.exit(1);
});
