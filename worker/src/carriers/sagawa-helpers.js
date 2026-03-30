/**
 * 佐川 e飛伝III 自動化の共通ヘルパー（processSagawa と診断スクリプトで共有）
 */
const { chromium } = require('playwright');

const DEFAULT_BROWSER_USER_AGENT =
  'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36';

/** 送り状メニューの「お知らせ」専用ページ（e飛伝本体ではない） */
const SAGAWA_INFO_POPUP_URL_RE = /\/outer_wtx\/info\/|\/info\/info_\d{6,8}/;

/**
 * @param {{ userAgent?: string }} [opts]
 */
async function launchBrowser(opts = {}) {
  const userAgent =
    opts.userAgent !== undefined
      ? opts.userAgent
      : process.env.BROWSER_USER_AGENT || DEFAULT_BROWSER_USER_AGENT;

  const browser = await chromium.launch({
    headless: false,
    args: [
      '--no-sandbox',
      '--disable-setuid-sandbox',
      '--disable-dev-shm-usage',
      '--disable-blink-features=AutomationControlled',
      '--headless=new',
    ],
  });
  const context = await browser.newContext({
    acceptDownloads: true,
    locale: 'ja-JP',
    viewport: { width: 1280, height: 800 },
    userAgent,
  });
  await context.addInitScript(() => {
    Object.defineProperty(navigator, 'webdriver', { get: () => false });
  });
  return { context, browser };
}

async function removeWalkMe(page) {
  await page.evaluate(() => {
    document.querySelectorAll('[id^="walkme-"]').forEach((el) => el.remove());
  }).catch(() => {});
}

/**
 * Element UI のメッセージボックスを閉じる
 * @returns {Promise<boolean>}
 */
async function dismissMessageBox(page) {
  try {
    const msgBox = page.locator('.el-message-box__wrapper:visible .el-message-box__btns button').first();
    if (await msgBox.isVisible({ timeout: 2000 }).catch(() => false)) {
      await msgBox.click();
      console.log('e飛伝III: メッセージボックスを閉じました');
      await page.waitForTimeout(1000);
      return true;
    }
  } catch {
    /* ダイアログが無ければ何もしない */
  }
  return false;
}

/**
 * 既知の全画面お知らせだけを閉じる（閉じるボタンが無ければ何もしない）
 * @returns {Promise<boolean>}
 */
async function dismissPortalNotice(page) {
  try {
    const noticeMarker = page
      .getByText(/Microsoft Edge利用時|飛脚機密文書リサイクル/)
      .first();
    if (!(await noticeMarker.isVisible({ timeout: 2500 }).catch(() => false))) {
      return false;
    }
    const closeBtn = page.getByRole('button', { name: /閉じる/ }).first();
    if (!(await closeBtn.isVisible({ timeout: 2000 }).catch(() => false))) {
      return false;
    }
    await closeBtn.click();
    await page.waitForTimeout(800);
    console.log('e飛伝III: スマートクラブのお知らせを閉じました');
    return true;
  } catch {
    /* ignore */
  }
  return false;
}

module.exports = {
  SAGAWA_INFO_POPUP_URL_RE,
  launchBrowser,
  removeWalkMe,
  dismissMessageBox,
  dismissPortalNotice,
};
