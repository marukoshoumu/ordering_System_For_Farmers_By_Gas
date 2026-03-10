const { chromium } = require('playwright');
const path = require('path');
const fs = require('fs');
const os = require('os');
const { getSelectors } = require('../load-selectors');

const LOGIN_URL = process.env.SAGAWA_LOGIN_URL || 'https://www.e-hidenweb2.sagawa-exp.co.jp/';
const LOGIN_ID = process.env.SAGAWA_LOGIN_ID;
const LOGIN_PASSWORD = process.env.SAGAWA_PASSWORD;
const HEADLESS = process.env.HEADLESS !== 'false';
const TIMEOUT = 60_000;

async function launchBrowser() {
  const browser = await chromium.launch({
    headless: HEADLESS,
    args: ['--no-sandbox', '--disable-setuid-sandbox', '--disable-dev-shm-usage'],
  });
  const context = await browser.newContext({
    acceptDownloads: true,
    locale: 'ja-JP',
    viewport: { width: 1280, height: 800 },
  });
  return { context, browser };
}

async function processSagawa(csvContent, shippingDate) {
  const sel = getSelectors('sagawa');
  const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'sagawa-'));
  const csvPath = path.join(tmpDir, `sagawa_${shippingDate}.csv`);
  fs.writeFileSync(csvPath, csvContent, 'utf-8');
  console.log('CSV一時ファイル作成', { csvPath });

  const downloadDir = path.join(tmpDir, 'downloads');
  fs.mkdirSync(downloadDir, { recursive: true });

  let context, browser;
  try {
    ({ context, browser } = await launchBrowser());
    const page = await context.newPage();
    page.setDefaultTimeout(TIMEOUT);

    // Step 1: ログイン
    console.log('e飛伝Web: ページ読み込み開始');
    await page.goto(LOGIN_URL, { waitUntil: 'networkidle' });

    if (!LOGIN_ID || !LOGIN_PASSWORD) {
      throw new Error('SAGAWA_LOGIN_ID / SAGAWA_PASSWORD が未設定です');
    }
    console.log('e飛伝Web: ログイン実行');
    await page.fill(sel.loginId, LOGIN_ID);
    await page.fill(sel.password, LOGIN_PASSWORD);
    await page.click(sel.loginButton);
    await page.waitForLoadState('networkidle');
    console.log('e飛伝Web: ログイン完了');

    // Step 2: CSV取込画面へ遷移
    console.log('e飛伝Web: CSV取込画面へ遷移');
    await page.click(sel.csvImportMenu);
    await page.waitForLoadState('networkidle');

    // Step 3: CSVアップロード
    console.log('e飛伝Web: CSVアップロード');
    await page.setInputFiles(sel.fileInput, csvPath);
    await page.click(sel.uploadButton);
    await page.waitForLoadState('networkidle');
    await page.waitForTimeout(3000);

    // エラーチェック
    const errorEl = await page.$('.error, .alert-danger, [class*="error"], .errMsg');
    if (errorEl) {
      const errorText = await errorEl.textContent();
      if (errorText?.trim()) throw new Error('e飛伝WebCSV取込エラー: ' + errorText.trim());
    }

    // Step 4: 伝票PDF出力
    console.log('e飛伝Web: 伝票PDF出力開始');
    await page.click(sel.printButton);
    await page.waitForTimeout(2000);

    // Step 5: PDFダウンロード
    console.log('e飛伝Web: PDFダウンロード待機');
    const downloadPromise = page.waitForEvent('download', { timeout: TIMEOUT });
    await page.click(sel.downloadButton);

    let pdfPath;
    try {
      const download = await downloadPromise;
      pdfPath = path.join(downloadDir, download.suggestedFilename() || `sagawa_${shippingDate}.pdf`);
      await download.saveAs(pdfPath);
      console.log('e飛伝Web: PDFダウンロード完了', { pdfPath });
    } catch (dlErr) {
      console.warn('ダウンロードイベント未検出。ページPDF化を試行', { error: dlErr.message });
      pdfPath = path.join(downloadDir, `sagawa_${shippingDate}.pdf`);
      await page.pdf({ path: pdfPath, format: 'A4' });
    }

    return pdfPath;
  } catch (error) {
    try {
      const screenshotPath = path.join(tmpDir, `sagawa_error_${Date.now()}.png`);
      const page = context?.pages()[0];
      if (page) await page.screenshot({ path: screenshotPath, fullPage: true });
      console.error('e飛伝Web: エラースクリーンショット保存', { screenshotPath });
    } catch { /* ignore */ }
    console.error('e飛伝Web: 処理エラー', { error: error.message });
    throw error;
  } finally {
    if (context) await context.close().catch(() => {});
    if (browser) await browser.close().catch(() => {});
    try { fs.unlinkSync(csvPath); } catch { /* ignore */ }
    console.log('e飛伝Web: ブラウザ終了');
  }
}

module.exports = { processSagawa };
