const { chromium } = require('playwright');
const path = require('path');
const fs = require('fs');
const os = require('os');
const { getSelectors } = require('../load-selectors');

const LOGIN_URL = process.env.YAMATO_LOGIN_URL || 'https://bmypage.kuronekoyamato.co.jp/bmypage/servlet/jp.co.kuronekoyamato.wur.hmp.servlet.user.HMPLGI0010JspServlet';
const LOGIN_ID = process.env.YAMATO_LOGIN_ID;
const LOGIN_PASSWORD = process.env.YAMATO_PASSWORD;
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

async function processYamato(csvContent, shippingDate) {
  const sel = getSelectors('yamato');
  const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'yamato-'));
  const csvPath = path.join(tmpDir, `yamato_${shippingDate}.csv`);
  fs.writeFileSync(csvPath, csvContent, 'utf-8');
  console.log('CSV一時ファイル作成', { csvPath });

  const downloadDir = path.join(tmpDir, 'downloads');
  fs.mkdirSync(downloadDir, { recursive: true });

  if (!LOGIN_ID || !LOGIN_PASSWORD) {
    throw new Error('YAMATO_LOGIN_ID / YAMATO_PASSWORD が未設定です');
  }

  let context, browser, workPage;
  try {
    ({ context, browser } = await launchBrowser());
    const page = await context.newPage();
    page.setDefaultTimeout(TIMEOUT);

    // === Step 1: ログイン ===
    console.log('B2クラウド: ページ読み込み開始');
    await page.goto(LOGIN_URL, { waitUntil: 'load' });

    console.log('B2クラウド: ログイン実行');
    await page.locator(sel.loginId).first().click();
    await page.locator(sel.loginId).first().fill(LOGIN_ID);
    await page.locator(sel.password).first().click();
    await page.locator(sel.password).first().fill(LOGIN_PASSWORD);

    // a.login は javascript:void(0) + onclick でJS経由のフォーム送信
    const loginPageUrl = page.url();
    const loginBtn = page.locator(sel.loginButton).first();
    await loginBtn.waitFor({ state: 'visible', timeout: 5000 });
    await loginBtn.click();

    // URL変更を待つ。click() でJSイベントが発火しない場合はevaluateでフォールバック
    try {
      await page.waitForURL((url) => url.href !== loginPageUrl, { timeout: 15000 });
    } catch {
      console.log('B2クラウド: click() で遷移せず、evaluate でログイン試行');
      await page.evaluate(() => document.querySelector('a.login')?.click());
      await page.waitForURL((url) => url.href !== loginPageUrl, { timeout: 15000 });
    }
    await page.waitForLoadState('load');
    await page.waitForTimeout(2000);
    console.log('B2クラウド: ログイン完了', { url: page.url() });

    // === Step 2: B2クラウド画面への遷移 ===
    // ビジネスメンバーズ → 「送り状発行システムB2クラウド」 → 新規ウインドウで開く
    workPage = page;
    for (const selector of (sel.entrySteps || [])) {
      const loc = workPage.locator(selector).first();
      try {
        await loc.waitFor({ state: 'visible', timeout: 20000 });
      } catch {
        console.log('B2クラウド: 入口ステップが見つからず、スキップ', { selector: selector.slice(0, 50) });
        continue;
      }
      console.log('B2クラウド: 入口ステップ', selector.slice(0, 50));
      // B2クラウドは新規ウインドウで開く
      const popupPromise = context.waitForEvent('page', { timeout: 15000 }).catch(() => null);
      await loc.click();
      const newPage = await popupPromise;
      if (newPage) {
        await newPage.waitForLoadState('load').catch(() => {});
        workPage = newPage;
        workPage.setDefaultTimeout(TIMEOUT);
        console.log('B2クラウド: 新規ウインドウに切り替え', { url: workPage.url() });
      } else {
        await workPage.waitForLoadState('load');
        console.log('B2クラウド: 同一ウインドウで遷移', { url: workPage.url() });
      }
    }

    // === Step 3: 外部データから発行 ===
    // #ex_data_import 内のリンクをクリック（FAQリンクと混同しないよう ID 指定）
    const extLink = workPage.locator(sel.externalDataLink).first();
    await extLink.waitFor({ state: 'visible', timeout: 15000 });
    console.log('B2クラウド: 外部データから発行へ');
    await extLink.click();
    await workPage.waitForLoadState('load');
    await workPage.waitForTimeout(1000);

    // === Step 4: 取込み開始行を 1 に変更 ===
    // ページ初期化JSが torikomi_strat_row を 2 にリセットするため、
    // 初期化完了（値が安定）を待ってから変更する
    if (sel.importStartRowInput) {
      const rowInput = workPage.locator(sel.importStartRowInput).first();
      try {
        await rowInput.waitFor({ state: 'visible', timeout: 5000 });
        // ページJSの初期化完了を待つ（値が2にセットされるまで）
        await workPage.waitForFunction(() => {
          const el = document.querySelector('input#torikomi_strat_row');
          return el && el.value !== '';
        }, { timeout: 5000 });
        await workPage.waitForTimeout(1000);
        await rowInput.click({ clickCount: 3 });
        await rowInput.fill('1');
        await rowInput.press('Tab');
        console.log('B2クラウド: 取込み開始行を1に変更');
      } catch {
        console.warn('B2クラウド: 取込み開始行フィールドが見つかりません');
      }
    }

    // === Step 5: CSVファイルアップロード ===
    // B2クラウド: a#file_button click → input#filename.click() → fileChooser
    //   → input#filename change → handleFiles() → ボタン有効化
    console.log('B2クラウド: CSVアップロード');
    const fileChooserPromise = workPage.waitForEvent('filechooser', { timeout: 10000 });
    // a#file_button の jQuery ハンドラ: $('#filename').val('').click()
    // Playwright native click では jQuery .on() が反応しないことがあるため evaluate で発火
    // jQuery が未読み込みの場合はネイティブ DOM でクリック
    await workPage.evaluate(() => {
      var btn = document.querySelector('#file_button');
      if (btn) {
        if (typeof window.jQuery !== 'undefined') {
          window.jQuery('#file_button').click();
        } else {
          btn.click();
        }
      }
    });
    const fileChooser = await fileChooserPromise;
    await fileChooser.setFiles(csvPath);
    console.log('B2クラウド: ファイル選択完了');

    // handleFiles() の FileReader 処理完了を待つ
    await workPage.waitForTimeout(3000);

    // 取込み開始ボタンが有効になるまで待つ
    const uploadBtn = workPage.locator(sel.uploadButton).first();
    try {
      await workPage.waitForFunction(() => {
        const btn = document.querySelector('a#import_start');
        return btn && !btn.classList.contains('disable');
      }, { timeout: 15000 });
      console.log('B2クラウド: 取込み開始ボタンが有効化');
    } catch {
      console.warn('B2クラウド: 取込み開始ボタンが無効のまま — CSV形式エラーの可能性');
    }

    // 取込み開始
    console.log('B2クラウド: 取込み開始ボタンクリック');
    await uploadBtn.click();
    await workPage.waitForLoadState('load');
    await workPage.waitForTimeout(3000);

    // エラーチェック
    const errorSelector = sel.errorSelectors || '.error, .alert-danger, [class*="error"]';
    const errorEl = await workPage.locator(errorSelector).first().elementHandle().catch(() => null);
    if (errorEl) {
      const errorText = await errorEl.textContent();
      if (errorText?.trim()) throw new Error('B2クラウドCSV取込エラー: ' + errorText.trim());
    }

    // === Step 6: 全選択 ===
    await workPage.waitForTimeout(1000);
    const checkAll = workPage.locator(sel.selectAllCheckbox).first();
    try {
      await checkAll.waitFor({ state: 'visible', timeout: 10000 });
      await checkAll.check();
      console.log('B2クラウド: 全選択');
    } catch {
      console.warn('B2クラウド: 全選択チェックボックスが見つかりません');
    }
    await workPage.waitForTimeout(500);

    // === Step 7: 印刷内容の確認へ ===
    const confirmBtn = workPage.locator(sel.confirmPrintButton).first();
    await confirmBtn.waitFor({ state: 'visible', timeout: 10000 });
    console.log('B2クラウド: 印刷内容の確認へ');
    await confirmBtn.click();
    await workPage.waitForLoadState('load');

    // === Step 8: 発行開始 → PDF取得 ===
    // B2クラウドは「発行開始」後、PDF表示用iFrameが表示される
    // iFrame内のPDFのURLを取得して直接ダウンロードする
    const issueBtn = workPage.locator(sel.startIssueButton).first();
    await issueBtn.waitFor({ state: 'visible', timeout: 15000 });
    // jQuery の読み込み完了を待つ
    await workPage.waitForFunction(() => typeof $ !== 'undefined' || typeof jQuery !== 'undefined', { timeout: 15000 });
    await workPage.waitForTimeout(1000);

    let pdfPath = path.join(downloadDir, `yamato_${shippingDate}.pdf`);

    // page.route() でPDFリクエストを傍受し、レスポンスを中継しつつバッファを保存
    let pdfBuffer = null;
    await workPage.route('**/B2_OKURIJYO*fileonly*', async (route) => {
      const response = await route.fetch();
      pdfBuffer = await response.body();
      console.log('B2クラウド: route傍受', { size: pdfBuffer.length, ct: response.headers()['content-type'] });
      await route.fulfill({ response });
    });

    console.log('B2クラウド: 発行開始');
    await workPage.evaluate(() => (typeof $ !== 'undefined' ? $ : jQuery)('#start_print').click());

    // fancybox iFrame が表示されるまで待つ（= PDF が読み込まれた）
    await workPage.waitForSelector('iframe.fancybox-iframe', { state: 'visible', timeout: 60000 });
    console.log('B2クラウド: PDF iFrame 表示');
    await workPage.waitForTimeout(2000);

    // route 解除
    await workPage.unroute('**/B2_OKURIJYO*fileonly*');

    if (pdfBuffer && pdfBuffer.length > 1000) {
      fs.writeFileSync(pdfPath, pdfBuffer);
      console.log('B2クラウド: PDFダウンロード完了', { pdfPath, size: pdfBuffer.length });
    } else {
      throw new Error('B2クラウド: PDF取得に失敗しました (size=' + (pdfBuffer?.length || 0) + ')');
    }

    return { pdfPath, tmpDir };
  } catch (error) {
    const screenshotPath = path.join(tmpDir, `yamato_error_${Date.now()}.png`);
    try {
      const errorPage = workPage || context?.pages()[0];
      if (errorPage) await errorPage.screenshot({ path: screenshotPath, fullPage: true });
      console.error('B2クラウド: エラースクリーンショット保存', { screenshotPath });
    } catch (screenErr) {
      console.error('B2クラウド: スクリーンショット保存失敗', { error: screenErr?.message });
    }
    console.error('B2クラウド: 処理エラー', { error: error.message });
    throw error;
  } finally {
    if (context) await context.close().catch(() => {});
    if (browser) await browser.close().catch(() => {});
    try { fs.unlinkSync(csvPath); } catch { /* ignore */ }
    console.log('B2クラウド: ブラウザ終了');
  }
}

module.exports = { processYamato };
