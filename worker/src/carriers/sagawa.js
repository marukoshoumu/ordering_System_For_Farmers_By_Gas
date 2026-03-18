const { chromium } = require('playwright');
const path = require('path');
const fs = require('fs');
const os = require('os');
const { getSelectors } = require('../load-selectors');
const { MIN_PDF_SIZE } = require('../constants');

const LOGIN_URL = process.env.SAGAWA_LOGIN_URL || 'https://www.e-service.sagawa-exp.co.jp/';
const LOGIN_ID = process.env.SAGAWA_LOGIN_ID;
const LOGIN_PASSWORD = process.env.SAGAWA_PASSWORD;
const TIMEOUT = 60_000;

async function launchBrowser() {
  // 佐川WAF(Akamai)は従来のheadlessを検知してブロックする。
  // --headless=new（新headlessモード）は通常ブラウザと同等の挙動でWAFを回避でき、
  // かつウィンドウも表示されない。
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
    userAgent: process.env.BROWSER_USER_AGENT || 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36',
  });
  // 佐川のWAF/bot検知を回避
  await context.addInitScript(() => {
    Object.defineProperty(navigator, 'webdriver', { get: () => false });
  });
  return { context, browser };
}

/**
 * WalkMeオーバーレイを除去する。
 * e飛伝IIIではWalkMeが操作をブロックするため、ページ遷移後に毎回呼ぶ。
 */
async function removeWalkMe(page) {
  await page.evaluate(() => {
    document.querySelectorAll('[id^="walkme-"]').forEach(el => el.remove());
  }).catch(() => {});
}

/**
 * Element UIのメッセージボックス（荷物受渡書印刷案内等）を閉じる。
 * 未印刷の荷物受渡書がある場合、ログイン後にダイアログが表示され操作をブロックする。
 */
/**
 * @returns {boolean} ダイアログを閉じたかどうか
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
  } catch { /* ダイアログが無ければ何もしない */ }
  return false;
}

async function processSagawa(csvContent, shippingDate) {
  const sel = getSelectors('sagawa');
  const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'sagawa-'));
  const csvPath = path.join(tmpDir, `sagawa_${shippingDate}.csv`);
  fs.writeFileSync(csvPath, csvContent, 'utf-8');
  console.log('CSV一時ファイル作成', { csvPath });

  const downloadDir = path.join(tmpDir, 'downloads');
  fs.mkdirSync(downloadDir, { recursive: true });

  if (!LOGIN_ID || !LOGIN_PASSWORD) {
    throw new Error('SAGAWA_LOGIN_ID / SAGAWA_PASSWORD が未設定です');
  }

  let context, browser, workPage;
  try {
    ({ context, browser } = await launchBrowser());
    const page = await context.newPage();
    page.setDefaultTimeout(TIMEOUT);

    // Step 1: ログイン（OpenID Connect認証ページ経由）
    console.log('e飛伝III: ページ読み込み開始');
    await page.goto(LOGIN_URL, { waitUntil: 'load' });

    // 法人タブを選択
    if (sel.businessTab) {
      const bizTab = page.locator(sel.businessTab).first();
      try {
        await bizTab.waitFor({ state: 'visible', timeout: 5000 });
        await bizTab.click();
        // タブ切替後のフォーム表示を待つ（ログインフィールドが使えるまで）
        await page.locator(sel.loginId).first().waitFor({ state: 'visible', timeout: 5000 }).catch(() => {});
        console.log('e飛伝III: 法人タブ選択');
      } catch {
        console.log('e飛伝III: 法人タブが見つかりません（単一フォームの可能性）');
      }
    }

    console.log('e飛伝III: ログイン実行');
    await page.locator(sel.loginId).first().fill(LOGIN_ID);
    await page.locator(sel.password).first().fill(LOGIN_PASSWORD);
    await page.locator(sel.loginButton).first().click();
    await page.waitForLoadState('load');
    console.log('e飛伝III: ログイン完了');

    // Step 2: スマートクラブ → e飛伝III（新規タブで開く）
    workPage = page;
    // ログイン後のダッシュボード/メニュー表示を待つ（外部ページのため短い待機）
    await page.waitForLoadState('networkidle').catch(() => {});
    await page.waitForTimeout(2000);
    const entrySteps = Array.isArray(sel.entrySteps) ? sel.entrySteps : [];
    for (let i = 0; i < entrySteps.length; i++) {
      const selector = entrySteps[i];
      await removeWalkMe(workPage);
      const loc = workPage.locator(selector).first();
      try {
        await loc.waitFor({ state: 'visible', timeout: 15000 });
      } catch {
        console.log('e飛伝III: 入口ステップ', i + 1, 'が見つからず、スキップ');
        continue;
      }
      console.log('e飛伝III: 入口ステップ', i + 1);
      const popupPromise = context.waitForEvent('page', { timeout: 15000 }).catch(() => null);
      await loc.click();
      const newPage = await popupPromise;
      if (newPage) {
        await newPage.waitForLoadState('domcontentloaded').catch(() => {});
        // 新規タブのコンテンツ読み込みを待つ（e飛伝はSPAのため短い待機を維持）
        await newPage.waitForTimeout(3000);
        workPage = newPage;
        workPage.setDefaultTimeout(TIMEOUT);
        console.log('e飛伝III: 新しいタブに切り替え');
      } else {
        await workPage.waitForLoadState('load');
        await workPage.waitForTimeout(3000);
      }
    }

    // WalkMeオーバーレイ除去
    await removeWalkMe(workPage);

    // 荷物受渡書印刷ダイアログ等のElement UIメッセージボックスを閉じる
    // （未印刷の荷物受渡書がある場合にリダイレクトされるダイアログ）
    const dismissed = await dismissMessageBox(workPage);
    if (dismissed) {
      // ダイアログOKで荷物受渡書印刷画面にリダイレクトされる場合がある
      // 送り状データ取込ボタンが見えるまで待ち、見えなければ画面遷移済み
      await workPage.waitForTimeout(2000);
      await removeWalkMe(workPage);
    }

    // Step 3: テンプレート選択（Element UIカスタムドロップダウン）
    // e飛伝IIIではネイティブ<select>ではなくElement UIのカスタムドロップダウンを使用。
    // テンプレートドロップダウンをクリックしてリストを開き、目的の項目をクリックする。
    if (sel.templateDropdownTrigger) {
      const trigger = workPage.locator(sel.templateDropdownTrigger).first();
      try {
        await trigger.waitFor({ state: 'visible', timeout: 5000 });
        await trigger.click();
        await workPage.waitForTimeout(300);

        // ドロップダウンリスト内から「標準_飛脚宅配便_CSV_ヘッダ無」を選択
        const optionLoc = workPage.locator(sel.templateOption).first();
        await optionLoc.waitFor({ state: 'visible', timeout: 3000 });
        await optionLoc.click();
        await workPage.waitForTimeout(500);
        console.log('e飛伝III: テンプレート選択完了');
      } catch {
        console.warn('e飛伝III: テンプレート選択が見つかりません（デフォルトを使用）');
      }
    }

    // Step 4: CSVファイルアップロード
    // e飛伝IIIでは<input type="file">が非表示。fileChooserイベントを使用。
    console.log('e飛伝III: CSVアップロード');
    await removeWalkMe(workPage);

    const fileChooserPromise = workPage.waitForEvent('filechooser', { timeout: 10000 });
    await workPage.locator(sel.fileSelectButton).first().click();
    const fileChooser = await fileChooserPromise;
    await fileChooser.setFiles(csvPath);
    console.log('e飛伝III: ファイル選択完了');

    // Step 5: 直接取込
    // エラーダイアログがあればOK押す
    const okBtn = workPage.locator('button:has-text("OK")').first();
    if (await okBtn.isVisible().catch(() => false)) {
      await okBtn.click();
      await workPage.waitForTimeout(1000);
    }
    const directBtn = workPage.locator(sel.directImportButton).first();
    await directBtn.waitFor({ state: 'visible', timeout: 10000 });
    await directBtn.click();
    await workPage.waitForTimeout(5000);
    console.log('e飛伝III: 直接取込実行');

    // Step 6: 登録（取込確認ダイアログの「登録」ボタン）
    if (sel.registerButton) {
      const regBtn = workPage.locator(sel.registerButton).first();
      try {
        await regBtn.waitFor({ state: 'visible', timeout: 15000 });
        await regBtn.click();
        await workPage.waitForTimeout(5000);
        console.log('e飛伝III: 登録完了');
      } catch {
        console.warn('e飛伝III: 登録ボタンが見つかりません');
      }
    }

    // Step 7: 送り状印刷一覧で全選択
    await removeWalkMe(workPage);
    await workPage.waitForTimeout(1000);

    // e飛伝IIIでは「全て選択/取消」はカスタム要素（title属性で識別）
    const selectAllEl = workPage.locator(sel.selectAllToggle).first();
    try {
      await selectAllEl.waitFor({ state: 'visible', timeout: 10000 });
      await selectAllEl.click();
      await workPage.waitForTimeout(500);
      console.log('e飛伝III: 全選択');
    } catch {
      console.warn('e飛伝III: 全選択ボタンが見つかりません');
    }

    // Step 8: 印刷ボタン → 印刷条件設定ダイアログ → 印刷 → PDF取得
    const printBtn = workPage.locator(sel.printButton).first();
    await printBtn.waitFor({ state: 'visible', timeout: 10000 });
    await printBtn.click();
    console.log('e飛伝III: 印刷条件設定ダイアログ');

    const dialogPrintBtn = workPage.locator(sel.dialogPrintButton).first();
    await dialogPrintBtn.waitFor({ state: 'visible', timeout: 10000 });

    // directdownload のレスポンスを route で傍受（ワンタイムURL対策）
    let pdfBuffer = null;
    await context.route('**/directdownload**', async (route) => {
      const response = await route.fetch();
      const body = await response.body();
      const ct = response.headers()['content-type'] || '';
      console.log('e飛伝III: route傍受', { ct, size: body.length });
      if (ct.includes('pdf') || (!ct || ct === 'application/octet-stream') && body.length > MIN_PDF_SIZE) {
        pdfBuffer = body;
      }
      await route.fulfill({ response });
    });

    await dialogPrintBtn.click();
    console.log('e飛伝III: ダイアログ印刷ボタンクリック');

    // 「帳票出力中」ダイアログが出るので消えるまで待機
    try {
      await workPage.locator('text=帳票出力中').first().waitFor({ state: 'visible', timeout: 10000 });
      await workPage.locator('text=帳票出力中').first().waitFor({ state: 'hidden', timeout: TIMEOUT });
    } catch {
      // ダイアログが出ない場合もある
    }

    // PDFタブ読み込み完了を待つ（外部サーバ応答のため短い待機を維持）
    await workPage.waitForTimeout(5000);

    // route 解除
    await context.unroute('**/directdownload**');

    let pdfPath = path.join(downloadDir, `sagawa_${shippingDate}.pdf`);
    if (pdfBuffer && pdfBuffer.length > MIN_PDF_SIZE) {
      fs.writeFileSync(pdfPath, pdfBuffer);
      console.log('e飛伝III: PDFダウンロード完了', { pdfPath, size: pdfBuffer.length });
    } else {
      throw new Error('e飛伝III: PDF取得に失敗しました (size=' + (pdfBuffer?.length || 0) + ', 最小必要: ' + MIN_PDF_SIZE + ')');
    }

    return { pdfPath, tmpDir };
  } catch (error) {
    const screenshotPath = path.join(tmpDir, `sagawa_error_${Date.now()}.png`);
    try {
      const errorPage = workPage || context?.pages()[0];
      if (errorPage) await errorPage.screenshot({ path: screenshotPath, fullPage: true });
      console.error('e飛伝III: エラースクリーンショット保存', { screenshotPath });
    } catch (screenErr) {
      console.error('e飛伝III: エラースクリーンショット保存に失敗', { error: screenErr?.message || screenErr, screenshotPath });
    }
    console.error('e飛伝III: 処理エラー', { error: error.message });
    error.screenshotPath = screenshotPath;
    throw error;
  } finally {
    if (context) await context.close().catch(() => {});
    if (browser) await browser.close().catch(() => {});
    try { fs.unlinkSync(csvPath); } catch { /* ignore */ }
    console.log('e飛伝III: ブラウザ終了');
  }
}

module.exports = { processSagawa };
