/**
 * サイドバーHTMLを生成するヘルパー関数
 * createTemplateFromFile を使用してGASスクリプトレットを正しく評価する。
 * createHtmlOutputFromFile ではスクリプトレットが処理されないため、この関数を使う。
 *
 * @param {string} deployURL - アプリケーションのデプロイURL
 * @param {string} sessionId - セッションID
 * @param {string} userRole - ユーザー権限（admin/viewer）
 * @returns {string} 評価済みのサイドバーHTML文字列
 */
function includeSidebar(deployURL, sessionId, userRole) {
  const template = HtmlService.createTemplateFromFile('sidebar');
  template.deployURL = deployURL || '';
  template.sessionId = sessionId || '';
  template.userRole = userRole || 'viewer';
  return template.evaluate().getContent();
}

/**
 * GETリクエスト処理（アプリケーションエントリーポイント）
 *
 * Google Apps ScriptのWebアプリケーションに対するGETリクエストを処理します。
 * URLパラメータに応じてテストページまたはログイン画面を表示します。
 *
 * 処理フロー:
 * 1. URLパラメータをチェック
 * 2. ?test=components の場合: test-components.html を表示（開発用）
 * 3. その他の場合: index.html（ログイン画面）を表示
 * 4. URLパラメータ（tempOrderId, redirectTo）を保持してログイン後に引き継ぎ
 *
 * URLパラメータ:
 * - test=components: コンポーネントテストページを表示
 * - tempOrderId: 仮受注ID（AI取込一覧からの遷移時）
 * - aiImportList: AI取込一覧への直接リンク用フラグ
 * - shipping: 受注画面への直接リンク用フラグ
 *
 * @param {Object} e - GETリクエストイベントオブジェクト
 * @param {Object} e.parameter - URLパラメータのkey-valueオブジェクト
 * @returns {HtmlOutput} HTMLページ（test-components.html または index.html）
 *
 * 使用例:
 * - https://script.google.com/.../exec → ログイン画面
 * - https://script.google.com/.../exec?test=components → テストページ
 * - https://script.google.com/.../exec?tempOrderId=abc123 → ログイン後に仮受注abc123を開く
 *
 * @see doPost() - POSTリクエスト処理（ログイン、画面遷移）
 * @see index.html - ログイン画面テンプレート
 * @see test-components.html - コンポーネントテストページ
 *
 * 呼び出し元: ブラウザからのGETリクエスト
 */
function doGet(e) {
  // テスト用：?test=components でコンポーネントテストページを表示
  if (e.parameter.test === 'components') {
    try {
      const template = HtmlService.createTemplateFromFile('test-components');
      const htmlOutput = template.evaluate();
      htmlOutput.addMetaTag('viewport', 'width=device-width, initial-scale=1');
      htmlOutput.setTitle('コンポーネントテスト');
      return htmlOutput;
    } catch (error) {
      // エラーが発生した場合、エラー情報を表示
      return HtmlService.createHtmlOutput(
        '<html><body style="padding:20px;font-family:Arial;"><h1>エラー発生</h1><p>' +
        error.message + '</p><pre>' + error.stack + '</pre></body></html>'
      );
    }
  }

  const template = HtmlService.createTemplateFromFile('index');
  template.deployURL = ScriptApp.getService().getUrl();

  // URLパラメータを保持してログイン画面に渡す
  template.tempOrderId = e.parameter.tempOrderId || '';
  template.redirectTo = e.parameter.aiImportList ? 'aiImportList'
    : e.parameter.shipping ? 'shipping'
      : '';

  const htmlOutput = template.evaluate();
  htmlOutput.addMetaTag('viewport', 'width=device-width, initial-scale=1');
  htmlOutput.setTitle('ログイン画面');
  return htmlOutput;
}

/**
 * ログインエラーメッセージHTML生成
 *
 * ログイン失敗時に表示するエラーメッセージのHTMLを生成します。
 * Bootstrap の text-danger クラスで赤色表示されます。
 *
 * @param {string} alert - エラーメッセージ（デフォルト: 空文字）
 * @returns {string} エラーメッセージのHTML文字列（<p class="text-danger">...</p>）
 *
 * 使用例:
 * const html = getLoginHTML('IDまたはパスワードが間違っています。');
 * // 返却: '<p class="text-danger">IDまたはパスワードが間違っています。</p>'
 *
 * @see doPost() - ログイン処理でこの関数を呼び出し
 * @see login.html - ログイン画面テンプレート
 */
function getLoginHTML(alert = '') {
  let html = ``;
  html += `<p class="text-danger">${alert}</p>`
  return html;
}

/** セッション有効期限（秒）6時間 */
const SESSION_TTL_SEC = 6 * 60 * 60;

/**
 * サーバー側セッションからユーザー情報を取得する。認証・権限はクライアント送信値に依存せずここで検証する。
 * @param {string} sessionId - ログイン時に発行されたセッションID
 * @returns {{ userRole: string, loginId: string } | null} セッションがあれば { userRole, loginId }、なければ null
 */
function getSession(sessionId) {
  if (!sessionId || typeof sessionId !== 'string') return null;
  try {
    const cache = CacheService.getScriptCache();
    const raw = cache.get('session_' + sessionId);
    return raw ? JSON.parse(raw) : null;
  } catch (err) {
    return null;
  }
}

/**
 * ログイン成功時にセッションを保存する。
 * @param {string} sessionId - Utilities.getUuid() 等で生成した一意のセッションID
 * @param {string} userRole - ユーザー権限（admin/viewer）。未設定・空文字の場合は最小権限（viewer）として保存する。
 * @param {string} loginId - ログインID
 */
function setSession(sessionId, userRole, loginId) {
  if (!sessionId) return;
  const cache = CacheService.getScriptCache();
  const role = (userRole && String(userRole).trim()) ? String(userRole).trim() : 'viewer';
  cache.put('session_' + sessionId, JSON.stringify({ userRole: role === 'admin' ? 'admin' : 'viewer', loginId: loginId || '' }), SESSION_TTL_SEC);
}

/**
 * ログアウト時にセッションをキャッシュから明示的に削除する。setSession と同じキー（session_ + sessionId）を使用。
 * @param {string} sessionId - 削除するセッションID
 * @returns {boolean} sessionId が有効な場合に削除して true、無効な場合は false
 */
function deleteSession(sessionId) {
  if (!sessionId || typeof sessionId !== 'string') return false;
  try {
    const cache = CacheService.getScriptCache();
    cache.remove('session_' + sessionId);
    return true;
  } catch (err) {
    return false;
  }
}

/**
 * POSTリクエストからセッションを解決し、サーバー側で確定した userRole と sessionId を返す。
 * セッションが無い場合は最小権限（viewer）とする。
 * @param {Object} e - POST イベント（e.parameter を使用）
 * @returns {{ userRole: string, sessionId: string }}
 */
function resolveSessionFromPost(e) {
  const sessionId = (e.parameter && e.parameter.sessionId) ? String(e.parameter.sessionId).trim() : '';
  const session = getSession(sessionId);
  const userRole = (session && session.userRole) ? session.userRole : 'viewer';
  return { userRole: userRole, sessionId: sessionId };
}

/**
 * POSTリクエスト処理（ログイン認証・画面遷移ルーティング）
 *
 * Google Apps ScriptのWebアプリケーションに対するPOSTリクエストを処理します。
 * ログイン認証、画面遷移、権限チェックを一元管理するルーティング関数です。
 *
 * 主要な処理分岐:
 * - e.parameter.logout: ログアウト処理 → セッション削除後にログイン画面を表示
 * - e.parameter.login: ログイン認証処理 → 成功時にホーム画面またはリダイレクト先へ遷移
 * - e.parameter.main / mainTop: ホーム画面表示
 * - e.parameter.phoneOrder: 電話受注モード（viewer権限不可）
 * - e.parameter.shipping: 受注画面（編集モード・AI取込対応）
 * - e.parameter.shippingComfirm: 受注確認画面
 * - e.parameter.shippingModify: 受注修正画面（確認画面から戻る）
 * - e.parameter.shippingSubmit: 受注登録実行 → 完了画面
 * - e.parameter.createShippingSlips: 送り状作成画面
 * - e.parameter.createBill: 請求書作成画面
 * - e.parameter.csvImport: CSV取込画面
 * - e.parameter.quotation: 見積書作成画面
 * - e.parameter.salesDashboard: 売上ダッシュボード（admin権限のみ）
 * - e.parameter.orderList: 製造数一覧（ヒートマップ）
 * - e.parameter.orderListPage: 受注一覧
 * - e.parameter.aiImportList: AI取込一覧（viewer権限不可）
 *
 * 権限制御:
 * - admin: 全機能アクセス可
 * - viewer: 参照系のみ（受注登録・AI取込一覧・電話受注モード不可）
 *
 * ログイン処理フロー:
 * 1. passシートから認証情報取得
 * 2. ID・パスワード照合
 * 3. 認証成功 → 権限情報取得、リダイレクト先チェック
 * 4. viewer権限の場合、AI取込一覧へのアクセスをブロック
 * 5. tempOrderIdがあれば仮受注データを受注画面に展開
 * 6. redirectToパラメータに応じて遷移先を決定
 *
 * @param {Object} e - POSTリクエストイベントオブジェクト
 * @param {Object} e.parameter - フォームパラメータのkey-valueオブジェクト
 * @param {string} e.parameter.login - ログインボタン押下フラグ
 * @param {string} e.parameter.loginId - ログインID
 * @param {string} e.parameter.loginPass - ログインパスワード
 * @param {string} e.parameter.sessionId - セッションID（ログイン後発行、権限はサーバー側セッションで検証）
 * @param {string} e.parameter.tempOrderId - 仮受注ID（AI取込一覧から遷移時）
 * @param {string} e.parameter.redirectTo - リダイレクト先（aiImportList/shipping）
 * @param {string} e.parameter.editOrderId - 編集対象の受注ID
 * @param {string} e.parameter.editMode - 編集モードフラグ（true/false）
 * @returns {HtmlOutput} 各画面のHTMLページ
 *
 * @see doGet() - GETリクエスト処理
 * @see redirectToHome() - 権限不足時のホーム画面リダイレクト
 * @see getshippingHTML() - 受注画面HTML生成
 * @see getshippingHTMLForTempOrder() - 仮受注データから受注画面HTML生成
 * @see createOrder() - 受注登録処理
 * @see home.html - ホーム画面テンプレート
 * @see shipping.html - 受注画面テンプレート
 * @see aiImportList.html - AI取込一覧テンプレート
 *
 * 呼び出し元: 各HTMLページのフォーム送信
 */
function doPost(e) {
  Logger.log(e);
  if (e.parameter.logout) {
    const sessionId = (e.parameter.sessionId) ? String(e.parameter.sessionId).trim() : '';
    deleteSession(sessionId);
    const template = HtmlService.createTemplateFromFile('login');
    template.deployURL = ScriptApp.getService().getUrl();
    template.loginHTML = getLoginHTML('');
    const htmlOutput = template.evaluate();
    htmlOutput.addMetaTag('viewport', 'width=device-width, initial-scale=1');
    htmlOutput.setTitle('ログイン画面');
    return htmlOutput;
  }
  if (e.parameter.login) {
    const items = getAllRecords('pass');
    // 入力されたIDとパスワードに一致するユーザーを検索
    const user = items.find(item =>
      item['id'] === e.parameter.loginId && item['pass'] === e.parameter.loginPass
    );

    if (!user) {
      // ユーザーが見つからない場合はログインエラー
      const template = HtmlService.createTemplateFromFile('login');
      template.deployURL = ScriptApp.getService().getUrl();
      const alert = 'IDまたはパスワードが間違っています。';
      template.loginHTML = getLoginHTML(alert);
      const htmlOutput = template.evaluate();
      htmlOutput.addMetaTag('viewport', 'width=device-width, initial-scale=1');
      htmlOutput.setTitle('ログイン画面');
      return htmlOutput;
    }
    // ログイン成功 - 権限情報を取得しサーバー側セッションを発行（クライアントの userRole 送信に依存しない）。権限が未設定・空の場合は最小権限（viewer）とする。
    const rawRole = (user['権限'] != null && typeof user['権限'] === 'string') ? String(user['権限']).trim() : '';
    const userRole = rawRole === 'admin' ? 'admin' : 'viewer';
    const sessionId = Utilities.getUuid();
    setSession(sessionId, userRole, e.parameter.loginId);
    const tempOrderId = e.parameter.tempOrderId || '';
    const redirectTo = e.parameter.redirectTo || '';

    // viewerの場合、AI取込一覧へのアクセスを制限
    if (userRole === 'viewer' && (redirectTo === 'aiImportList' || tempOrderId)) {
      return redirectToHome(userRole, sessionId);
    }

    // tempOrderIdがあれば受注画面へ直接遷移
    if (tempOrderId) {
      const template = HtmlService.createTemplateFromFile('shipping');
      template.deployURL = ScriptApp.getService().getUrl();
      template.sessionId = sessionId;
      template.userRole = userRole;
      template.isEditMode = false;
      template.isCopyMode = false;
      template.editOrderId = '';
      template.tempOrderId = tempOrderId;
      template.shippingHTML = getshippingHTMLForTempOrder(tempOrderId);
      template.autoOpenAI = false;
      template.aiAnalysisResult = '';
      // 検索条件引継用
      template.prevPeriod = e.parameter.prevPeriod || '';
      template.prevDateFrom = e.parameter.prevDateFrom || '';
      template.prevDateTo = e.parameter.prevDateTo || '';
      template.prevDestination = e.parameter.prevDestination || '';
      template.prevCustomer = e.parameter.prevCustomer || '';
      template.prevStatus = e.parameter.prevStatus || '';
      template.prevIncludeOverdue = e.parameter.prevIncludeOverdue || '';
      const htmlOutput = template.evaluate();
      htmlOutput.addMetaTag('viewport', 'width=device-width, initial-scale=1');
      htmlOutput.setTitle('受注画面（AI入力）');
      return htmlOutput;
    }

    // リダイレクト先パラメータに応じて遷移
    if (redirectTo === 'aiImportList') {
      const template = HtmlService.createTemplateFromFile('aiImportList');
      template.deployURL = ScriptApp.getService().getUrl();
      template.sessionId = sessionId;
      template.userRole = userRole;
      const htmlOutput = template.evaluate();
      htmlOutput.addMetaTag('viewport', 'width=device-width, initial-scale=1');
      htmlOutput.setTitle('AI取込一覧');
      return htmlOutput;
    }

    if (redirectTo === 'shipping') {
      const template = HtmlService.createTemplateFromFile('shipping');
      template.deployURL = ScriptApp.getService().getUrl();
      template.sessionId = sessionId;
      template.userRole = userRole;
      template.isEditMode = false;
      template.isCopyMode = false;
      template.editOrderId = '';
      template.tempOrderId = '';
      template.autoOpenAI = false;
      template.aiAnalysisResult = '';
      template.shippingHTML = getshippingHTML(e);
      // 検索条件引継用
      template.prevPeriod = e.parameter.prevPeriod || '';
      template.prevDateFrom = e.parameter.prevDateFrom || '';
      template.prevDateTo = e.parameter.prevDateTo || '';
      template.prevDestination = e.parameter.prevDestination || '';
      template.prevCustomer = e.parameter.prevCustomer || '';
      template.prevStatus = e.parameter.prevStatus || '';
      template.prevIncludeOverdue = e.parameter.prevIncludeOverdue || '';
      const htmlOutput = template.evaluate();
      htmlOutput.addMetaTag('viewport', 'width=device-width, initial-scale=1');
      htmlOutput.setTitle('受注画面');
      return htmlOutput;
    }

    const template = HtmlService.createTemplateFromFile('home');
    template.deployURL = ScriptApp.getService().getUrl();
    template.sessionId = sessionId;
    template.userRole = userRole;
    const htmlOutput = template.evaluate();
    htmlOutput.addMetaTag('viewport', 'width=device-width, initial-scale=1');
    htmlOutput.setTitle('ホーム画面');
    return htmlOutput;
  }
  else if (e.parameter.main || e.parameter.mainTop) {
    // ホーム画面への遷移（権限はセッションから取得、送信値は使わない）
    const _s = resolveSessionFromPost(e);
    const template = HtmlService.createTemplateFromFile('home');
    template.deployURL = ScriptApp.getService().getUrl();
    template.sessionId = _s.sessionId;
    template.userRole = _s.userRole;
    const htmlOutput = template.evaluate();
    htmlOutput.addMetaTag('viewport', 'width=device-width, initial-scale=1');
    htmlOutput.setTitle('ホーム画面');
    return htmlOutput;
  }
  // 電話受注モード（権限はセッションで検証）
  else if (e.parameter.phoneOrder) {
    const _s = resolveSessionFromPost(e);
    if (_s.userRole === 'viewer') {
      return redirectToHome(_s.userRole, _s.sessionId);
    }
    const template = HtmlService.createTemplateFromFile('phoneOrder');
    template.deployURL = ScriptApp.getService().getUrl();
    template.sessionId = _s.sessionId;
    template.userRole = _s.userRole;
    const htmlOutput = template.evaluate();
    htmlOutput.addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no');
    htmlOutput.setTitle('電話受注モード');
    return htmlOutput;
  }
  // 「受注」ボタンが押されたらshipping.htmlを返す（権限はセッションで検証）
  else if (e.parameter.shipping) {
    const _s = resolveSessionFromPost(e);
    if (_s.userRole === 'viewer') {
      return redirectToHome(_s.userRole, _s.sessionId);
    }
    const template = HtmlService.createTemplateFromFile('shipping');
    template.deployURL = ScriptApp.getService().getUrl();
    template.sessionId = _s.sessionId;
    template.userRole = _s.userRole;

    // 編集モードの設定
    const editOrderId = e.parameter.editOrderId || '';
    const editMode = e.parameter.editMode === 'true';
    const copyMode = e.parameter.actionMode === 'inherit';
    template.isEditMode = editMode;
    template.isCopyMode = copyMode;
    template.editOrderId = editOrderId;

    // 検索条件引継用
    template.prevPeriod = e.parameter.prevPeriod || '';
    template.prevDateFrom = e.parameter.prevDateFrom || '';
    template.prevDateTo = e.parameter.prevDateTo || '';
    template.prevDestination = e.parameter.prevDestination || '';
    template.prevCustomer = e.parameter.prevCustomer || '';
    template.prevStatus = e.parameter.prevStatus || '';
    template.prevIncludeOverdue = e.parameter.prevIncludeOverdue || '';

    // AI取込一覧からの遷移チェック
    const tempOrderId = e.parameter.tempOrderId || '';
    template.tempOrderId = tempOrderId;

    if (tempOrderId) {
      // AI取込一覧から遷移した場合
      const tempOrderData = getTempOrderData(tempOrderId);
      if (tempOrderData && tempOrderData.analysisResult) {
        // 解析結果をテンプレートに渡す
        // analysisResult.data に実際の解析データがある場合はそちらを使用
        const analysisData = tempOrderData.analysisResult.data || tempOrderData.analysisResult;
        template.aiAnalysisResult = JSON.stringify(analysisData);
        template.autoOpenAI = true;
      } else {
        template.autoOpenAI = false;
        template.aiAnalysisResult = '';
      }
      template.shippingHTML = getshippingHTML(e);
    } else {
      // 通常の受注画面
      template.autoOpenAI = false;
      template.aiAnalysisResult = '';
      template.shippingHTML = getshippingHTML(e);
    }

    const htmlOutput = template.evaluate();
    htmlOutput.addMetaTag('viewport', 'width=device-width, initial-scale=1');
    htmlOutput.setTitle(editMode ? '受注修正画面' : '受注画面（AI入力）');
    return htmlOutput;
  }
  // 「受注確認」ボタンが押されたらconfirm.htmlを返す（権限はセッションで検証）
  else if (e.parameter.shippingComfirm) {
    const _s = resolveSessionFromPost(e);
    if (_s.userRole === 'viewer') {
      return redirectToHome(_s.userRole, _s.sessionId);
    }
    if (isZero(e)) {
      const template = HtmlService.createTemplateFromFile('shipping');
      const alert = '少なくとも1個以上注文してください。';
      template.deployURL = ScriptApp.getService().getUrl();
      template.sessionId = _s.sessionId;
      template.userRole = _s.userRole;
      template.shippingHTML = getshippingHTML(e, alert);
      template.autoOpenAI = false;
      template.aiAnalysisResult = '';
      template.isEditMode = e.parameter.editMode === 'true';
      template.isCopyMode = e.parameter.actionMode === 'inherit';
      template.editOrderId = e.parameter.editOrderId || '';
      template.tempOrderId = e.parameter.tempOrderId || '';
      // 検索条件引継用
      template.prevPeriod = e.parameter.prevPeriod || '';
      template.prevDateFrom = e.parameter.prevDateFrom || '';
      template.prevDateTo = e.parameter.prevDateTo || '';
      template.prevDestination = e.parameter.prevDestination || '';
      template.prevCustomer = e.parameter.prevCustomer || '';
      template.prevStatus = e.parameter.prevStatus || '';
      template.prevIncludeOverdue = e.parameter.prevIncludeOverdue || '';
      const htmlOutput = template.evaluate();
      htmlOutput.addMetaTag('viewport', 'width=device-width, initial-scale=1');
      htmlOutput.setTitle('受注画面');
      return htmlOutput;
    }

    const template = HtmlService.createTemplateFromFile('shippingComfirm');
    template.deployURL = ScriptApp.getService().getUrl();
    template.sessionId = _s.sessionId;
    template.userRole = _s.userRole;
    // 編集モードの設定を追加
    const editOrderId = e.parameter.editOrderId || '';
    const editMode = e.parameter.editMode === 'true';
    template.isEditMode = editMode;
    template.editOrderId = editOrderId;
    // 仮受注IDを引き継ぎ
    template.tempOrderId = e.parameter.tempOrderId || '';
    // 検索条件引継用
    template.prevPeriod = e.parameter.prevPeriod || '';
    template.prevDateFrom = e.parameter.prevDateFrom || '';
    template.prevDateTo = e.parameter.prevDateTo || '';
    template.prevDestination = e.parameter.prevDestination || '';
    template.prevCustomer = e.parameter.prevCustomer || '';
    template.prevStatus = e.parameter.prevStatus || '';
    template.prevIncludeOverdue = e.parameter.prevIncludeOverdue || '';
    template.confirmHTML = getShippingComfirmHTML(e);
    const htmlOutput = template.evaluate();
    htmlOutput.addMetaTag('viewport', 'width=device-width, initial-scale=1');
    htmlOutput.setTitle('確認画面');
    return htmlOutput;
  }
  // 確認画面で「修正する」ボタンが押されたらform.htmlを返す（権限はセッションで検証）
  else if (e.parameter.shippingModify) {
    const _s = resolveSessionFromPost(e);
    if (_s.userRole === 'viewer') {
      return redirectToHome(_s.userRole, _s.sessionId);
    }
    const template = HtmlService.createTemplateFromFile('shipping');
    template.deployURL = ScriptApp.getService().getUrl();
    template.sessionId = _s.sessionId;
    template.userRole = _s.userRole;

    // 編集モードの設定を追加
    const editOrderId = e.parameter.editOrderId || '';
    const editMode = e.parameter.editMode === 'true';
    template.isEditMode = editMode;
    template.isCopyMode = false;
    template.editOrderId = editOrderId;
    // 仮受注IDを引き継ぎ
    template.tempOrderId = e.parameter.tempOrderId || '';

    // fromConfirmフラグを設定（再検索防止）
    e.parameter.fromConfirm = 'true';

    template.shippingHTML = getshippingHTML(e);
    template.autoOpenAI = false;
    template.aiAnalysisResult = '';
    // 検索条件引継用
    template.prevPeriod = e.parameter.prevPeriod || '';
    template.prevDateFrom = e.parameter.prevDateFrom || '';
    template.prevDateTo = e.parameter.prevDateTo || '';
    template.prevDestination = e.parameter.prevDestination || '';
    template.prevCustomer = e.parameter.prevCustomer || '';
    template.prevStatus = e.parameter.prevStatus || '';
    template.prevIncludeOverdue = e.parameter.prevIncludeOverdue || '';
    const htmlOutput = template.evaluate();
    htmlOutput.addMetaTag('viewport', 'width=device-width, initial-scale=1');
    htmlOutput.setTitle(editMode ? '受注修正画面' : '受注画面');
    return htmlOutput;
  }
  // 確認画面で「受注する」ボタンが押されたらcomplete画面へ（権限はセッションで検証）
  else if (e.parameter.shippingSubmit) {
    const _s = resolveSessionFromPost(e);
    if (_s.userRole === 'viewer') {
      return redirectToHome(_s.userRole, _s.sessionId);
    }
    createOrder(e);
    //updateZaiko(e);
    const template = HtmlService.createTemplateFromFile('shippingComplete');
    template.deployURL = ScriptApp.getService().getUrl();
    template.sessionId = _s.sessionId;
    template.userRole = _s.userRole;
    const htmlOutput = template.evaluate();
    htmlOutput.addMetaTag('viewport', 'width=device-width, initial-scale=1');
    htmlOutput.setTitle('受注完了');
    return htmlOutput;
  }
  else if (e.parameter.createShippingSlips) {
    const _s = resolveSessionFromPost(e);
    if (_s.userRole === 'viewer') {
      return redirectToHome(_s.userRole, _s.sessionId);
    }
    const template = HtmlService.createTemplateFromFile('createShippingSlips');
    template.deployURL = ScriptApp.getService().getUrl();
    template.sessionId = _s.sessionId;
    template.userRole = _s.userRole;
    const htmlOutput = template.evaluate();
    htmlOutput.addMetaTag('viewport', 'width=device-width, initial-scale=1');
    htmlOutput.setTitle('発注伝票作成');
    return htmlOutput;
  }
  else if (e.parameter.createBill) {
    const _s = resolveSessionFromPost(e);
    if (_s.userRole === 'viewer') {
      return redirectToHome(_s.userRole, _s.sessionId);
    }
    const template = HtmlService.createTemplateFromFile('createBill');
    template.deployURL = ScriptApp.getService().getUrl();
    template.sessionId = _s.sessionId;
    template.userRole = _s.userRole;
    const htmlOutput = template.evaluate();
    htmlOutput.addMetaTag('viewport', 'width=device-width, initial-scale=1');
    htmlOutput.setTitle('請求書作成');
    return htmlOutput;
  }
  else if (e.parameter.createFreeeDeliveryNote) {
    const _s = resolveSessionFromPost(e);
    if (_s.userRole === 'viewer') {
      return redirectToHome(_s.userRole, _s.sessionId);
    }
    const template = HtmlService.createTemplateFromFile('createFreeeDeliveryNote');
    template.deployURL = ScriptApp.getService().getUrl();
    template.sessionId = _s.sessionId;
    template.userRole = _s.userRole;
    const htmlOutput = template.evaluate();
    htmlOutput.addMetaTag('viewport', 'width=device-width, initial-scale=1');
    htmlOutput.setTitle('freee納品書CSV作成');
    return htmlOutput;
  }
  else if (e.parameter.csvImport) {
    const _s = resolveSessionFromPost(e);
    if (_s.userRole === 'viewer') {
      return redirectToHome(_s.userRole, _s.sessionId);
    }
    const template = HtmlService.createTemplateFromFile('csvImport');
    template.deployURL = ScriptApp.getService().getUrl();
    template.sessionId = _s.sessionId;
    template.userRole = _s.userRole;
    const htmlOutput = template.evaluate();
    htmlOutput.addMetaTag('viewport', 'width=device-width, initial-scale=1');
    htmlOutput.setTitle('受注CSV取込');
    return htmlOutput;
  }
  else if (e.parameter.masterImport) {
    const _s = resolveSessionFromPost(e);
    if (_s.userRole === 'viewer') {
      return redirectToHome(_s.userRole, _s.sessionId);
    }
    const template = HtmlService.createTemplateFromFile('masterImport');
    template.deployURL = ScriptApp.getService().getUrl();
    template.sessionId = _s.sessionId;
    template.userRole = _s.userRole;
    const htmlOutput = template.evaluate();
    htmlOutput.addMetaTag('viewport', 'width=device-width, initial-scale=1');
    htmlOutput.setTitle('一括マスタ登録');
    return htmlOutput;
  }
  else if (e.parameter.quotation) {
    const _s = resolveSessionFromPost(e);
    if (_s.userRole === 'viewer') {
      return redirectToHome(_s.userRole, _s.sessionId);
    }
    const template = HtmlService.createTemplateFromFile('quotation');
    template.deployURL = ScriptApp.getService().getUrl();
    template.sessionId = _s.sessionId;
    template.userRole = _s.userRole;
    const htmlOutput = template.evaluate();
    htmlOutput.addMetaTag('viewport', 'width=device-width, initial-scale=1');
    htmlOutput.setTitle('見積書作成');
    return htmlOutput;
  }
  else if (e.parameter.salesDashboard) {
    const _s = resolveSessionFromPost(e);

    // admin権限はセッションで検証（送信値は使わない）
    if (_s.userRole !== 'admin') {
      const errorTemplate = HtmlService.createTemplateFromFile('error');
      errorTemplate.errorMessage = 'この機能はadmin権限が必要です。';
      errorTemplate.deployURL = ScriptApp.getService().getUrl();
      errorTemplate.userRole = _s.userRole;
      errorTemplate.sessionId = _s.sessionId;
      return errorTemplate.evaluate()
        .setTitle('アクセス拒否')
        .addMetaTag('viewport', 'width=device-width, initial-scale=1');
    }

    const template = HtmlService.createTemplateFromFile('salesDashboard');
    template.deployURL = ScriptApp.getService().getUrl();
    template.sessionId = _s.sessionId;
    template.userRole = _s.userRole;
    return template.evaluate()
      .setTitle('売上ダッシュボード')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
  else if (e.parameter.orderList) {
    const _s = resolveSessionFromPost(e);
    const template = HtmlService.createTemplateFromFile('HeatmapOrderList');
    template.deployURL = ScriptApp.getService().getUrl();
    template.sessionId = _s.sessionId;
    template.userRole = _s.userRole;
    return template.evaluate()
      .setTitle('製造数一覧')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
  else if (e.parameter.manufacturingOrder) {
    const _s = resolveSessionFromPost(e);
    const page = HtmlService.createTemplateFromFile('manufacturingOrder');
    page.deployURL = ScriptApp.getService().getUrl();
    page.sessionId = _s.sessionId;
    page.userRole = _s.userRole;
    return page.evaluate()
      .setTitle('製造指示書')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
  else if (e.parameter.orderListPage) {
    const _s = resolveSessionFromPost(e);
    const template = HtmlService.createTemplateFromFile('orderList');
    template.deployURL = ScriptApp.getService().getUrl();
    template.sessionId = _s.sessionId;
    template.userRole = _s.userRole;

    // 検索条件引継用
    template.prevPeriod = e.parameter.prevPeriod || '';
    template.prevDateFrom = e.parameter.prevDateFrom || '';
    template.prevDateTo = e.parameter.prevDateTo || '';
    template.prevDestination = e.parameter.prevDestination || '';
    template.prevCustomer = e.parameter.prevCustomer || '';
    template.prevStatus = e.parameter.prevStatus || '';
    template.prevIncludeOverdue = e.parameter.prevIncludeOverdue || '';

    return template.evaluate()
      .setTitle('受注一覧')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
  // AI取込一覧画面（viewerはアクセス不可、権限はセッションで検証）
  else if (e.parameter.aiImportList) {
    const _s = resolveSessionFromPost(e);
    if (_s.userRole === 'viewer') {
      return redirectToHome(_s.userRole, _s.sessionId);
    }
    const template = HtmlService.createTemplateFromFile('aiImportList');
    template.deployURL = ScriptApp.getService().getUrl();
    template.sessionId = _s.sessionId;
    template.userRole = _s.userRole;
    return template.evaluate()
      .setTitle('AI取込一覧')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
  // 定期便一覧画面（viewerはアクセス不可、権限はセッションで検証）
  else if (e.parameter.recurringOrderList) {
    const _s = resolveSessionFromPost(e);
    if (_s.userRole === 'viewer') {
      return redirectToHome(_s.userRole, _s.sessionId);
    }
    const template = HtmlService.createTemplateFromFile('recurringOrderList');
    template.deployURL = ScriptApp.getService().getUrl();
    template.sessionId = _s.sessionId;
    template.userRole = _s.userRole;
    return template.evaluate()
      .setTitle('定期便一覧')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
  // 商品マスタ一覧画面（viewerはアクセス不可、権限はセッションで検証）
  else if (e.parameter.productList) {
    const _s = resolveSessionFromPost(e);
    if (_s.userRole === 'viewer') {
      return redirectToHome(_s.userRole, _s.sessionId);
    }
    const template = HtmlService.createTemplateFromFile('productList');
    template.deployURL = ScriptApp.getService().getUrl();
    template.sessionId = _s.sessionId;
    template.userRole = _s.userRole;
    return template.evaluate()
      .setTitle('商品マスタ')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
  // 顧客マスタ一覧画面（viewerはアクセス不可、権限はセッションで検証）
  else if (e.parameter.customerList) {
    const _s = resolveSessionFromPost(e);
    if (_s.userRole === 'viewer') {
      return redirectToHome(_s.userRole, _s.sessionId);
    }
    const template = HtmlService.createTemplateFromFile('customerList');
    template.deployURL = ScriptApp.getService().getUrl();
    template.sessionId = _s.sessionId;
    template.userRole = _s.userRole;
    return template.evaluate()
      .setTitle('顧客マスタ')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
  // 発送先マスタ一覧画面（viewerはアクセス不可、権限はセッションで検証）
  else if (e.parameter.shippingToList) {
    const _s = resolveSessionFromPost(e);
    if (_s.userRole === 'viewer') {
      return redirectToHome(_s.userRole, _s.sessionId);
    }
    const template = HtmlService.createTemplateFromFile('shippingToList');
    template.deployURL = ScriptApp.getService().getUrl();
    template.sessionId = _s.sessionId;
    template.userRole = _s.userRole;
    return template.evaluate()
      .setTitle('発送先マスタ')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
  else {
    const template = HtmlService.createTemplateFromFile('error');
    template.deployURL = ScriptApp.getService().getUrl();
    const htmlOutput = template.evaluate();
    htmlOutput.addMetaTag('viewport', 'width=device-width, initial-scale=1');
    htmlOutput.setTitle('エラー画面');
    return htmlOutput;
  }
}

/**
 * 仮受注IDから受注画面HTMLを生成（AI取込一覧からの遷移用）
 *
 * LINE BotスプレッドシートまたはAI取込一覧から仮受注データを取得し、
 * AI解析結果を受注画面のフォームに展開するHTMLを生成します。
 * 仮受注データには顧客情報、発送先情報、商品情報が含まれます。
 *
 * 処理フロー:
 * 1. LINE BotスプレッドシートID取得（getLineBotSpreadsheetId()）
 * 2. '仮受注'シートを開く
 * 3. 仮受注IDで検索（シートの1列目）
 * 4. 該当レコードの解析結果JSON（8列目）を取得
 * 5. buildMockParameterFromAnalysis() でフォームパラメータに変換
 * 6. getshippingHTML() で受注画面HTMLを生成
 * 7. 該当データなしの場合: エラーメッセージ付きHTMLを返却
 *
 * @param {string} tempOrderId - 仮受注ID（例: "temp_abc123"）
 * @returns {string} 受注画面のHTML文字列（フォームにAI解析データが展開済み）
 *
 * 仮受注シート構造:
 * - 列A: 仮受注ID
 * - 列B: 登録日時
 * - 列C: ステータス
 * - 列D: 顧客名
 * - 列E: 発送先名
 * - 列F: 商品JSON
 * - 列G: ファイル名
 * - 列H: 解析結果JSON
 * - 列I: ファイルURL
 *
 * @see buildMockParameterFromAnalysis() - AI解析結果からフォームパラメータ生成
 * @see getshippingHTML() - 受注画面HTML生成
 * @see getTempOrderData() - 仮受注データ取得
 * @see getLineBotSpreadsheetId() - LINE Botスプレッドシート ID取得（config.js）
 * @see doPost() - 呼び出し元（tempOrderIdパラメータがある場合）
 *
 * 呼び出し元: doPost() の tempOrderId パラメータ処理
 */
function getshippingHTMLForTempOrder(tempOrderId) {
  // LINE Bot側のスプレッドシートから仮受注データを取得
  const ss = SpreadsheetApp.openById(getLineBotSpreadsheetId());
  const sheet = ss.getSheetByName('仮受注');

  if (!sheet) {
    return getshippingHTML({ parameter: {} }, '仮受注データが見つかりません');
  }

  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  // 仮受注IDで検索
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === tempOrderId) {
      const analysisResult = JSON.parse(data[i][7]); // 解析結果JSON

      // eパラメータを模擬して既存関数を利用
      const mockParam = buildMockParameterFromAnalysis(analysisResult);
      return getshippingHTML({
        parameter: mockParam,
        parameters: {}  // 配列パラメータ用（チェックボックスなど）
      });
    }
  }

  return getshippingHTML({
    parameter: {},
    parameters: {}
  }, '該当する仮受注が見つかりません');
}

/**
 * AI解析結果からフォームパラメータオブジェクトを生成（仮受注データ展開用）
 *
 * AI（Gemini）の画像解析結果を受注画面のフォーム入力形式に変換します。
 * 顧客情報、発送先情報、日付、商品情報（最大10件）を含むパラメータオブジェクトを生成し、
 * getshippingHTML() に渡すことで受注画面にデータを展開します。
 *
 * 処理フロー:
 * 1. fromTempOrder フラグを設定（仮受注からの遷移であることを示す）
 * 2. analysis.customer から顧客情報を抽出:
 *    - masterData があればマスタデータ使用、なければrawデータ使用
 *    - displayName, zipcode, address, tel
 * 3. analysis.shippingTo から発送先情報を抽出:
 *    - masterData があればマスタデータ使用、なければrawデータ使用
 *    - companyName, zipcode, address, tel
 * 4. analysis.shippingDate, deliveryDate を設定
 * 5. analysis.items から商品情報を抽出（最大10件）:
 *    - bunrui1-10: 商品分類
 *    - product1-10: 商品名
 *    - quantity1-10: 個数
 *    - price1-10: 価格
 *
 * @param {Object} analysis - AI解析結果オブジェクト
 * @param {Object} analysis.customer - 顧客情報
 * @param {Object} analysis.customer.masterData - マスタデータ（displayName, zipcode, address, tel）
 * @param {string} analysis.customer.rawCompanyName - 生の会社名（マスタなし時）
 * @param {Object} analysis.shippingTo - 発送先情報
 * @param {Object} analysis.shippingTo.masterData - マスタデータ（companyName, zipcode, address, tel）
 * @param {string} analysis.shippingTo.rawCompanyName - 生の会社名（マスタなし時）
 * @param {string} analysis.shippingTo.rawZipcode - 生の郵便番号（マスタなし時）
 * @param {string} analysis.shippingTo.rawAddress - 生の住所（マスタなし時）
 * @param {string} analysis.shippingTo.rawTel - 生の電話番号（マスタなし時）
 * @param {string} analysis.shippingDate - 発送日（yyyy/MM/dd）
 * @param {string} analysis.deliveryDate - 納品日（yyyy/MM/dd）
 * @param {Array} analysis.items - 商品情報配列
 * @param {string} analysis.items[].category - 商品分類
 * @param {string} analysis.items[].productName - 商品名
 * @param {number} analysis.items[].quantity - 個数
 * @param {number} analysis.items[].price - 価格
 * @returns {Object} フォームパラメータオブジェクト
 *   {
 *     fromTempOrder: 'true',
 *     customerName: string, customerZipcode: string, customerAddress: string, customerTel: string,
 *     shippingToName: string, shippingToZipcode: string, shippingToAddress: string, shippingToTel: string,
 *     shippingDate: string, deliveryDate: string,
 *     bunrui1-10: string, product1-10: string, quantity1-10: number, price1-10: number
 *   }
 *
 * @see getshippingHTMLForTempOrder() - 呼び出し元
 * @see getshippingHTML() - 生成したパラメータを使用してHTML生成
 * @see getTempOrderData() - 仮受注データ取得
 *
 * 呼び出し元: getshippingHTMLForTempOrder()
 */
function buildMockParameterFromAnalysis(analysis) {
  const param = {
    fromTempOrder: 'true'
  };

  // 顧客情報
  if (analysis.customer) {
    param.customerName = analysis.customer.masterData?.displayName ||
      analysis.customer.rawCompanyName || '';
    param.customerZipcode = analysis.customer.masterData?.zipcode || '';
    param.customerAddress = analysis.customer.masterData?.address || '';
    param.customerTel = analysis.customer.masterData?.tel || '';
  }

  // 発送先情報
  if (analysis.shippingTo) {
    param.shippingToName = analysis.shippingTo.masterData?.companyName ||
      analysis.shippingTo.rawCompanyName || '';
    param.shippingToZipcode = analysis.shippingTo.masterData?.zipcode ||
      analysis.shippingTo.rawZipcode || '';
    param.shippingToAddress = analysis.shippingTo.masterData?.address ||
      analysis.shippingTo.rawAddress || '';
    param.shippingToTel = analysis.shippingTo.masterData?.tel ||
      analysis.shippingTo.rawTel || '';
  }

  // 日付
  param.shippingDate = analysis.shippingDate || '';
  param.deliveryDate = analysis.deliveryDate || '';

  // 商品情報（最大10件）
  if (analysis.items && analysis.items.length > 0) {
    analysis.items.slice(0, 10).forEach((item, i) => {
      const num = i + 1;
      param['bunrui' + num] = item.category || '';
      param['product' + num] = item.productName || '';
      param['quantity' + num] = item.quantity || '';
      param['price' + num] = item.price || '';
    });
  }

  return param;
}

/**
 * 仮受注IDから仮受注データを取得（AI解析結果を含む）
 *
 * LINE BotスプレッドシートまたはAI取込一覧から仮受注データを検索し、
 * AI解析結果を含む完全なデータオブジェクトを返却します。
 * 受注画面へのデータ展開時に使用されます。
 *
 * 処理フロー:
 * 1. LINE BotスプレッドシートID取得（getLineBotSpreadsheetId()）
 * 2. '仮受注'シートを開く
 * 3. 仮受注IDで全行を検索（列A）
 * 4. 該当レコードが見つかった場合:
 *    - 列Hの解析結果JSON文字列をパース
 *    - 全カラムの情報を含むオブジェクトを返却
 * 5. 該当レコードなしまたはエラー: null返却
 *
 * @param {string} tempOrderId - 仮受注ID（例: "temp_abc123"）
 * @returns {Object|null} 仮受注データオブジェクト、見つからない場合はnull
 *   {
 *     tempOrderId: string,        // 仮受注ID
 *     registeredAt: Date,         // 登録日時
 *     status: string,             // ステータス（例: "未処理", "処理済み"）
 *     customerName: string,       // 顧客名
 *     shippingToName: string,     // 発送先名
 *     itemsJson: string,          // 商品情報JSON文字列
 *     fileName: string,           // 元ファイル名
 *     analysisResult: Object,     // AI解析結果オブジェクト（JSONパース済み）
 *     fileUrl: string             // ファイルURL
 *   }
 *
 * 仮受注シート構造:
 * - 列A: 仮受注ID
 * - 列B: 登録日時
 * - 列C: ステータス
 * - 列D: 顧客名
 * - 列E: 発送先名
 * - 列F: 商品JSON
 * - 列G: ファイル名
 * - 列H: 解析結果JSON
 * - 列I: ファイルURL
 *
 * @see getshippingHTMLForTempOrder() - 呼び出し元（HTML生成時）
 * @see doPost() - 呼び出し元（受注画面遷移時）
 * @see getLineBotSpreadsheetId() - LINE Botスプレッドシート ID取得（config.js）
 *
 * 呼び出し元: doPost() の shipping パラメータ処理、getshippingHTMLForTempOrder()
 */
function getTempOrderData(tempOrderId) {
  try {
    const ss = SpreadsheetApp.openById(getLineBotSpreadsheetId());
    const sheet = ss.getSheetByName('仮受注');

    if (!sheet) {
      Logger.log('仮受注シートが見つかりません');
      return null;
    }

    const data = sheet.getDataRange().getValues();
    const headers = data[0];

    // 仮受注IDで検索
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === tempOrderId) {
        const analysisResultJson = data[i][7]; // 解析結果JSON（列H）

        if (!analysisResultJson) {
          Logger.log('解析結果が空です: ' + tempOrderId);
          return null;
        }

        const analysisResult = JSON.parse(analysisResultJson);

        return {
          tempOrderId: tempOrderId,
          registeredAt: data[i][1],
          status: data[i][2],
          customerName: data[i][3],
          shippingToName: data[i][4],
          itemsJson: data[i][5],
          fileName: data[i][6],
          analysisResult: analysisResult,
          fileUrl: data[i][8]
        };
      }
    }

    Logger.log('該当する仮受注が見つかりません: ' + tempOrderId);
    return null;

  } catch (error) {
    Logger.log('getTempOrderData エラー: ' + error.toString());
    return null;
  }
}

/**
 * スプレッドシートのシート名からヘッダ付きレコード配列を取得
 *
 * 指定されたシート名の全レコードを、ヘッダ行をキーとした連想配列の配列形式で取得します。
 * 内部的に getAllRecordsInternal() を呼び出し、オブジェクト形式で返却します。
 *
 * 処理フロー:
 * 1. getAllRecordsInternal(sheetName, false) を呼び出し
 * 2. シートの1行目をヘッダとして取得
 * 3. 2行目以降の各行を { ヘッダ名: セル値 } の連想配列に変換
 * 4. 連想配列の配列を返却
 *
 * @param {string} sheetName - シート名（例: "受注", "顧客情報", "商品"）
 * @returns {Array<Object>} レコード配列
 *   [
 *     { ヘッダ1: 値1, ヘッダ2: 値2, ... },
 *     { ヘッダ1: 値1, ヘッダ2: 値2, ... },
 *     ...
 *   ]
 *
 * 使用例:
 * const orders = getAllRecords('受注');
 * // 返却: [{ '受注ID': 'ORD001', '発送先名': '株式会社ABC', ... }, ...]
 *
 * @see getAllRecordsInternal() - 内部実装（JSON文字列/オブジェクト形式切り替え可能）
 * @see getLastData() - 前回商品反映機能で使用
 * @see customerSearch() - 顧客検索で使用
 * @see shippingToSearch() - 発送先検索で使用
 *
 * 呼び出し元: customerCode.js, orderCode.js, quotationCode.js など多数
 */
function getAllRecords(sheetName) {
  return getAllRecordsInternal(sheetName, false);
}

/**
 * スプレッドシート全レコード取得（内部実装）
 *
 * 指定されたシート名の全レコードを取得し、ヘッダ行をキーとした連想配列形式で返却します。
 * flgパラメータによってJSON文字列またはオブジェクト配列を選択できます。
 * Webアプリ実行時はマスタスプレッドシートIDを使用してシートを開きます。
 *
 * 処理フロー:
 * 1. アクティブスプレッドシート取得試行
 * 2. 取得できない場合（Webアプリ実行時）: getMasterSpreadsheetId() でマスタIDを取得
 * 3. 指定されたシート名のシートを取得
 * 4. シートの全データを取得（getDataRange().getValues()）
 * 5. 1行目をヘッダとして取得（shift()）
 * 6. 2行目以降の各行を { ヘッダ名: セル値 } の連想配列に変換
 * 7. flg=true の場合: JSON.stringify() で文字列化
 * 8. flg=false の場合: オブジェクト配列をそのまま返却
 *
 * @param {string} sheetName - シート名（例: "受注", "顧客情報", "商品"）
 * @param {boolean} flg - 返却形式フラグ（true=JSON文字列, false=オブジェクト配列）
 * @returns {Array<Object>|string} レコード配列またはJSON文字列、シートなしの場合は [] または '[]'
 *   flg=false: [{ ヘッダ1: 値1, ... }, ...]
 *   flg=true: '[{"ヘッダ1":"値1",...},...]'
 *
 * エラーハンドリング:
 * - スプレッドシートを開けない: [] または '[]' を返却
 * - シートが見つからない: [] または '[]' を返却
 * - Logger.log() でエラー内容を記録
 *
 * @see getAllRecords() - 公開API（flg=false で呼び出し）
 * @see getMasterSpreadsheetId() - マスタスプレッドシートID取得（config.js）
 *
 * 呼び出し元: getAllRecords()
 */
function getAllRecordsInternal(sheetName, flg) {
  // WebアプリではgetActiveSpreadsheet()がnullになるため、マスタIDを使用
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    const masterSpreadsheetId = getMasterSpreadsheetId();
    if (masterSpreadsheetId) {
      ss = SpreadsheetApp.openById(masterSpreadsheetId);
    } else {
      Logger.log('getAllRecords: スプレッドシートを開けませんでした');
      return flg ? '[]' : [];
    }
  }

  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    Logger.log('getAllRecords: シート「' + sheetName + '」が見つかりません');
    return flg ? '[]' : [];
  }

  const values = sheet.getDataRange().getValues();
  const labels = values.shift();

  const records = [];
  for (const value of values) {
    const record = {};
    labels.forEach((label, index) => {
      record[label] = value[index];
    });
    records.push(record);
  }
  if (flg) {
    return JSON.stringify(records);
  } else {
    return records;
  }
}

/**
 * 商品個数のゼロ入力チェック（受注確認画面でのバリデーション）
 *
 * 受注画面で入力された商品の個数が全てゼロまたは未入力かチェックします。
 * quantity1～quantity8 の合計がゼロの場合、受注確認画面への遷移をブロックします。
 *
 * 処理フロー:
 * 1. quantity1～quantity8 のパラメータをループ
 * 2. 各個数を数値に変換して合計を計算
 * 3. 合計が0の場合: true を返却（エラー）
 * 4. 合計が1以上の場合: false を返却（正常）
 *
 * @param {Object} e - POSTリクエストイベントオブジェクト
 * @param {Object} e.parameter - フォームパラメータのkey-valueオブジェクト
 * @param {string} e.parameter.quantity1-8 - 商品1～8の個数
 * @returns {boolean} true=全ての個数がゼロ（エラー）, false=1個以上入力あり（正常）
 *
 * 使用例:
 * if (isZero(e)) {
 *   // エラー処理: 「少なくとも1個以上注文してください」というアラート表示
 * }
 *
 * @see doPost() - 受注確認画面遷移時にバリデーション実行
 * @see shipping.html - 受注画面フォーム
 *
 * 呼び出し元: doPost() の shippingComfirm パラメータ処理
 */
function isZero(e) {
  let total = 0;
  var rowNum = 0;
  for (let i = 0; i < 8; i++) {
    rowNum++;
    var quantity = "quantity" + rowNum;
    const count = Number(e.parameter[quantity]);
    if (count) {
      total += count;
    }
  }
  if (total == 0) {
    return true;
  }
  return false;
}

/**
 * ホーム画面にリダイレクト（権限不足時）
 * @param {string} userRole - ユーザー権限（admin/viewer）
 * @param {string} [sessionId] - セッションID（フォーム引き継ぎ用）
 * @returns {HtmlOutput} ホーム画面
 */
function redirectToHome(userRole, sessionId) {
  const template = HtmlService.createTemplateFromFile('home');
  template.deployURL = ScriptApp.getService().getUrl();
  template.sessionId = sessionId || '';
  template.userRole = userRole || 'viewer';
  const htmlOutput = template.evaluate();
  htmlOutput.addMetaTag('viewport', 'width=device-width, initial-scale=1');
  htmlOutput.setTitle('ホーム画面');
  return htmlOutput;
}
