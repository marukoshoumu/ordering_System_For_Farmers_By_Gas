/**
 * 受注シートからヘッダーを取得
 * シートのヘッダーが真実の情報源（Single Source of Truth）
 * @returns {Array} ヘッダー配列
 */
function getOrderSheetHeaders() {
  const ss = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.openById(getMasterSpreadsheetId());
  const sheet = ss.getSheetByName('受注');
  return sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
}

/**
 * 受注レコード配列から指定ヘッダーの値を取得
 * @param {Array} record - 受注レコード配列
 * @param {string} headerName - ヘッダー名
 * @param {Array} headers - ヘッダー配列
 * @returns {*} 値
 * @throws {Error} ヘッダーが見つからない場合
 */
function getOrderValue(record, headerName, headers) {
  const index = headers.indexOf(headerName);
  if (index === -1) {
    throw new Error(`受注シートヘッダー「${headerName}」が見つかりません。利用可能なヘッダー: [${headers.join(', ')}]`);
  }
  return record[index];
}

/**
 * recordオブジェクトをヘッダー順の配列に変換
 * @param {Object} record - キーが列名のオブジェクト
 * @param {Array} headers - ヘッダー配列
 * @returns {Array} ヘッダー順に並んだ値の配列
 */
function recordToArray(record, headers) {
  return headers.map(header => record[header] !== undefined ? record[header] : '');
}

/**
 * マスタデータキャッシュをクリアして再取得
 *
 * CacheService に保存されているマスタデータキャッシュを削除し、
 * 最新のマスタデータを取得して新しいキャッシュを作成します。
 *
 * 使用シーン:
 * - 商品マスタ、発送元情報を更新した後
 * - マスタデータが最新でない可能性がある場合
 * - 手動でキャッシュをリフレッシュしたい場合
 *
 * 処理フロー:
 * 1. CacheService.getScriptCache() 取得
 * 2. 'masterData_v5' キーのキャッシュを削除
 * 3. getMasterDataCached() で最新データ取得 & 新規キャッシュ作成
 * 4. 成功時: 商品数、発送元数を含む結果オブジェクト返却
 * 5. エラー時: エラーメッセージを含む結果オブジェクト返却
 *
 * @returns {Object} 実行結果オブジェクト
 *   成功時: { success: true, message: string, productCount: number, shippingFromCount: number }
 *   失敗時: { success: false, message: string }
 *
 * @see getMasterDataCached() - マスタデータ取得とキャッシュ管理
 */
function refreshMasterDataCache() {
  try {
    const cache = CacheService.getScriptCache();
    cache.remove('masterData_v5');
    const newData = getMasterDataCached();
    return {
      success: true,
      message: 'キャッシュを更新しました',
      productCount: newData.products ? newData.products.length : 0,
      shippingFromCount: newData.shippingFromList ? newData.shippingFromList.length : 0
    };
  } catch (error) {
    Logger.log('refreshMasterDataCache エラー: ' + error.toString());
    return {
      success: false,
      message: 'キャッシュ更新に失敗しました: ' + error.toString()
    };
  }
}

/**
 * 配送伝票スキャンモーダル用のHTMLコンテンツを取得
 * 
 * @param {string} orderId - 対象の受注ID
 * @param {string} customerName - 顧客名（モーダルタイトル表示用）
 * @returns {string} scanner.html を evaluate したHTML文字列
 */
function getScannerHTML(orderId, customerName) {
  const template = HtmlService.createTemplateFromFile('scanner');
  template.orderId = orderId;
  template.customerName = customerName;
  return template.evaluate()
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .getContent();
}


/**
 * マスタデータをキャッシュから取得（2時間キャッシュ、パフォーマンス最適化）
 *
 * スプレッドシートから全マスタデータを取得し、CacheService に2時間保存します。
 * キャッシュが存在する場合はスプレッドシート読み取りをスキップして高速にデータを返します。
 *
 * パフォーマンス改善:
 * - キャッシュヒット時: 数十ミリ秒（スプレッドシート読み取りなし）
 * - キャッシュミス時: 数秒（全マスタシートを読み取り）
 * - キャッシュ有効期間: 2時間（7200秒）
 *
 * 処理フロー:
 * 1. CacheService から 'masterData_v5' キーでキャッシュ取得
 * 2. キャッシュヒット & 有効なデータ構造:
 *    - JSON.parse() してそのまま返却
 * 3. キャッシュミス または 古いキャッシュ構造:
 *    - getAllRecords() で各マスタシートからデータ取得
 *    - masterData オブジェクトを構築
 *    - cache.put() で2時間キャッシュに保存
 * 4. エラー時: 空配列・空オブジェクトのデフォルト値を返却
 *
 * キャッシュキー履歴:
 * - v5: shippingFromList 追加（複数発送元対応）
 * - 古いバージョンのキャッシュは自動的に破棄される
 *
 * 返却データ構造:
 * {
 *   products: Array<Object>,         // 商品マスタ
 *   recipients: Array<Object>,        // 担当者マスタ
 *   deliveryMethods: Array<Object>,   // 納品方法マスタ
 *   receipts: Array<Object>,          // 受付方法マスタ
 *   deliveryTimes: Array<Object>,     // 配送時間帯マスタ
 *   invoiceTypes: Array<Object>,      // 送り状種別マスタ
 *   coolClss: Array<Object>,          // クール区分マスタ
 *   cargos: Array<Object>,            // 荷扱いマスタ
 *   shippingFrom: Object,             // デフォルト発送元（1件目）
 *   shippingFromList: Array<Object>   // 全発送元リスト
 * }
 *
 * @returns {Object} マスタデータオブジェクト（キャッシュまたはスプレッドシートから取得）
 *
 * @see refreshMasterDataCache() - キャッシュクリアと再取得
 * @see getAllRecords() - スプレッドシートからレコード取得
 *
 * 呼び出し元:
 * - getshippingHTML() - 受注入力画面の初期化
 * - refreshMasterDataCache() - キャッシュ更新
 */
function getMasterDataCached() {
  try {
    const cache = CacheService.getScriptCache();
    const cacheKey = 'masterData_v5';

    let cached = cache.get(cacheKey);
    if (cached) {
      const parsed = JSON.parse(cached);
      // productsとshippingFromListが含まれているか確認（古いキャッシュ対策）
      if (parsed && parsed.products && parsed.shippingFromList) {
        return parsed;
      }
    }

    // キャッシュがなければ取得
    // 会社名シート（COMPANY_DISPLAY_NAMEプロパティの値と同じ名前）から発送元情報を取得
    const companyDisplayName = getCompanyDisplayName();
    const shippingFromRecords = getAllRecords(companyDisplayName) || [];
    const shippingFromList = shippingFromRecords.map(function (r) {
      return {
        name: r['名前'] || '',
        zipcode: r['郵便番号'] || '',
        address: r['住所'] || '',
        tel: r['電話番号'] || ''
      };
    });
    const shippingFrom = shippingFromList.length > 0 ? shippingFromList[0] : { name: '', zipcode: '', address: '', tel: '' };

    const masterData = {
      products: getAllRecords('商品') || [],
      recipients: getAllRecords('担当者') || [],
      deliveryMethods: getAllRecords('納品方法') || [],
      receipts: getAllRecords('受付方法') || [],
      deliveryTimes: getAllRecords('配送時間帯') || [],
      invoiceTypes: getAllRecords('送り状種別') || [],
      coolClss: getAllRecords('クール区分') || [],
      cargos: getAllRecords('荷扱い') || [],
      shippingFrom: shippingFrom,
      shippingFromList: shippingFromList
    };

    // 2時間キャッシュ（7200秒）
    try {
      cache.put(cacheKey, JSON.stringify(masterData), 7200);
    } catch (cacheError) {
      Logger.log('キャッシュ保存エラー: ' + cacheError.toString());
    }
    return masterData;
  } catch (error) {
    Logger.log('getMasterDataCached エラー: ' + error.toString());
    // エラー時はデフォルト値を返す
    return {
      products: [],
      recipients: [],
      deliveryMethods: [],
      receipts: [],
      deliveryTimes: [],
      invoiceTypes: [],
      coolClss: [],
      cargos: [],
      shippingFrom: { name: '', zipcode: '', address: '', tel: '' },
      shippingFromList: []
    };
  }
}
/**
 * 受注入力画面（shipping.html）のメイン生成関数
 * 
 * 編集モード、引き継ぎモード、新規入力の各状態に応じてフォームの初期値を設定し、
 * 受注入力画面のHTMLパーツを構築します。
 * 
 * @param {Object} e - POSTリクエストイベントオブジェクト
 * @param {string} alert - エラーメッセージ（バリデーション失敗時など）
 * @returns {string} 受注入力フォームのHTML文字列
 * 
 * @see getOrderByOrderId - 編集モード時のデータ取得
 * @see getMasterDataCached - マスタデータの取得
 */
function getshippingHTML(e, alert = '') {

  // 編集モードの判定
  const editOrderId = e.parameter.editOrderId || '';
  const actionMode = e.parameter.actionMode || '';
  const isInheritMode = actionMode === 'inherit';
  let editData = null;

  if (editOrderId && !e.parameter.fromConfirm) {
    editData = getOrderByOrderId(editOrderId);
  }

  if (editData) {
    e.parameter.shippingToName = editData.shippingToName;
    e.parameter.shippingToZipcode = editData.shippingToZipcode;
    e.parameter.shippingToAddress = editData.shippingToAddress;
    e.parameter.shippingToTel = editData.shippingToTel;
    e.parameter.customerName = editData.customerName;
    e.parameter.customerZipcode = editData.customerZipcode;
    e.parameter.customerAddress = editData.customerAddress;
    e.parameter.customerTel = editData.customerTel;
    e.parameter.shippingFromName = editData.shippingFromName;
    e.parameter.shippingFromZipcode = editData.shippingFromZipcode;
    e.parameter.shippingFromAddress = editData.shippingFromAddress;
    e.parameter.shippingFromTel = editData.shippingFromTel;

    // 複数日程対応: dates配列からデータを設定
    if (isInheritMode) {
      e.parameter.shippingDate1 = '';
      e.parameter.deliveryDate1 = '';
      // コピー時も納品書テキストは引継ぐ（1日程目を deliveryNoteText1 に設定）
      e.parameter.deliveryNoteText1 = (editData.dates && editData.dates[0]) ? (editData.dates[0].deliveryNoteText || '') : '';
    } else {
      // 後方互換性: 旧形式（dates配列がない場合）にも対応
      if (editData.dates && editData.dates.length > 0) {
        // 複数日程の場合: dates配列から設定
        editData.dates.forEach((dateInfo, index) => {
          const dateNum = index + 1;
          e.parameter['shippingDate' + dateNum] = dateInfo.shippingDate;
          e.parameter['deliveryDate' + dateNum] = dateInfo.deliveryDate;
          e.parameter['deliveryNoteText' + dateNum] = dateInfo.deliveryNoteText || '';
        });
      } else {
        // 旧形式の場合: トップレベルのプロパティから設定
        e.parameter.shippingDate1 = editData.shippingDate;
        e.parameter.deliveryDate1 = editData.deliveryDate;
      }
    }

    e.parameter.receiptWay = editData.receiptWay;
    e.parameter.recipient = editData.recipient;
    e.parameter.deliveryMethod = editData.deliveryMethod;
    e.parameter.deliveryTime = editData.deliveryTime;

    e.parameters = e.parameters || {};
    e.parameters.checklist = [];
    if (editData.checklist.deliverySlip) e.parameters.checklist.push('納品書');
    if (editData.checklist.bill) e.parameters.checklist.push('請求書');
    if (editData.checklist.receipt) e.parameters.checklist.push('領収書');
    if (editData.checklist.pamphlet) e.parameters.checklist.push('パンフ');
    if (editData.checklist.recipe) e.parameters.checklist.push('レシピ');

    e.parameter.otherAttach = editData.otherAttach;

    editData.items.forEach((item, index) => {
      const rowNum = index + 1;
      e.parameter['bunrui' + rowNum] = item.bunrui;
      e.parameter['product' + rowNum] = item.product;
      e.parameter['price' + rowNum] = item.price;
      if (isInheritMode) {
        e.parameter['quantity' + rowNum] = '';
      } else {
        e.parameter['quantity' + rowNum] = item.quantity;
      }
    });

    e.parameter.sendProduct = editData.sendProduct;
    e.parameter.invoiceType = editData.invoiceType;
    e.parameter.coolCls = editData.coolCls;
    e.parameter.cargo1 = editData.cargo1;
    e.parameter.cargo2 = editData.cargo2;
    e.parameter.cargo3 = editData.cargo3;
    e.parameter.cashOnDelivery = editData.cashOnDelivery;
    e.parameter.cashOnDeliTax = editData.cashOnDeliTax;
    e.parameter.copiePrint = editData.copiePrint;
    e.parameter.csvmemo = editData.csvmemo;
    e.parameter.internalMemo = editData.internalMemo;
    e.parameter.deliveryMemo = editData.deliveryMemo;
    e.parameter.memo = editData.memo;
  }

  const master = getMasterDataCached();
  const items = getAllRecords('商品');
  const recipients = master.recipients;
  const deliveryMethods = master.deliveryMethods;
  const receipts = master.receipts;
  const deliveryTimes = master.deliveryTimes;
  const invoiceTypes = master.invoiceTypes;
  const coolClss = master.coolClss;
  const cargos = master.cargos;

  var nDate = new Date();
  var strDate = Utilities.formatDate(nDate, 'JST', 'yyyy-MM-dd')
  const orderDate = e.parameter.orderDate ? e.parameter.orderDate : strDate;
  var n = 2;
  nDate.setDate(nDate.getDate() + n);
  strDate = Utilities.formatDate(nDate, 'JST', 'yyyy-MM-dd')
  const shippingDate = e.parameter.shippingDate ? e.parameter.shippingDate : strDate;
  nDate.setDate(nDate.getDate() + 1);
  strDate = Utilities.formatDate(nDate, 'JST', 'yyyy-MM-dd')
  const deliveryDate = e.parameter.deliveryDate ? e.parameter.deliveryDate : strDate;

  // 受注シートのヘッダー更新（追跡番号がない場合のみ追加）
  setupOrderSheetHeaders();

  // ============================================
  // CSS スタイル
  // ============================================
  let html = `
<style>
/* セクションヘッダー共通スタイル */
.section-header {
  color: white;
  padding: 8px 12px;
  margin-top: 16px;
  border-radius: 6px 6px 0 0;
}
.section-header:first-of-type {
  margin-top: 0;
}
.section-header-row {
  display: flex;
  flex-wrap: wrap;
  align-items: center;
  gap: 6px;
  margin-bottom: 4px;
}
.section-header-row:last-child {
  margin-bottom: 0;
}
.section-header-label {
  font-weight: bold;
  font-size: 0.95rem;
  display: flex;
  align-items: center;
  gap: 6px;
  margin-right: 8px;
}
.section-header button {
  padding: 4px 10px;
  font-size: 0.85rem;
  border: 1px solid rgba(255,255,255,0.5);
  background: rgba(255,255,255,0.15);
  color: white;
  border-radius: 4px;
  cursor: pointer;
  white-space: nowrap;
  transition: background 0.2s;
}
.section-header button:hover {
  background: rgba(255,255,255,0.3);
}
.section-header button:disabled {
  opacity: 0.6;
  cursor: not-allowed;
}
.section-header select {
  padding: 4px 8px;
  font-size: 0.85rem;
  border-radius: 4px;
  border: 1px solid #ccc;
  background: white;
  max-width: 280px;
  color: #333;
}

/* セクション本体 */
.section-body {
  background: #fff;
  border: 1px solid #e0e0e0;
  border-top: none;
  border-radius: 0 0 6px 6px;
  padding: 12px;
  margin-bottom: 8px;
}

/* 各セクションのカラー */
.section-shipping-to {
  background: linear-gradient(135deg, #b8860b 0%, #daa520 100%);
}
.section-customer {
  background: linear-gradient(135deg, #c71585 0%, #db7093 100%);
}
.section-shipping-from {
  background: linear-gradient(135deg, #8b4513 0%, #a0522d 100%);
}
.section-order-basic {
  background: linear-gradient(135deg, #008b8b 0%, #20b2aa 100%);
}
.section-products {
  background: linear-gradient(135deg, #228b22 0%, #32cd32 100%);
}
.section-shipping-info {
  background: linear-gradient(135deg, #1e3a5f 0%, #4169e1 100%);
}

/* サブ行（履歴など） */
.section-sub-row {
  border-top: 1px solid rgba(255,255,255,0.2);
  padding-top: 6px;
  margin-top: 4px;
}
.section-sub-label {
  font-size: 0.8rem;
  color: rgba(255,255,255,0.85);
  margin-right: 6px;
}

/* チェックボックスグループ */
.checkbox-group {
  display: flex;
  flex-wrap: wrap;
  gap: 12px;
  padding: 8px 0;
}
.checkbox-group .form-check {
  margin: 0;
}

/* レスポンシブ */
@media (max-width: 768px) {
  .section-header {
    padding: 6px 8px;
  }
  .section-header-row {
    gap: 4px;
  }
  .section-header button {
    padding: 4px 8px;
    font-size: 0.8rem;
  }
  .section-header select {
    padding: 4px 6px;
    font-size: 0.8rem;
    max-width: 180px;
  }
  .section-header-label {
    font-size: 0.9rem;
  }
  .section-body {
    padding: 10px;
  }
}
</style>
`;

  html += `<p class="text-danger">${alert}</p>`;

  // ============================================
  // 発送先情報セクション
  // ============================================
  html += `
<div class="confirm-section section-shipping-to"><div class="section-header">
  <div class="section-header-row">
    <span class="section-header-label">📦 発送先情報</span>
    <button type='button' class="shippingToInsertBtn_open" title="新規登録">➕ 新規</button>
    <button type='button' class="shippingToSearchBtn_open" title="発送先検索">🔍 検索</button>
    <button type='button' id="productSearch" onclick="setProductSearch()" title="前回商品反映">📋 前回</button>
  </div>
  <div class="section-header-row section-sub-row">
    <span class="section-sub-label">📂 履歴:</span>
    <button type='button' id="getOrderHistoryBtn" onclick="getOrderHistory()" title="過去の受注から取得">取得</button>
    <select id="orderHistorySelect" style="display:none;" onchange="onOrderHistorySelect()">
      <option value="">発送日を選択...</option>
    </select>
    <button type='button' id="applyOrderHistoryBtn" style="display:none;" onclick="applyOrderHistory()">反映</button>
  </div>
</div>
<div class="section-body">
`;
  html += `<div>
    <label for="shippingToName" class="text-left form-label">発送先名</label>
    <input type="text" class="form-control" id="shippingToName" name="shippingToName" required value="${e.parameter.shippingToName ? e.parameter.shippingToName : ""}" >
  </div>`;
  html += `<div class="mt-2 mb-2 row g-3 align-items-center">
    <div class="col-auto">
      <label for="shippingToZipcode" class="col-form-label">郵便番号</label>
    </div>
    <div class="col-auto">
      <input type="text" class="form-control" id="shippingToZipcode" name="shippingToZipcode" required value="${e.parameter.shippingToZipcode ? e.parameter.shippingToZipcode : ""}" maxlength=7 pattern="[0-9]{7}" title="7桁の数字のみを入力してください。">
    </div>
  </div>`;
  html += `<div>
    <label for="shippingToAddress" class="text-left form-label">住所</label>
    <input type="text" class="form-control" id="shippingToAddress" name="shippingToAddress" required value="${e.parameter.shippingToAddress ? e.parameter.shippingToAddress : ""}" >
  </div>`;
  html += `<div class="mt-2 mb-2 row g-3 align-items-center">
    <div class="col-auto">
      <label for="shippingToTel" class="col-form-label">電話番号</label>
    </div>
    <div class="col-auto">
      <input type="text" class="form-control" id="shippingToTel" name="shippingToTel" required value="${e.parameter.shippingToTel ? e.parameter.shippingToTel : ""}" maxlength=11 pattern="^0[0-9]{9,10}$" title="10~11桁の数字のみを入力してください。">
    </div>
  </div>`;
  html += `</div>`; // section-body閉じ

  // ============================================
  // 顧客情報セクション
  // ============================================
  html += `
<div class="confirm-section section-customer"><div class="section-header">
  <div class="section-header-row">
    <span class="section-header-label">👤 顧客情報</span>
    <button type='button' class="customerInsertBtn_open" title="新規登録">➕ 新規</button>
    <button type='button' class="customerSearchBtn_open" title="顧客検索">🔍 検索</button>
    <button type='button' class="customerSameBtn" onclick="customerSame()" title="発送先と同じ">📋 同上</button>
    <button type='button' id="quotationSearch" onclick="setQuotationSearch()" title="見積書から反映">📄 見積書</button>
  </div>
</div>
<div class="section-body">
`;
  html += `<div>
    <label for="customerName" class="text-left form-label">顧客名</label>
    <input type="text" class="form-control" id="customerName" name="customerName" required value="${e.parameter.customerName ? e.parameter.customerName : ""}" >
  </div>`;
  html += `<div class="mt-2 mb-2 row g-3 align-items-center">
    <div class="col-auto">
      <label for="customerZipcode" class="col-form-label">郵便番号</label>
    </div>
    <div class="col-auto">
      <input type="text" class="form-control" id="customerZipcode" name="customerZipcode" required value="${e.parameter.customerZipcode ? e.parameter.customerZipcode : ""}" maxlength=7 pattern="[0-9]{7}" title="7桁の数字のみを入力してください。">
    </div>
  </div>`;
  html += `<div>
    <label for="customerAddress" class="text-left form-label">住所</label>
    <input type="text" class="form-control" id="customerAddress" name="customerAddress" required value="${e.parameter.customerAddress ? e.parameter.customerAddress : ""}" >
  </div>`;
  html += `<div class="mt-2 mb-2 row g-3 align-items-center">
    <div class="col-auto">
      <label for="customerTel" class="col-form-label">電話番号</label>
    </div>
    <div class="col-auto">
      <input type="text" class="form-control" id="customerTel" name="customerTel" required value="${e.parameter.customerTel ? e.parameter.customerTel : ""}" maxlength=11 pattern="^0[0-9]{9,10}$" title="10~11桁の数字のみを入力してください。">
    </div>
  </div>`;
  html += `</div>`;

  // ============================================
  // 発送元情報セクション
  // ============================================
  const companyDisplayName = getCompanyDisplayName();
  html += `
<div class="confirm-section section-shipping-from"><div class="section-header">
  <div class="section-header-row">
    <span class="section-header-label">🏭 発送元情報</span>
    <button type="button" id="companyBtn" onclick="companyChange()" title="${companyDisplayName}の情報を入力">🌿 ${companyDisplayName}</button>
    <button type="button" id="custCopyBtn" onclick="custCopy()" title="顧客情報をコピー">📋 顧客</button>
    <button type="button" id="sendCopyBtn" onclick="sendCopy()" title="発送先情報をコピー">📋 発送先</button>
  </div>
</div>
<div class="section-body">
`;
  html += `<div>
    <label for="shippingFromName" class="text-left form-label">発送元名</label>
    <input type="text" class="form-control" id="shippingFromName" name="shippingFromName" required value="${e.parameter.shippingFromName ? e.parameter.shippingFromName : ""}" >
  </div>`;
  html += `<div class="mt-2 mb-2 row g-3 align-items-center">
    <div class="col-auto">
      <label for="shippingFromZipcode" class="col-form-label">郵便番号</label>
    </div>
    <div class="col-auto">
      <input type="text" class="form-control" id="shippingFromZipcode" name="shippingFromZipcode" required value="${e.parameter.shippingFromZipcode ? e.parameter.shippingFromZipcode : ""}" maxlength=7 pattern="[0-9]{7}" title="7桁の数字のみを入力してください。">
    </div>
  </div>`;
  html += `<div>
    <label for="shippingFromAddress" class="text-left form-label">住所</label>
    <input type="text" class="form-control" id="shippingFromAddress" name="shippingFromAddress" required value="${e.parameter.shippingFromAddress ? e.parameter.shippingFromAddress : ""}" >
  </div>`;
  html += `<div class="mt-2 mb-2 row g-3 align-items-center">
    <div class="col-auto">
      <label for="shippingFromTel" class="col-form-label">電話番号</label>
    </div>
    <div class="col-auto">
      <input type="text" class="form-control" id="shippingFromTel" name="shippingFromTel" required value="${e.parameter.shippingFromTel ? e.parameter.shippingFromTel : ""}" maxlength=11 pattern="^0[0-9]{9,10}$" title="10~11桁の数字のみを入力してください。">
    </div>
  </div>`;
  html += `</div>`;

  // ============================================
  // 受注基本情報セクション
  // ============================================
  html += `
<div class="confirm-section section-order-basic"><div class="section-header">
  <div class="section-header-row">
    <span class="section-header-label">📝 受注基本情報</span>
  </div>
</div>
<div class="section-body">
`;

  // 複数日程登録UI
  let existingDateCount = 0;
  for (let i = 1; i <= 10; i++) {
    if (e.parameter['shippingDate' + i]) {
      existingDateCount = i;
    } else {
      break;
    }
  }
  if (existingDateCount === 0) {
    existingDateCount = 1;
  }

  html += `<div class="mb-3">`;
  html += `  <div class="d-flex align-items-center mb-2">`;
  html += `    <span class="fw-bold">📅 発送日程</span>`;
  html += `    <button type="button" class="btn btn-sm btn-outline-primary ms-3" onclick="addShippingDate()">＋ 日程追加</button>`;
  html += `  </div>`;
  html += `  <div id="shippingDateContainer">`;

  for (let i = 1; i <= existingDateCount; i++) {
    const sd = e.parameter['shippingDate' + i] || (i === 1 ? shippingDate : '');
    const dd = e.parameter['deliveryDate' + i] || (i === 1 ? deliveryDate : '');
    const dnt = e.parameter['deliveryNoteText' + i] || '';
    const disabledAttr = (existingDateCount === 1) ? 'disabled' : '';

    html += `    <div class="shipping-date-row d-flex align-items-center gap-2 mb-2" data-row="${i}">`;
    html += `      <span class="badge bg-secondary">#${i}</span>`;
    html += `      <label class="col-form-label">発送日</label>`;
    html += `      <input type="date" class="form-control" style="width:160px;" name="shippingDate${i}" required value="${sd}">`;
    html += `      <label class="col-form-label">納品日</label>`;
    html += `      <input type="date" class="form-control" style="width:160px;" name="deliveryDate${i}" required value="${dd}">`;
    html += `      <label class="col-form-label">納品書テキスト</label>`;
    html += `      <input type="text" class="form-control" style="width:300px;" name="deliveryNoteText${i}" placeholder="例: 月曜コース分の商品" value="${dnt}">`;
    html += `      <button type="button" class="btn btn-sm btn-outline-danger" onclick="removeShippingDate(this)" ${disabledAttr}>✕</button>`;
    html += `    </div>`;
  }

  html += `  </div>`;
  html += `</div>`;

  // 受付方法
  html += `<div class="mb-2">`;
  html += `<label for="receiptWay" class="text-left form-label">受付方法</label>`;
  html += `<select class="form-select" id="receiptWay" name="receiptWay" required>`;
  for (const receipt of receipts) {
    const receiptWay = receipt['受付方法'];
    if (receiptWay == e.parameter.receiptWay) {
      html += `<option value="${receiptWay}" selected>${receiptWay}</option>`;
    } else {
      html += `<option value="${receiptWay}">${receiptWay}</option>`;
    }
  }
  html += `</select>`;
  html += `</div>`;

  // 受付者
  html += `<div class="mb-2">`;
  html += `<label for="recipient" class="text-left form-label">受付者</label>`;
  html += `<select class="form-select" id="recipient" name="recipient" required>`;
  for (const recipient of recipients) {
    const recip = recipient['名前'];
    if (recip == e.parameter.recipient) {
      html += `<option value="${recip}" selected>${recip}</option>`;
    } else {
      html += `<option value="${recip}">${recip}</option>`;
    }
  }
  html += `</select>`;
  html += `</div>`;

  // 納品方法・配達時間帯
  html += `<div class="mt-2 mb-2 row g-3 align-items-center">`;
  html += `  <div class="col-auto">`;
  html += `    <label for="deliveryMethod" class="col-form-label">納品方法</label>`;
  html += `  </div>`;
  html += `  <div class="col-auto">`;
  html += `    <select class="form-select" id="deliveryMethod" name="deliveryMethod" required onchange=deliveryMethodChange() >`;
  html += `      <option value=""></option>`;
  for (const deliveryMethod of deliveryMethods) {
    const deliMethod = deliveryMethod['納品方法'];
    if (deliMethod == e.parameter.deliveryMethod) {
      html += `<option value="${deliMethod}" selected>${deliMethod}</option>`;
    } else {
      html += `<option value="${deliMethod}">${deliMethod}</option>`;
    }
  }
  html += `    </select>`;
  html += `  </div>`;
  html += `  <div class="col-auto">`;
  html += `    <label for="deliveryTime" class="col-form-label">配達時間帯</label>`;
  html += `  </div>`;
  html += `  <div class="col-auto">`;
  html += `    <select class="form-select" id="deliveryTime" name="deliveryTime" >`;
  html += `      <option value=""></option>`;
  for (const deliveryTime of deliveryTimes) {
    const deliTime = deliveryTime['時間指定'];
    const deliTimeVal = deliveryTime['時間指定値'];
    const deliveryMethod = deliveryTime['納品方法'];
    if (deliTimeVal == e.parameter.deliveryTime) {
      html += `<option value="${deliTimeVal}" data-val="${deliveryMethod}" selected>${deliTime}</option>`;
    } else {
      html += `<option value="${deliTimeVal}" data-val="${deliveryMethod}">${deliTime}</option>`;
    }
  }
  html += `    </select>`;
  html += `  </div>`;
  html += `</div>`;

  // チェックリスト
  html += `<div class="checkbox-group">`;
  const checklistItems = [
    { id: 'deliveryChk', value: '納品書', label: '納品書' },
    { id: 'billChk', value: '請求書', label: '請求書' },
    { id: 'receiptChk', value: '領収書', label: '領収書' },
    { id: 'pamphletChk', value: 'パンフ', label: 'パンフ' },
    { id: 'recipeChk', value: 'レシピ', label: 'レシピ' }
  ];
  for (const item of checklistItems) {
    const checked = e.parameters.checklist && e.parameters.checklist.includes(item.value) ? 'checked' : '';
    html += `<div class="form-check">`;
    html += `  <input class="form-check-input" type="checkbox" value="${item.value}" id="${item.id}" name="checklist" ${checked}>`;
    html += `  <label class="form-check-label" for="${item.id}">${item.label}</label>`;
    html += `</div>`;
  }
  html += `</div>`;

  // その他添付
  html += `<div class="mb-2">`;
  html += `  <label for="otherAttach" class="col-form-label">その他添付</label>`;
  html += `  <input type="text" class="form-control" id="otherAttach" name="otherAttach" value="${e.parameter.otherAttach ? e.parameter.otherAttach : ""}" >`;
  html += `</div>`;
  html += `</div>`; // section-body閉じ

  // ============================================
  // 商品情報セクション
  // ============================================
  html += `
<div class="confirm-section section-products"><div class="section-header">
  <div class="section-header-row">
    <span class="section-header-label">🛒 商品情報</span>
  </div>
</div>
<div class="section-body">
`;

  html += `
    <table class="table text-center">
      <thead>
        <tr>
          <th scope="col">商品分類</th>
          <th scope="col">商品名</th>
          <th scope="col">価格</th>
          <th scope="col">個数</th>
        </tr>
      </thead>
      <tbody>
  `;

  const categorys = getAllRecords('商品分類');
  var rowNum = 0;
  for (let i = 0; i < 10; i++) {
    rowNum++;
    html += `<tr>`;
    html += `<td>`;
    var bunrui = "bunrui" + rowNum;
    html += `<select class="form-select" id="${bunrui}" name="${bunrui}" onchange=bunruiChange(${rowNum}) >`;
    html += `<option value=""></option>`;
    for (const category of categorys) {
      if (e.parameter[bunrui] == category['商品分類']) {
        html += `<option value="${category['商品分類']}" selected>${category['商品分類']}</option>`;
      } else {
        html += `<option value="${category['商品分類']}" >${category['商品分類']}</option>`;
      }
    }
    html += `</select>`;
    html += `</td>`;
    var product = "product" + rowNum;
    html += `<td>`;
    html += `<select class="form-select" id="${product}" name="${product}" onchange=productChange(${rowNum}) >`;
    html += `<option value="" data-val="" data-zaiko=""></option>`;
    for (const item of items) {
      if (item['在庫数'] > 0) {
        if (e.parameter[product] == item['商品名']) {
          html += `<option value="${item['商品名']}" data-val="${item['商品分類']}" data-name="${item['商品名']}" data-abbreviation="${item['送り状品名']}" data-zaiko="${item['在庫数']}" 
          data-price="${item['価格（P)']}" selected>${item['商品名']}</option>`;
        } else {
          html += `<option value="${item['商品名']}" data-val="${item['商品分類']}" data-name="${item['商品名']}" data-abbreviation="${item['送り状品名']}" data-zaiko="${item['在庫数']}"
          data-price="${item['価格（P)']}" >${item['商品名']}</option>`;
        }
      }
    }
    html += `</select>`;
    html += `</td>`;
    var price = "price" + rowNum;
    html += `<td>`;
    html += `<input type="number" class="form-control no-spin" id="${price}" name="${price}" min='0'  value="${e.parameter[price] ? e.parameter[price] : ""}" >`;
    html += `</td>`;
    var quantity = "quantity" + rowNum;
    html += `<td>`;
    html += `<div class="d-flex align-items-center gap-1">`;
    html += `<input type="number" class="form-control no-spin flex-grow-1" id="${quantity}" name="${quantity}" min='0' max='999' step="0.1" title="整数部3桁小数部1桁の数字のみを入力してください。" value="${e.parameter[quantity] ? e.parameter[quantity] : ""}" >`;
    html += `<button type="button" class="btn btn-sm btn-outline-danger px-1 py-1" style="min-width: 20px; width: 20px; font-size: 0.7rem; line-height: 1;" onclick="clearProductRow(${rowNum})" title="この行をクリア">✖</button>`;
    html += `</div>`;
    html += `</td>`;
    html += `</tr>`;
  }
  html += `</tbody>`;
  html += `</table>`;
  html += `</div>`; // section-body閉じ

  // ============================================
  // 発送情報セクション
  // ============================================
  html += `
<div class="confirm-section section-shipping-info"><div class="section-header">
  <div class="section-header-row">
    <span class="section-header-label">🚚 発送情報</span>
  </div>
</div>
<div class="section-body">
`;

  html += `<div class="mb-2">
    <label for="sendProduct" class="text-left form-label">品名</label>
    <input type="text" class="form-control" id="sendProduct" name="sendProduct" value="${e.parameter.sendProduct ? e.parameter.sendProduct : ""}">
  </div>`;

  // 送り状種別・クール区分
  html += `<div class="mt-2 mb-2 row g-3 align-items-center">`;
  html += `  <div class="col-auto">`;
  html += `    <label for="invoiceType" class="col-form-label">送り状種別</label>`;
  html += `  </div>`;
  html += `  <div class="col-auto">`;
  html += `<select class="form-select" id="invoiceType" name="invoiceType" >`;
  html += `<option value=""></option>`;
  for (const invoiceType of invoiceTypes) {
    if (e.parameter.invoiceType == invoiceType['種別値']) {
      html += `<option value="${invoiceType['種別値']}" data-val="${invoiceType['納品方法']}" selected>${invoiceType['種別']}</option>`;
    } else {
      html += `<option value="${invoiceType['種別値']}" data-val="${invoiceType['納品方法']}" >${invoiceType['種別']}</option>`;
    }
  }
  html += `</select>`;
  html += `  </div>`;
  html += `  <div class="col-auto">`;
  html += `    <label for="coolCls" class="col-form-label">クール区分</label>`;
  html += `  </div>`;
  html += `  <div class="col-auto">`;
  html += `<select class="form-select" id="coolCls" name="coolCls" >`;
  html += `<option value=""></option>`;
  for (const coolCls of coolClss) {
    if (e.parameter.coolCls == coolCls['種別値']) {
      html += `<option value="${coolCls['種別値']}" data-val="${coolCls['納品方法']}" selected>${coolCls['種別']}</option>`;
    } else {
      html += `<option value="${coolCls['種別値']}" data-val="${coolCls['納品方法']}" >${coolCls['種別']}</option>`;
    }
  }
  html += `</select>`;
  html += `  </div>`;
  html += `</div>`;

  // 荷扱い１・２
  html += `<div class="mt-2 mb-2 row g-3 align-items-center">`;
  html += `  <div class="col-auto">`;
  html += `    <label for="cargo1" class="col-form-label">荷扱い１</label>`;
  html += `  </div>`;
  html += `  <div class="col-auto">`;
  html += `<select class="form-select" id="cargo1" name="cargo1" >`;
  html += `<option value=""></option>`;
  for (const cargo of cargos) {
    if (e.parameter.cargo1 == cargo['種別値']) {
      html += `<option value="${cargo['種別値']}" data-val="${cargo['納品方法']}" selected>${cargo['種別']}</option>`;
    } else {
      html += `<option value="${cargo['種別値']}" data-val="${cargo['納品方法']}" >${cargo['種別']}</option>`;
    }
  }
  html += `</select>`;
  html += `  </div>`;
  html += `  <div class="col-auto">`;
  html += `    <label for="cargo2" class="col-form-label">荷扱い２</label>`;
  html += `  </div>`;
  html += `  <div class="col-auto">`;
  html += `<select class="form-select" id="cargo2" name="cargo2" >`;
  html += `<option value=""></option>`;
  for (const cargo of cargos) {
    if (e.parameter.cargo2 == cargo['種別値']) {
      html += `<option value="${cargo['種別値']}" data-val="${cargo['納品方法']}" selected>${cargo['種別']}</option>`;
    } else {
      html += `<option value="${cargo['種別値']}" data-val="${cargo['納品方法']}" >${cargo['種別']}</option>`;
    }
  }
  html += `</select>`;
  html += `  </div>`;
  html += `</div>`;

  // 荷扱い３（佐川の場合のみ表示）
  html += `<div class="mt-2 mb-2 row g-3 align-items-center" id="cargo3Container" style="display:none;">`;
  html += `  <div class="col-auto">`;
  html += `    <label for="cargo3" class="col-form-label">荷扱い３</label>`;
  html += `  </div>`;
  html += `  <div class="col-auto">`;
  html += `<select class="form-select" id="cargo3" name="cargo3" >`;
  html += `<option value=""></option>`;
  for (const cargo of cargos) {
    if (e.parameter.cargo3 == cargo['種別値']) {
      html += `<option value="${cargo['種別値']}" data-val="${cargo['納品方法']}" selected>${cargo['種別']}</option>`;
    } else {
      html += `<option value="${cargo['種別値']}" data-val="${cargo['納品方法']}" >${cargo['種別']}</option>`;
    }
  }
  html += `</select>`;
  html += `  </div>`;
  html += `</div>`;

  // 代引総額・代引内税
  html += `<div class="mt-2 mb-2 row g-3 align-items-center">`;
  html += `  <div class="col-auto">`;
  html += `    <label for="cashOnDelivery" class="col-form-label">代引総額</label>`;
  html += `  </div>`;
  html += `  <div class="col-auto">`;
  html += `    <input type="number" class="form-control" id="cashOnDelivery" name="cashOnDelivery" min='0'  value="${e.parameter.cashOnDelivery ? e.parameter.cashOnDelivery : ""}">`;
  html += `  </div>`;
  html += `  <div class="col-auto">`;
  html += `    <label for="cashOnDeliTax" class="col-form-label">代引内税</label>`;
  html += `  </div>`;
  html += `  <div class="col-auto">`;
  html += `    <input type="number" class="form-control" id="cashOnDeliTax" name="cashOnDeliTax" min='0'  value="${e.parameter.cashOnDeliTax ? e.parameter.cashOnDeliTax : ""}">`;
  html += `  </div>`;
  html += `</div>`;

  // 発行枚数
  html += `<div class="mt-2 mb-2 row g-3 align-items-center">`;
  html += `  <div class="col-auto">`;
  html += `    <label for="copiePrint" class="col-form-label">発行枚数</label>`;
  html += `  </div>`;
  html += `  <div class="col-auto">`;
  html += `    <input type="number" class="form-control" id="copiePrint" name="copiePrint" min='0'  value="${e.parameter.copiePrint ? e.parameter.copiePrint : ""}">`;
  html += `  </div>`;
  html += `</div>`;

  // 備考欄
  html += `<div class="mb-3">`;
  html += `<label for="internalMemo" class="text-left form-label">社内メモ</label>`;
  html += `<textarea class="form-control" id="internalMemo" name="internalMemo" rows="2" cols="30" placeholder="freeeの社内メモ列に出力されます">${e.parameter.internalMemo ? e.parameter.internalMemo : ""}</textarea>`;
  html += `</div>`;
  html += `<div class="mb-3">`;
  html += `<label for="csvmemo" class="text-left form-label">送り状　備考欄</label>`;
  html += `<textarea class="form-control" id="csvmemo" name="csvmemo" rows="2" cols="30" maxlength="22" placeholder="freeeの備考列に出力されます">${e.parameter.csvmemo ? e.parameter.csvmemo : ""}</textarea>`;
  html += `</div>`;
  html += `<div class="mb-3">`;
  html += `<label for="deliveryMemo" class="text-left form-label">納品書　備考欄</label>`;
  html += `<textarea class="form-control" id="deliveryMemo" name="deliveryMemo" rows="3" cols="30" maxlength="90">${e.parameter.deliveryMemo ? e.parameter.deliveryMemo : ""}</textarea>`;
  html += `</div>`;
  html += `<div class="mb-3">`;
  html += `<label for="memo" class="text-left form-label">メモ</label>`;
  html += `<textarea class="form-control" id="memo" name="memo" rows="3" cols="30">${e.parameter.memo ? e.parameter.memo : ""}</textarea>`;
  html += `</div>`;
  html += `</div>`; // section-body閉じ

  // 編集モードの場合のみhiddenフィールドを追加
  if (editOrderId && !isInheritMode) {
    html += `<input type="hidden" name="editOrderId" value="${editOrderId}">`;
    html += `<input type="hidden" name="editMode" value="true">`;
  }

  return html;
}

/**
 * 受注確認画面（shippingComfirm.html）のメイン生成関数
 * 
 * 受注入力画面から送信されたパラメータを元に、確認用の読み取り専用画面を構築します。
 * 過去注文との差異チェック（AIによる入力ミス防止）もここで行います。
 * 
 * @param {Object} e - POSTリクエストイベントオブジェクト
 * @returns {string} 受注確認画面のHTML文字列
 * 
 * @see checkOrderAgainstHistory - 過去注文との比較ロジック
 */
function getShippingComfirmHTML(e) {
  const master = getMasterDataCached();
  const recipients = master.recipients;
  const deliveryMethods = master.deliveryMethods;
  const receipts = master.receipts;
  const deliveryTimes = master.deliveryTimes;
  const invoiceTypes = master.invoiceTypes;
  const coolClss = master.coolClss;
  const cargos = master.cargos;

  var nDate = new Date();
  var strDate = Utilities.formatDate(nDate, 'JST', 'yyyy-MM-dd')
  const orderDate = e.parameter.orderDate ? e.parameter.orderDate : strDate;
  var n = 2;
  nDate.setDate(nDate.getDate() + n);
  strDate = Utilities.formatDate(nDate, 'JST', 'yyyy-MM-dd')
  const shippingDate = e.parameter.shippingDate ? e.parameter.shippingDate : strDate;
  nDate.setDate(nDate.getDate() + 1);
  strDate = Utilities.formatDate(nDate, 'JST', 'yyyy-MM-dd')
  const deliveryDate = e.parameter.deliveryDate ? e.parameter.deliveryDate : strDate;

  // ============================================
  // CSS スタイル（確認画面用 - デザインシステム統一）
  // ============================================
  let html = `
<style>
/* セクションカード（.card-unified ベース） */
.confirm-section {
  background: var(--bg-card);
  border: 1px solid var(--border-color);
  border-radius: var(--radius-md);
  margin-bottom: var(--space-4);
  box-shadow: var(--shadow);
  overflow: hidden;
}

.section-header {
  color: white;
  padding: var(--space-3) var(--space-4);
  font-weight: 600;
}
.section-header-label {
  font-size: 1rem;
  display: flex;
  align-items: center;
  gap: var(--space-2);
}

/* 各セクションのカラー（グラデーション統一） */
.section-shipping-to .section-header {
  background: linear-gradient(135deg, #b8860b 0%, #daa520 100%);
}
.section-customer .section-header {
  background: var(--gradient-purple);
}
.section-shipping-from .section-header {
  background: linear-gradient(135deg, #8b4513 0%, #a0522d 100%);
}
.section-order-basic .section-header {
  background: linear-gradient(135deg, #008b8b 0%, #20b2aa 100%);
}
.section-products .section-header {
  background: linear-gradient(135deg, #228b22 0%, #32cd32 100%);
}
.section-shipping-info .section-header {
  background: linear-gradient(135deg, #1e3a5f 0%, #4169e1 100%);
}

/* セクション本体 */
.section-body {
  background: var(--bg-card);
  padding: var(--space-4);
}

/* 確認画面用：読み取り専用フィールド */
.confirm-field {
  background: var(--bg-body);
  border: 1px solid var(--border-color);
  border-radius: var(--radius);
  padding: var(--space-2) var(--space-3);
  margin-bottom: var(--space-2);
}
.confirm-field-label {
  font-size: 0.75rem;
  color: var(--text-secondary);
  margin-bottom: var(--space-1);
}
.confirm-field-value {
  font-size: 0.95rem;
  color: var(--text-primary);
  font-weight: 500;
}

/* 確認画面用：インライン表示 */
.confirm-row {
  display: flex;
  flex-wrap: wrap;
  gap: var(--space-4);
  margin-bottom: var(--space-2);
}
.confirm-item {
  display: flex;
  align-items: center;
  gap: var(--space-2);
}
.confirm-item-label {
  font-size: 0.85rem;
  color: var(--text-secondary);
}
.confirm-item-value {
  font-size: 0.95rem;
  color: var(--text-primary);
  font-weight: 500;
}

/* チェックリストバッジ */
.checklist-badges {
  display: flex;
  flex-wrap: wrap;
  gap: var(--space-2);
  margin: var(--space-2) 0;
}
.checklist-badge {
  padding: var(--space-1) var(--space-2);
  border-radius: var(--radius-full);
  font-size: 0.8rem;
  font-weight: 500;
}
.checklist-badge.checked {
  background: var(--status-success-light);
  color: #155724;
  border: 1px solid var(--status-success);
}
.checklist-badge.unchecked {
  background: var(--bg-body);
  color: var(--text-tertiary);
  border: 1px solid var(--border-color);
  text-decoration: line-through;
}

/* 日程カード */
.date-card {
  background: var(--gradient-purple);
  color: white;
  border-radius: var(--radius);
  padding: var(--space-3) var(--space-4);
  margin-bottom: var(--space-2);
  display: flex;
  align-items: center;
  gap: 16px;
}
.date-card .badge {
  background: rgba(255,255,255,0.2);
  padding: 4px 10px;
  border-radius: 20px;
  font-size: 0.8rem;
}
.date-card .dates {
  display: flex;
  gap: 20px;
}
.date-card .date-item {
  display: flex;
  align-items: center;
  gap: 6px;
}
.date-card .date-label {
  font-size: 0.8rem;
  opacity: 0.9;
}
.date-card .date-value {
  font-weight: 600;
  font-size: 1rem;
}

/* 商品テーブル（PC） */
.product-table-pc { display: none; }
.product-cards-sp { display: block; }

@media (min-width: 768px) {
  .product-table-pc { display: block; }
  .product-cards-sp { display: none; }
}

/* 商品カード（スマホ） */
.product-card {
  background: #f8f9fa;
  border: 1px solid #dee2e6;
  border-radius: 8px;
  padding: 12px;
  margin-bottom: 10px;
}
.product-card-header {
  display: flex;
  justify-content: space-between;
  align-items: center;
  margin-bottom: 8px;
}
.product-card-header .category {
  background: #6c757d;
  color: white;
  padding: 2px 8px;
  border-radius: 4px;
  font-size: 0.75rem;
}
.product-card-header .product-name {
  font-weight: 600;
  font-size: 1rem;
}
.product-card-body {
  display: flex;
  justify-content: space-between;
  align-items: center;
}
.product-card-calc {
  color: #666;
  font-size: 0.9rem;
}
.product-card-total {
  font-weight: 600;
  font-size: 1.1rem;
  color: #28a745;
}

/* 合計セクション */
.total-section {
  background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
  color: white;
  padding: 16px;
  border-radius: 8px;
  margin-top: 16px;
  margin-bottom: 16px;
}
.total-section .label {
  font-size: 1rem;
}
.total-section .amount {
  font-size: 1.5rem;
  font-weight: 700;
}

/* レスポンシブ */
@media (max-width: 768px) {
  .section-header {
    padding: 6px 8px;
  }
  .section-header-label {
    font-size: 0.9rem;
  }
  .section-body {
    padding: 10px;
  }
  .confirm-row {
    flex-direction: column;
    gap: 8px;
  }
  .date-card {
    flex-direction: column;
    align-items: flex-start;
    gap: 8px;
  }
  .date-card .dates {
    flex-direction: column;
    gap: 4px;
  }
}
</style>
`;

  // ============================================
  // 発送先情報セクション
  // ============================================
  html += `
<div class="confirm-section section-shipping-to">
  <div class="section-header">
    <span class="section-header-label">📦 発送先情報</span>
  </div>
  <div class="section-body">
`;
  html += `<div class="confirm-field">
    <div class="confirm-field-label">発送先名</div>
    <div class="confirm-field-value">${e.parameter.shippingToName || '-'}</div>
  </div>`;
  html += `<div class="confirm-row">
    <div class="confirm-item">
      <span class="confirm-item-label">〒</span>
      <span class="confirm-item-value">${e.parameter.shippingToZipcode || '-'}</span>
    </div>
    <div class="confirm-item">
      <span class="confirm-item-label">TEL</span>
      <span class="confirm-item-value">${e.parameter.shippingToTel || '-'}</span>
    </div>
  </div>`;
  html += `<div class="confirm-field">
    <div class="confirm-field-label">住所</div>
    <div class="confirm-field-value">${e.parameter.shippingToAddress || '-'}</div>
  </div>`;
  // Hidden inputs
  html += `<input type="hidden" name="shippingToName" value="${e.parameter.shippingToName || ''}">`;
  html += `<input type="hidden" name="shippingToZipcode" value="${e.parameter.shippingToZipcode || ''}">`;
  html += `<input type="hidden" name="shippingToAddress" value="${e.parameter.shippingToAddress || ''}">`;
  html += `<input type="hidden" name="shippingToTel" value="${e.parameter.shippingToTel || ''}">`;
  html += `</div>`;

  // ============================================
  // 顧客情報セクション
  // ============================================
  html += `
<div class="confirm-section section-customer"><div class="section-header">
  <div class="section-header-row">
    <span class="section-header-label">👤 顧客情報</span>
  </div>
</div>
<div class="section-body">
`;
  html += `<div class="confirm-field">
    <div class="confirm-field-label">顧客名</div>
    <div class="confirm-field-value">${e.parameter.customerName || '-'}</div>
  </div>`;
  html += `<div class="confirm-row">
    <div class="confirm-item">
      <span class="confirm-item-label">〒</span>
      <span class="confirm-item-value">${e.parameter.customerZipcode || '-'}</span>
    </div>
    <div class="confirm-item">
      <span class="confirm-item-label">TEL</span>
      <span class="confirm-item-value">${e.parameter.customerTel || '-'}</span>
    </div>
  </div>`;
  html += `<div class="confirm-field">
    <div class="confirm-field-label">住所</div>
    <div class="confirm-field-value">${e.parameter.customerAddress || '-'}</div>
  </div>`;
  // Hidden inputs
  html += `<input type="hidden" name="customerName" value="${e.parameter.customerName || ''}">`;
  html += `<input type="hidden" name="customerZipcode" value="${e.parameter.customerZipcode || ''}">`;
  html += `<input type="hidden" name="customerAddress" value="${e.parameter.customerAddress || ''}">`;
  html += `<input type="hidden" name="customerTel" value="${e.parameter.customerTel || ''}">`;
  html += `</div>`;

  // ============================================
  // 発送元情報セクション
  // ============================================
  html += `
<div class="confirm-section section-shipping-from"><div class="section-header">
  <div class="section-header-row">
    <span class="section-header-label">🏭 発送元情報</span>
  </div>
</div>
<div class="section-body">
`;
  html += `<div class="confirm-field">
    <div class="confirm-field-label">発送元名</div>
    <div class="confirm-field-value">${e.parameter.shippingFromName || '-'}</div>
  </div>`;
  html += `<div class="confirm-row">
    <div class="confirm-item">
      <span class="confirm-item-label">〒</span>
      <span class="confirm-item-value">${e.parameter.shippingFromZipcode || '-'}</span>
    </div>
    <div class="confirm-item">
      <span class="confirm-item-label">TEL</span>
      <span class="confirm-item-value">${e.parameter.shippingFromTel || '-'}</span>
    </div>
  </div>`;
  html += `<div class="confirm-field">
    <div class="confirm-field-label">住所</div>
    <div class="confirm-field-value">${e.parameter.shippingFromAddress || '-'}</div>
  </div>`;
  // Hidden inputs
  html += `<input type="hidden" name="shippingFromName" value="${e.parameter.shippingFromName || ''}">`;
  html += `<input type="hidden" name="shippingFromZipcode" value="${e.parameter.shippingFromZipcode || ''}">`;
  html += `<input type="hidden" name="shippingFromAddress" value="${e.parameter.shippingFromAddress || ''}">`;
  html += `<input type="hidden" name="shippingFromTel" value="${e.parameter.shippingFromTel || ''}">`;
  html += `</div>`;

  // ============================================
  // 受注基本情報セクション
  // ============================================
  html += `
<div class="confirm-section section-order-basic"><div class="section-header">
  <div class="section-header-row">
    <span class="section-header-label">📝 受注基本情報</span>
  </div>
</div>
<div class="section-body">
`;

  // 日程数をカウント
  let dateCount = 0;
  for (let i = 1; i <= 10; i++) {
    if (e.parameter['shippingDate' + i]) {
      dateCount = i;
    } else {
      break;
    }
  }

  // 日程カード表示
  html += `<div class="mb-3">`;
  html += `<div class="fw-bold mb-2">📅 発送日程（${dateCount}件）</div>`;

  for (let i = 1; i <= dateCount; i++) {
    const sd = e.parameter['shippingDate' + i] || '';
    const dd = e.parameter['deliveryDate' + i] || '';
    const dnt = e.parameter['deliveryNoteText' + i] || '';
    html += `<div class="date-card">`;
    html += `  <span class="badge">#${i}</span>`;
    html += `  <div class="dates">`;
    html += `    <div class="date-item">`;
    html += `      <span class="date-label">発送日</span>`;
    html += `      <span class="date-value">${sd}</span>`;
    html += `    </div>`;
    html += `    <div class="date-item">`;
    html += `      <span class="date-label">→ 納品日</span>`;
    html += `      <span class="date-value">${dd}</span>`;
    html += `    </div>`;
    if (dnt) {
      html += `    <div class="date-item">`;
      html += `      <span class="date-label">📄</span>`;
      html += `      <span class="date-value">${dnt}</span>`;
      html += `    </div>`;
    }
    html += `  </div>`;
    html += `</div>`;
    html += `<input type="hidden" name="shippingDate${i}" value="${sd}">`;
    html += `<input type="hidden" name="deliveryDate${i}" value="${dd}">`;
    html += `<input type="hidden" name="deliveryNoteText${i}" value="${dnt}">`;
  }
  html += `</div>`;

  // 受付方法・受付者
  html += `<div class="confirm-row">`;
  html += `  <div class="confirm-item">`;
  html += `    <span class="confirm-item-label">受付方法</span>`;
  html += `    <span class="confirm-item-value">${e.parameter.receiptWay || '-'}</span>`;
  html += `  </div>`;
  html += `  <div class="confirm-item">`;
  html += `    <span class="confirm-item-label">受付者</span>`;
  html += `    <span class="confirm-item-value">${e.parameter.recipient || '-'}</span>`;
  html += `  </div>`;
  html += `</div>`;
  html += `<input type="hidden" name="receiptWay" value="${e.parameter.receiptWay || ''}">`;
  html += `<input type="hidden" name="recipient" value="${e.parameter.recipient || ''}">`;

  // 納品方法・配達時間帯
  // 配達時間帯の表示名を取得
  let deliveryTimeDisplay = '-';
  for (const dt of deliveryTimes) {
    if (dt['時間指定値'] == e.parameter.deliveryTime) {
      deliveryTimeDisplay = dt['時間指定'];
      break;
    }
  }
  html += `<div class="confirm-row">`;
  html += `  <div class="confirm-item">`;
  html += `    <span class="confirm-item-label">納品方法</span>`;
  html += `    <span class="confirm-item-value">${e.parameter.deliveryMethod || '-'}</span>`;
  html += `  </div>`;
  html += `  <div class="confirm-item">`;
  html += `    <span class="confirm-item-label">配達時間帯</span>`;
  html += `    <span class="confirm-item-value">${deliveryTimeDisplay}</span>`;
  html += `  </div>`;
  html += `</div>`;
  html += `<input type="hidden" name="deliveryMethod" value="${e.parameter.deliveryMethod || ''}">`;
  html += `<input type="hidden" name="deliveryTime" value="${e.parameter.deliveryTime || ''}">`;

  // チェックリストバッジ
  const checklistItems = ['納品書', '請求書', '領収書', 'パンフ', 'レシピ'];
  html += `<div class="checklist-badges">`;
  for (const item of checklistItems) {
    const isChecked = e.parameters.checklist && e.parameters.checklist.includes(item);
    const badgeClass = isChecked ? 'checked' : 'unchecked';
    const icon = isChecked ? '✓' : '';
    html += `<span class="checklist-badge ${badgeClass}">${icon} ${item}</span>`;
    if (isChecked) {
      html += `<input type="hidden" name="checklist" value="${item}">`;
    }
  }
  html += `</div>`;

  // 定期便情報（チェックが入っている場合のみ表示）
  if (e.parameter.isRecurringOrder === 'true') {
    // 個別パラメータから間隔オブジェクトを構築
    var intervalObj;
    if (e.parameter.recurringInterval) {
      // 既にJSON形式で渡されている場合（修正画面からの遷移等）
      try {
        intervalObj = JSON.parse(e.parameter.recurringInterval);
      } catch (ex) {
        // 旧形式（数値のみ）
        intervalObj = { type: 'nmonth', value: Number(e.parameter.recurringInterval) || 1 };
      }
    } else {
      // shipping.htmlからの個別パラメータを統合
      var recType = e.parameter.recurringType || 'monthly';
      intervalObj = { type: recType };

      if (recType === 'weekly') {
        intervalObj.value = Number(e.parameter.recurringWeekday) || 1;
      } else if (recType === 'biweekly' || recType === 'triweekly' || recType.endsWith('week')) {
        // biweekly, triweekly は "week" で終わらないため明示的に含める。4week, 5week... は endsWith('week')
        var weekNum = recType === 'biweekly' ? 2 : recType === 'triweekly' ? 3 : parseInt(recType.replace('week', '')) || 2;
        intervalObj.type = 'nweek';
        intervalObj.value = weekNum;
        intervalObj.weekday = Number(e.parameter.recurringWeekday) || 1;
      } else if (recType === 'monthly') {
        // 毎月X日(1-31) / 毎月初日('first') / 毎月末日('last') をそのまま渡す
        var monthDay = e.parameter.recurringMonthDay;
        if (monthDay === 'first' || monthDay === 'last') {
          intervalObj.value = monthDay;
        } else {
          intervalObj.value = Number(monthDay) || 1;
        }
      } else if (recType === '2month' || recType === '3month') {
        intervalObj.type = 'nmonth';
        intervalObj.value = parseInt(recType.replace('month', '')) || 1;
      } else {
        intervalObj.value = 1;
      }
    }

    var intervalDisplay = formatIntervalForDisplay(intervalObj);
    var intervalJson = JSON.stringify(intervalObj);

    // 次回発送日・納品日を計算
    var nextShippingDateDisplay = '-';
    var nextDeliveryDateDisplay = '-';
    var baseShippingDateStr = e.parameter.shippingDate1;
    var baseDeliveryDateStr = e.parameter.deliveryDate1;

    if (baseShippingDateStr && baseDeliveryDateStr) {
      var baseShippingDate = new Date(baseShippingDateStr);
      var baseDeliveryDate = new Date(baseDeliveryDateStr);
      if (!isNaN(baseShippingDate.getTime()) && !isNaN(baseDeliveryDate.getTime())) {
        var nextShippingDate = calcNextShippingDateFlexible(baseShippingDate, intervalObj);
        var daysDiff = Math.round((baseDeliveryDate - baseShippingDate) / (1000 * 60 * 60 * 24));
        var nextDeliveryDate = new Date(nextShippingDate);
        nextDeliveryDate.setDate(nextDeliveryDate.getDate() + daysDiff);
        nextShippingDateDisplay = Utilities.formatDate(nextShippingDate, 'JST', 'yyyy/MM/dd');
        nextDeliveryDateDisplay = Utilities.formatDate(nextDeliveryDate, 'JST', 'yyyy/MM/dd');
      }
    }

    html += `<div class="mt-3 p-3" style="background: linear-gradient(135deg, #17a2b8 0%, #20c997 100%); border-radius: 8px;">`;
    html += `  <div style="color: white; font-weight: 600; display: flex; align-items: center; gap: 8px;">`;
    html += `    <span style="font-size: 1.2rem;">🔄</span>`;
    html += `    <span>定期便として登録</span>`;
    html += `    <span style="background: rgba(255,255,255,0.2); padding: 2px 10px; border-radius: 20px; font-size: 0.85rem;">${intervalDisplay}</span>`;
    html += `  </div>`;
    html += `  <div style="color: rgba(255,255,255,0.9); font-size: 0.85rem; margin-top: 8px; padding-left: 28px;">`;
    html += `    <div>次回発送日: ${nextShippingDateDisplay}</div>`;
    html += `    <div>次回納品日: ${nextDeliveryDateDisplay}</div>`;
    html += `  </div>`;
    html += `</div>`;
    html += `<input type="hidden" name="isRecurringOrder" value="true">`;
    html += `<input type="hidden" name="recurringInterval" value='${intervalJson}'>`;
  }

  // その他添付
  if (e.parameter.otherAttach) {
    html += `<div class="confirm-row">`;
    html += `  <div class="confirm-item">`;
    html += `    <span class="confirm-item-label">その他添付</span>`;
    html += `    <span class="confirm-item-value">${e.parameter.otherAttach}</span>`;
    html += `  </div>`;
    html += `</div>`;
  }
  html += `<input type="hidden" name="otherAttach" value="${e.parameter.otherAttach || ''}">`;
  html += `</div>`; // section-body閉じ

  // ============================================
  // 商品情報セクション
  // ============================================
  html += `
<div class="confirm-section section-products"><div class="section-header">
  <div class="section-header-row">
    <span class="section-header-label">🛒 商品情報</span>
  </div>
</div>
<div class="section-body">
`;

  html += `<p class="fw-bold text-center mb-3">以下の内容で受注していいですか？</p>`;

  // 商品データを収集
  let total = 0;
  const products = [];

  for (let i = 1; i <= 10; i++) {
    const bunruiVal = e.parameter['bunrui' + i];
    const productVal = e.parameter['product' + i];
    const count = Number(e.parameter['quantity' + i]);
    const unitPrice = Number(e.parameter['price' + i]);

    if (count > 0) {
      const subtotal = unitPrice * count;
      total += subtotal;
      products.push({
        row: i,
        bunrui: bunruiVal,
        product: productVal,
        price: unitPrice,
        quantity: count,
        subtotal: subtotal
      });
    }
  }

  // Hidden inputs（パラメータ保持用）
  products.forEach(p => {
    html += `<input type="hidden" name="bunrui${p.row}" value="${p.bunrui}">`;
    html += `<input type="hidden" name="product${p.row}" value="${p.product}">`;
    html += `<input type="hidden" name="price${p.row}" value="${p.price}">`;
    html += `<input type="hidden" name="quantity${p.row}" value="${p.quantity}">`;
  });

  // PC用テーブル表示
  html += `<div class="product-table-pc">`;
  html += `<table class="table table-striped">`;
  html += `<thead class="table-dark">`;
  html += `<tr>`;
  html += `<th class="text-start">分類</th>`;
  html += `<th class="text-start">商品名</th>`;
  html += `<th class="text-end">単価</th>`;
  html += `<th class="text-end">数量</th>`;
  html += `<th class="text-end">金額</th>`;
  html += `</tr>`;
  html += `</thead>`;
  html += `<tbody>`;

  products.forEach(p => {
    html += `<tr>`;
    html += `<td class="text-start">${p.bunrui}</td>`;
    html += `<td class="text-start">${p.product}</td>`;
    html += `<td class="text-end">¥${p.price.toLocaleString()}</td>`;
    html += `<td class="text-end">${p.quantity}</td>`;
    html += `<td class="text-end fw-bold">¥${p.subtotal.toLocaleString()}</td>`;
    html += `</tr>`;
  });

  html += `</tbody>`;
  html += `</table>`;
  html += `</div>`;

  // スマホ用カード表示
  html += `<div class="product-cards-sp">`;

  products.forEach(p => {
    html += `<div class="product-card">`;
    html += `  <div class="product-card-header">`;
    html += `    <span class="category">${p.bunrui}</span>`;
    html += `    <span class="product-name">${p.product}</span>`;
    html += `  </div>`;
    html += `  <div class="product-card-body">`;
    html += `    <span class="product-card-calc">¥${p.price.toLocaleString()} × ${p.quantity}個</span>`;
    html += `    <span class="product-card-total">¥${p.subtotal.toLocaleString()}</span>`;
    html += `  </div>`;
    html += `</div>`;
  });

  html += `</div>`;

  // ============================================
  // 過去注文比較チェック（電話対応時の入力ミス防止）
  // ============================================
  const shippingToName = e.parameter.shippingToName || '';
  if (shippingToName && products.length > 0) {
    // 現在の注文商品リストを作成
    const currentItems = products.map(p => ({
      productName: p.product,
      quantity: p.quantity
    }));

    try {
      const checkResult = JSON.parse(checkOrderAgainstHistory(shippingToName, currentItems));

      if (checkResult.hasWarnings && checkResult.warnings.length > 0) {
        html += `
<div class="alert alert-warning mt-3" role="alert">
  <div class="d-flex align-items-center mb-2">
    <span class="me-2" style="font-size: 1.5rem;">⚠️</span>
    <strong>過去注文との比較で確認が必要な点があります</strong>
  </div>
  <small class="text-muted d-block mb-2">（直近${checkResult.recentOrderCount}回の注文と比較）</small>
  <ul class="mb-0 ps-3">`;

        // 警告の種類ごとに整理して表示
        const quantityWarnings = checkResult.warnings.filter(w => w.type === 'quantity');
        const missingWarnings = checkResult.warnings.filter(w => w.type === 'missing');
        const newWarnings = checkResult.warnings.filter(w => w.type === 'new');

        // 数量異常（最も重要）
        quantityWarnings.forEach(w => {
          html += `<li class="text-danger"><strong>📊 数量確認:</strong> ${w.message}</li>`;
        });

        // 常連商品の欠落
        missingWarnings.forEach(w => {
          html += `<li class="text-warning"><strong>📦 商品漏れ:</strong> ${w.message}</li>`;
        });

        // 初めての商品（情報のみ）
        newWarnings.forEach(w => {
          html += `<li class="text-info"><strong>🆕 新規:</strong> ${w.message}</li>`;
        });

        html += `
  </ul>
  <div class="mt-2">
    <small class="text-muted">問題がなければそのまま「受注する」ボタンを押してください</small>
  </div>
</div>`;
      } else if (checkResult.isNewCustomer) {
        html += `
<div class="alert alert-info mt-3" role="alert">
  <span class="me-2">ℹ️</span>
  <strong>新規発送先</strong>: この発送先への過去の注文履歴はありません
</div>`;
      }
    } catch (error) {
      // エラーの場合は警告セクションを表示しない（サイレントに失敗）
      Logger.log('過去注文比較チェックエラー: ' + error.toString());
    }
  }

  // 合計セクション
  html += `<div class="total-section d-flex justify-content-between align-items-center">`;
  html += `  <span class="label">合計金額</span>`;
  html += `  <span class="amount">¥${total.toLocaleString()}</span>`;
  html += `</div>`;
  html += `</div>`; // section-body閉じ

  // ============================================
  // 発送情報セクション
  // ============================================
  html += `
<div class="confirm-section section-shipping-info"><div class="section-header">
  <div class="section-header-row">
    <span class="section-header-label">🚚 発送情報</span>
  </div>
</div>
<div class="section-body">
`;

  // 品名
  html += `<div class="confirm-field">
    <div class="confirm-field-label">品名</div>
    <div class="confirm-field-value">${e.parameter.sendProduct || '-'}</div>
  </div>`;
  html += `<input type="hidden" name="sendProduct" value="${e.parameter.sendProduct || ''}">`;

  // 送り状種別・クール区分の表示名取得
  let invoiceTypeDisplay = '-';
  for (const it of invoiceTypes) {
    if (it['種別値'] == e.parameter.invoiceType) {
      invoiceTypeDisplay = it['種別'];
      break;
    }
  }
  let coolClsDisplay = '-';
  for (const cc of coolClss) {
    if (cc['種別値'] == e.parameter.coolCls) {
      coolClsDisplay = cc['種別'];
      break;
    }
  }

  html += `<div class="confirm-row">`;
  html += `  <div class="confirm-item">`;
  html += `    <span class="confirm-item-label">送り状種別</span>`;
  html += `    <span class="confirm-item-value">${invoiceTypeDisplay}</span>`;
  html += `  </div>`;
  html += `  <div class="confirm-item">`;
  html += `    <span class="confirm-item-label">クール区分</span>`;
  html += `    <span class="confirm-item-value">${coolClsDisplay}</span>`;
  html += `  </div>`;
  html += `</div>`;
  html += `<input type="hidden" name="invoiceType" value="${e.parameter.invoiceType || ''}">`;
  html += `<input type="hidden" name="coolCls" value="${e.parameter.coolCls || ''}">`;

  // 荷扱い１・２・３の表示名取得
  let cargo1Display = '-';
  let cargo2Display = '-';
  let cargo3Display = '-';
  for (const c of cargos) {
    if (c['種別値'] == e.parameter.cargo1) {
      cargo1Display = c['種別'];
    }
    if (c['種別値'] == e.parameter.cargo2) {
      cargo2Display = c['種別'];
    }
    if (c['種別値'] == e.parameter.cargo3) {
      cargo3Display = c['種別'];
    }
  }

  html += `<div class="confirm-row">`;
  html += `  <div class="confirm-item">`;
  html += `    <span class="confirm-item-label">荷扱い１</span>`;
  html += `    <span class="confirm-item-value">${cargo1Display}</span>`;
  html += `  </div>`;
  html += `  <div class="confirm-item">`;
  html += `    <span class="confirm-item-label">荷扱い２</span>`;
  html += `    <span class="confirm-item-value">${cargo2Display}</span>`;
  html += `  </div>`;
  if (e.parameter.cargo3) {
    html += `  <div class="confirm-item">`;
    html += `    <span class="confirm-item-label">荷扱い３</span>`;
    html += `    <span class="confirm-item-value">${cargo3Display}</span>`;
    html += `  </div>`;
  }
  html += `</div>`;
  html += `<input type="hidden" name="cargo1" value="${e.parameter.cargo1 || ''}">`;
  html += `<input type="hidden" name="cargo2" value="${e.parameter.cargo2 || ''}">`;
  html += `<input type="hidden" name="cargo3" value="${e.parameter.cargo3 || ''}">`;

  // 代引総額・代引内税
  if (e.parameter.cashOnDelivery || e.parameter.cashOnDeliTax) {
    html += `<div class="confirm-row">`;
    html += `  <div class="confirm-item">`;
    html += `    <span class="confirm-item-label">代引総額</span>`;
    html += `    <span class="confirm-item-value">${e.parameter.cashOnDelivery ? '¥' + Number(e.parameter.cashOnDelivery).toLocaleString() : '-'}</span>`;
    html += `  </div>`;
    html += `  <div class="confirm-item">`;
    html += `    <span class="confirm-item-label">代引内税</span>`;
    html += `    <span class="confirm-item-value">${e.parameter.cashOnDeliTax ? '¥' + Number(e.parameter.cashOnDeliTax).toLocaleString() : '-'}</span>`;
    html += `  </div>`;
    html += `</div>`;
  }
  html += `<input type="hidden" name="cashOnDelivery" value="${e.parameter.cashOnDelivery || ''}">`;
  html += `<input type="hidden" name="cashOnDeliTax" value="${e.parameter.cashOnDeliTax || ''}">`;

  // 発行枚数
  if (e.parameter.copiePrint) {
    html += `<div class="confirm-row">`;
    html += `  <div class="confirm-item">`;
    html += `    <span class="confirm-item-label">発行枚数</span>`;
    html += `    <span class="confirm-item-value">${e.parameter.copiePrint}枚</span>`;
    html += `  </div>`;
    html += `</div>`;
  }
  html += `<input type="hidden" name="copiePrint" value="${e.parameter.copiePrint || ''}">`;

  // 社内メモ
  if (e.parameter.internalMemo) {
    html += `<div class="confirm-field">
      <div class="confirm-field-label">社内メモ</div>
      <div class="confirm-field-value">${e.parameter.internalMemo}</div>
    </div>`;
  }
  html += `<input type="hidden" name="internalMemo" value="${e.parameter.internalMemo || ''}">`;

  // 備考欄
  if (e.parameter.csvmemo) {
    html += `<div class="confirm-field">
      <div class="confirm-field-label">送り状 備考欄</div>
      <div class="confirm-field-value">${e.parameter.csvmemo}</div>
    </div>`;
  }
  html += `<input type="hidden" name="csvmemo" value="${e.parameter.csvmemo || ''}">`;

  if (e.parameter.deliveryMemo) {
    html += `<div class="confirm-field">
      <div class="confirm-field-label">納品書 備考欄</div>
      <div class="confirm-field-value">${e.parameter.deliveryMemo}</div>
    </div>`;
  }
  html += `<input type="hidden" name="deliveryMemo" value="${e.parameter.deliveryMemo || ''}">`;

  if (e.parameter.memo) {
    html += `<div class="confirm-field">
      <div class="confirm-field-label">メモ</div>
      <div class="confirm-field-value">${e.parameter.memo}</div>
    </div>`;
  }
  html += `<input type="hidden" name="memo" value="${e.parameter.memo || ''}">`;

  html += `</div>`; // section-body閉じ

  // 編集モードの場合
  const editOrderIdConfirm = e.parameter.editOrderId || '';
  if (editOrderIdConfirm) {
    html += `<input type="hidden" name="editOrderId" value="${editOrderIdConfirm}">`;
    html += `<input type="hidden" name="editMode" value="true">`;
  }

  return html;
}

/**
 * 一意のランダムなIDを生成
 * 
 * @param {number} length - IDの長さ (デフォルト: 8)
 * @returns {string} 生成されたID
 */
function generateId(length = 8) {
  const [alphabets, numbers] = ['abcdefghijklmnopqrstuvwxyz', '0123456789'];
  const string = alphabets + numbers;
  let id = alphabets.charAt(Math.floor(Math.random() * alphabets.length));
  for (let i = 0; i < length - 1; i++) {
    id += string.charAt(Math.floor(Math.random() * string.length));
  }
  return id;
}

/**
 * 電話注文モードなどで入力された新規顧客情報をマスタシートに登録
 * 
 * 重複チェックを行い、存在しない場合のみ「顧客情報」および「発送先情報」シートに追加します。
 * 
 * @param {Object} e - POSTリクエストイベントオブジェクト
 */
function registerNewCustomerToMaster(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // 顧客情報シート
    const customerSheet = ss.getSheetByName('顧客情報');
    if (!customerSheet) {
      Logger.log('顧客情報シートが見つかりません');
      return;
    }
    const customerLastRow = customerSheet.getLastRow();
    const customerBkSheet = ss.getSheetByName('顧客情報BK');
    const customerBkLastRow = customerBkSheet ? customerBkSheet.getLastRow() : 0;

    // 発送先情報シート
    const shippingToSheet = ss.getSheetByName('発送先情報');
    if (!shippingToSheet) {
      Logger.log('発送先情報シートが見つかりません');
      return;
    }
    const shippingToLastRow = shippingToSheet.getLastRow();
    const shippingToBkSheet = ss.getSheetByName('発送先情報BK');
    const shippingToBkLastRow = shippingToBkSheet ? shippingToBkSheet.getLastRow() : 0;

    const customerName = e.parameter.customerName || '';
    const customerCompany = e.parameter.customerCompany || '';
    const customerZipcode = e.parameter.customerZipcode || '';
    const customerAddress = e.parameter.customerAddress || '';
    const customerTel = e.parameter.customerTel || '';

    const shippingToName = e.parameter.shippingToName || '';
    const shippingToCompany = e.parameter.shippingToCompany || '';
    const shippingToZipcode = e.parameter.shippingToZipcode || '';
    const shippingToAddress = e.parameter.shippingToAddress || '';
    const shippingToTel = e.parameter.shippingToTel || '';

    Logger.log('顧客登録: 会社名=' + customerCompany + ', 氏名=' + customerName);
    Logger.log('発送先登録: 会社名=' + shippingToCompany + ', 氏名=' + shippingToName);

    // 顧客情報に既に存在するかチェック（会社名と氏名の両方が一致する場合のみ重複とみなす）
    const existingCustomers = getAllRecords('顧客情報');
    const customerExists = existingCustomers.some(function (c) {
      // 会社名と氏名の両方が一致する場合のみ重複
      // 例: 既存「株式会社田中/田中太郎」と新規「田中太郎/田中太郎」は別人として登録可能
      return c['会社名'] === customerCompany && c['氏名'] === customerName;
    });

    // 顧客情報マスタに登録（存在しない場合のみ）
    if (!customerExists && customerName) {
      // カラム: 顧客分類,表示名,フリガナ,会社名,部署,役職,氏名,郵便番号,住所１,住所２,TEL,携帯電話,FAX,メールアドレス,請求書有無,入金期日,消費税の表示方法,備考
      const customerRow = [
        '',                    // 顧客分類
        '',                    // 表示名
        '',                    // フリガナ
        customerCompany,       // 会社名
        '',                    // 部署
        '',                    // 役職
        customerName,          // 氏名
        customerZipcode,       // 郵便番号
        customerAddress,       // 住所１
        '',                    // 住所２
        customerTel,           // TEL
        '',                    // 携帯電話
        '',                    // FAX
        '',                    // メールアドレス
        '',                    // 請求書有無
        '',                    // 入金期日
        '',                    // 消費税の表示方法
        '【電話注文から自動登録】'  // 備考
      ];
      customerSheet.getRange(customerLastRow + 1, 1, 1, customerRow.length).setNumberFormat('@').setValues([customerRow]).setBorder(true, true, true, true, true, true);
      if (customerBkSheet) {
        customerBkSheet.getRange(customerBkLastRow + 1, 1, 1, customerRow.length).setNumberFormat('@').setValues([customerRow]).setBorder(true, true, true, true, true, true);
      }
      Logger.log('顧客情報マスタに登録: 会社名=' + customerCompany + ', 氏名=' + customerName);
    }

    // 発送先情報に既に存在するかチェック（会社名と氏名の両方が一致する場合のみ重複とみなす）
    const existingShippingTo = getAllRecords('発送先情報');
    Logger.log('発送先情報マスタ件数: ' + existingShippingTo.length);
    Logger.log('登録しようとしている発送先: 会社名=' + shippingToCompany + ', 氏名=' + shippingToName);

    const shippingToExists = existingShippingTo.some(function (s) {
      // 会社名と氏名の両方が一致する場合のみ重複とみなす
      const match = s['会社名'] === shippingToCompany && s['氏名'] === shippingToName;
      if (match) {
        Logger.log('発送先一致: 会社名=' + s['会社名'] + ', 氏名=' + s['氏名']);
      }
      return match;
    });

    Logger.log('発送先存在チェック結果: ' + shippingToExists);

    // 発送先情報マスタに登録（存在しない場合のみ）
    if (!shippingToExists && shippingToName) {
      // カラム: 会社名,部署,氏名,郵便番号,住所１,住所２,TEL,メールアドレス,備考
      const shippingToRow = [
        shippingToCompany,     // 会社名
        '',                    // 部署
        shippingToName,        // 氏名
        shippingToZipcode,     // 郵便番号
        shippingToAddress,     // 住所１
        '',                    // 住所２
        shippingToTel,         // TEL
        '',                    // メールアドレス
        '【電話注文から自動登録】'  // 備考
      ];
      shippingToSheet.getRange(shippingToLastRow + 1, 1, 1, shippingToRow.length).setNumberFormat('@').setValues([shippingToRow]).setBorder(true, true, true, true, true, true);
      if (shippingToBkSheet) {
        shippingToBkSheet.getRange(shippingToBkLastRow + 1, 1, 1, shippingToRow.length).setNumberFormat('@').setValues([shippingToRow]).setBorder(true, true, true, true, true, true);
      }
      Logger.log('発送先情報マスタに登録: 会社名=' + shippingToCompany + ', 氏名=' + shippingToName);
    }

  } catch (error) {
    Logger.log('registerNewCustomerToMaster エラー: ' + error.toString());
    // エラーがあっても受注処理は継続
  }
}

/**
 * 受注登録（AI自動学習フック統合、複数日程対応）
 *
 * フォームから送信された受注データをスプレッドシートに登録します。
 * 複数の配送日程に対応し、各日程ごとに受注IDを発行します。
 * 受注完了後、自動学習システムにより発送先名と顧客名の関係を学習します。
 *
 * 主要処理フロー:
 * 1. 新規顧客の場合: registerNewCustomerToMaster() でマスタ登録
 * 2. 編集モードの場合: 既存受注データ削除（受注シート + CSV）
 * 3. 日程数カウント（shippingDate1～10）
 * 4. 日程ごとにループ処理:
 *    a) 受注ID発行（generateId()）
 *    b) 商品1～10をループして受注レコード作成
 *    c) addRecords('受注', records) で受注シート登録
 *    d) 納品方法に応じてCSV登録（ヤマト/佐川）
 *    e) チェックリストに応じてPDF作成（納品書/領収書）
 * 5. AI取込一覧から遷移した場合: 仮受注データ削除
 * 6. **自動学習フック**: recordShippingMapping() で発送先→顧客のマッピング学習
 *
 * **AI自動学習システム統合（Phase 2）**:
 * - 受注完了時に recordShippingMapping(shippingToName, shippingToName, customerName) を自動呼び出し
 * - 発送先名と顧客名の関係を学習し、次回以降の顧客推定精度を向上
 * - エラー時も受注処理は継続（学習失敗が受注に影響しない設計）
 *
 * 複数日程対応:
 * - shippingDate1, deliveryDate1, shippingDate2, deliveryDate2, ...
 * - 各日程ごとに別々の受注IDを発行
 * - 最大10日程まで対応
 *
 * @param {Object} e - フォーム送信イベントオブジェクト
 *   e.parameter.isNewCustomer: 'true' or undefined
 *   e.parameter.editOrderId: 編集モードの場合の受注ID
 *   e.parameter.shippingDate1～10: 発送日
 *   e.parameter.deliveryDate1～10: 納品日
 *   e.parameter.bunrui1～10: 商品分類
 *   e.parameter.product1～10: 商品名
 *   e.parameter.price1～10: 価格
 *   e.parameter.quantity1～10: 数量
 *   e.parameter.deliveryMethod: 'ヤマト' or '佐川' or その他
 *   e.parameters.checklist: ['納品書', '領収書', ...]
 *   e.parameter.tempOrderId: AI取込一覧からの遷移時の仮受注ID
 *   e.parameter.shippingToName: 発送先名（学習用）
 *   e.parameter.customerName: 顧客名（学習用）
 *
 * @see registerNewCustomerToMaster() - 新規顧客マスタ登録
 * @see generateId() - 受注ID発行
 * @see addRecords() - 受注シート登録
 * @see addRecordYamato() - ヤマトCSV登録
 * @see addRecordSagawa() - 佐川CSV登録
 * @see createFile() - 納品書PDF作成
 * @see createReceiptFile() - 領収書PDF作成
 * @see deleteTempOrder() - 仮受注データ削除
 * @see recordShippingMapping() - AI自動学習フック（shippingMappingCode.js）
 *
 * 呼び出し元: shippingComfirm.html の [受注する] ボタン送信時
 */
function createOrder(e) {
  // 新規顧客の場合、顧客情報・発送先情報マスタに登録
  if (e.parameter.isNewCustomer === 'true') {
    registerNewCustomerToMaster(e);
  }

  // 受注シートからヘッダーを取得（シートが真実の情報源）
  const orderHeaders = getOrderSheetHeaders();

  // 編集モードの場合、既存データを削除（紐付けグループ全体を削除）
  const editOrderId = e.parameter.editOrderId || '';
  if (editOrderId) {
    // 元の納品方法と紐付けグループ情報を取得
    const oldOrder = getOrderByOrderId(editOrderId);

    if (oldOrder) {
      // 複数日程の場合: dates配列から全受注IDを取得して削除
      if (oldOrder.dates && oldOrder.dates.length > 0) {
        Logger.log('複数日程の受注グループを削除: ' + oldOrder.dates.length + '件');

        oldOrder.dates.forEach(function(dateInfo) {
          const orderIdToDelete = dateInfo.orderId;

          // CSVデータを削除
          if (oldOrder.deliveryMethod === 'ヤマト') {
            const yamatoDeleted = deleteYamatoCSVByOrderId(orderIdToDelete);
            Logger.log('ヤマトCSV削除 (' + orderIdToDelete + '): ' + yamatoDeleted + '件');
          } else if (oldOrder.deliveryMethod === '佐川') {
            const sagawaDeleted = deleteSagawaCSVByOrderId(orderIdToDelete);
            Logger.log('佐川CSV削除 (' + orderIdToDelete + '): ' + sagawaDeleted + '件');
          }

          // 受注シート削除
          const deletedCount = deleteOrderByOrderId(orderIdToDelete);
          Logger.log('受注データ削除 (' + orderIdToDelete + '): ' + deletedCount + '件');
        });
      } else {
        // 単一日程の場合: 従来通り単一IDのみ削除
        Logger.log('単一日程の受注を削除');

        // CSVデータを削除
        if (oldOrder.deliveryMethod === 'ヤマト') {
          const yamatoDeleted = deleteYamatoCSVByOrderId(editOrderId);
          Logger.log('ヤマトCSV削除: ' + yamatoDeleted + '件');
        } else if (oldOrder.deliveryMethod === '佐川') {
          const sagawaDeleted = deleteSagawaCSVByOrderId(editOrderId);
          Logger.log('佐川CSV削除: ' + sagawaDeleted + '件');
        }

        // 受注シート削除
        const deletedCount = deleteOrderByOrderId(editOrderId);
        Logger.log('受注データ削除: ' + deletedCount + '件');
      }
    }
  }
  // 日程数を取得（連続性を前提としない堅牢なカウント方法）
  let dateCount = 0;
  for (let i = 1; i <= 10; i++) {
    if (e.parameter['shippingDate' + i]) {
      dateCount = i;  // 最後に値がある日程を記録
    }
    // breakを削除: 非連続の日程入力にも対応
  }

  // 紐付け受注ID管理用（複数日程対応）
  let firstOrderId = null;

  // 日程ごとにループして受注登録
  for (let dateIndex = 1; dateIndex <= dateCount; dateIndex++) {
    const shippingDate = e.parameter['shippingDate' + dateIndex];
    const deliveryDate = e.parameter['deliveryDate' + dateIndex];
    if (!shippingDate) continue;  // 空の出荷日はスキップ（受注・納品ID作成しない）
    const deliveryNoteText = e.parameter['deliveryNoteText' + dateIndex] || '';

    // 納品ID（日程ごとに別ID）
    const deliveryId = generateId();

    // 紐付け受注IDの設定
    let linkedOrderId = '';
    if (!firstOrderId && shippingDate) {
      firstOrderId = deliveryId;  // 最初の非空出荷日の受注IDを保持
    } else {
      linkedOrderId = firstOrderId;  // 2つ目以降は最初のIDを紐付け
    }

    // 受注テーブルに複数レコードを追加する
    const records = [];
    const createRecords = [];
    var rowNum = 0;
    var dateNow = Utilities.formatDate(new Date(), 'JST', 'yyyy/MM/dd');

    for (let i = 0; i < 10; i++) {
      rowNum++;
      var bunrui = "bunrui" + rowNum;
      var product = "product" + rowNum;
      var price = "price" + rowNum;
      var quantity = "quantity" + rowNum;
      const bunruiVal = e.parameter[bunrui];
      const productVal = e.parameter[product];
      const count = Number(e.parameter[quantity]);
      const unitPrice = Number(e.parameter[price]);
      if (count > 0) {
        var record = [];
        record['受注ID'] = deliveryId;
        record['受注日'] = dateNow;
        record['顧客名'] = e.parameter.customerName;
        record['顧客郵便番号'] = e.parameter.customerZipcode;
        record['顧客住所'] = e.parameter.customerAddress;
        record['顧客電話番号'] = e.parameter.customerTel;
        record['発送先名'] = e.parameter.shippingToName;
        record['発送先郵便番号'] = e.parameter.shippingToZipcode;
        record['発送先住所'] = e.parameter.shippingToAddress;
        record['発送先電話番号'] = e.parameter.shippingToTel;
        record['発送元名'] = e.parameter.shippingFromName;
        record['発送元郵便番号'] = e.parameter.shippingFromZipcode;
        record['発送元住所'] = e.parameter.shippingFromAddress;
        record['発送元電話番号'] = e.parameter.shippingFromTel;
        record['発送日'] = shippingDate;  // ← 日程ごとの発送日
        record['納品日'] = deliveryDate;  // ← 日程ごとの納品日
        record['受付方法'] = e.parameter.receiptWay;
        record['受付者'] = e.parameter.recipient;
        record['納品方法'] = e.parameter.deliveryMethod;
        record['配達時間帯'] = e.parameter.deliveryTime ? e.parameter.deliveryTime.split(":")[1] : "";
        record['納品書'] = e.parameters.checklist ? e.parameters.checklist.includes('納品書') ? "○" : "" : "";
        record['請求書'] = e.parameters.checklist ? e.parameters.checklist.includes('請求書') ? "○" : "" : "";
        record['領収書'] = e.parameters.checklist ? e.parameters.checklist.includes('領収書') ? "○" : "" : "";
        record['パンフ'] = e.parameters.checklist ? e.parameters.checklist.includes('パンフ') ? "○" : "" : "";
        record['レシピ'] = e.parameters.checklist ? e.parameters.checklist.includes('レシピ') ? "○" : "" : "";
        record['その他添付'] = e.parameter.otherAttach;
        record['品名'] = e.parameter.sendProduct;
        record['送り状種別'] = e.parameter.invoiceType ? e.parameter.invoiceType.split(':')[1] : "";
        record['クール区分'] = e.parameter.coolCls ? e.parameter.coolCls.split(':')[1] : "";
        record['荷扱い１'] = e.parameter.cargo1 ? e.parameter.cargo1.split(':')[1] : "";
        record['荷扱い２'] = e.parameter.cargo2 ? e.parameter.cargo2.split(':')[1] : "";
        record['荷扱い３'] = e.parameter.cargo3 ? e.parameter.cargo3.split(':')[1] : "";
        record['代引総額'] = e.parameter.cashOnDelivery;
        record['代引内税'] = e.parameter.cashOnDeliTax;
        record['発行枚数'] = e.parameter.copiePrint;
        record['社内メモ'] = e.parameter.internalMemo;
        record['送り状備考欄'] = e.parameter.csvmemo;
        record['納品書備考欄'] = e.parameter.deliveryMemo;
        record['メモ'] = e.parameter.memo;
        record['出荷済'] = "";  // デフォルトは未出荷
        record['ステータス'] = ""; // ステータス列を追加
        record['追跡番号'] = ""; // 追跡番号列を追加
        record['紐付け受注ID'] = linkedOrderId;  // 紐付け受注ID（複数日程対応）
        record['納品書テキスト'] = deliveryNoteText;  // 納品書テキスト（freee用）
        record['商品分類'] = bunruiVal;
        record['商品名'] = productVal;
        record['受注数'] = count;
        record['販売価格'] = unitPrice;
        record['小計'] = count * unitPrice;

        // ヘッダー順に配列化（シートのヘッダーが真実の情報源）
        const addRecord = recordToArray(record, orderHeaders);
        records.push(addRecord);
        createRecords.push(record);
      }
    }

    // 日程ごとに登録処理を実行
    addRecords('受注', records);

    if (e.parameter.deliveryMethod == 'ヤマト') {
      addRecordYamato('ヤマトCSV', records, e, deliveryId, orderHeaders);
    }
    if (e.parameter.deliveryMethod == '佐川') {
      addRecordSagawa('佐川CSV', records, e, deliveryId, orderHeaders);
    }
    if (e.parameters.checklist && e.parameters.checklist.includes('納品書')) {
      createFile(createRecords);
    }
    if (e.parameters.checklist && e.parameters.checklist.includes('領収書')) {
      createReceiptFile(createRecords);
    }
  }

  // 仮受注データ削除（AI取込一覧からの遷移時）
  const tempOrderId = e.parameter.tempOrderId || '';
  if (tempOrderId) {
    deleteTempOrder(tempOrderId);
  }

  // 自動学習: 発送先名と顧客名の関係をマッピングシートに記録
  try {
    const shippingToName = e.parameter.shippingToName || '';
    const customerName = e.parameter.customerName || '';

    if (shippingToName && customerName) {
      // マッピングシートに記録（既存の場合は信頼度を更新）
      recordShippingMapping(shippingToName, shippingToName, customerName);
      Logger.log(`自動学習完了: ${shippingToName} → ${customerName}`);
    }
  } catch (error) {
    // 学習処理のエラーは受注処理に影響させない
    Logger.log('自動学習エラー（処理続行）: ' + error.message);
  }

  // 定期便登録処理（新規登録時のみ、編集モードでは実行しない）
  if (!editOrderId && e.parameter.isRecurringOrder === 'true') {
    try {
      // 最初の日程の発送日・納品日を基準にする（createRecurringOrder 内で発送日→納品日の間隔を保持）
      const baseShippingDate = e.parameter.shippingDate1;
      const baseDeliveryDate = e.parameter.deliveryDate1;

      if (baseShippingDate && baseDeliveryDate) {
        createRecurringOrder(e);
        Logger.log('定期便登録完了: 受注ID=' + firstOrderId);
      }
    } catch (error) {
      // 定期便登録のエラーは受注処理に影響させない
      Logger.log('定期便登録エラー（処理続行）: ' + error.message);
    }
  }
}
/**
 * 名前文字列をキャリアCSV用に会社名・個人名に分割する
 * ヤマトB2: お届け先名(16文字), 会社・部門名１(25文字)
 * 佐川e飛伝: お届け先名称１(16文字), お届け先名称２(16文字)
 *
 * 分割ルール:
 * 1. nameMaxLen以下 → そのまま name に格納
 * 2. 全角/半角スペースがあれば、最後のスペースで分割 → 前半=company, 後半=name
 *    - 後半がnameMaxLenを超える場合は最初のスペースで再試行
 * 3. スペースなし or 分割後もname超過 → companyMaxLenで前方分割
 *
 * @param {string} fullName - 元の名前文字列
 * @param {number} nameMaxLen - 名前フィールドの最大文字数
 * @param {number} companyMaxLen - 会社名フィールドの最大文字数
 * @returns {{company: string, name: string}}
 */
function splitNameForCarrier(fullName, nameMaxLen, companyMaxLen) {
  fullName = (fullName || '').toString().trim();
  if (fullName.length <= nameMaxLen) {
    return { company: '', name: fullName };
  }

  // スペース（全角・半角）で分割を試みる
  var lastFullSpace = fullName.lastIndexOf('　');
  var lastHalfSpace = fullName.lastIndexOf(' ');
  var splitIdx = Math.max(lastFullSpace, lastHalfSpace);

  if (splitIdx > 0) {
    var companyPart = fullName.substring(0, splitIdx);
    var namePart = fullName.substring(splitIdx + 1);
    // 後半がnameMaxLen以内かつ前半がcompanyMaxLen以内なら採用
    if (namePart.length <= nameMaxLen && companyPart.length <= companyMaxLen) {
      return { company: companyPart, name: namePart };
    }
    // 最後のスペースでダメなら最初のスペースで試行
    var firstFullSpace = fullName.indexOf('　');
    var firstHalfSpace = fullName.indexOf(' ');
    var firstSplitIdx = -1;
    if (firstFullSpace > 0 && firstHalfSpace > 0) {
      firstSplitIdx = Math.min(firstFullSpace, firstHalfSpace);
    } else {
      firstSplitIdx = Math.max(firstFullSpace, firstHalfSpace);
    }
    if (firstSplitIdx > 0 && firstSplitIdx !== splitIdx) {
      companyPart = fullName.substring(0, firstSplitIdx);
      namePart = fullName.substring(firstSplitIdx + 1);
      if (namePart.length <= nameMaxLen && companyPart.length <= companyMaxLen) {
        return { company: companyPart, name: namePart };
      }
    }
  }

  // スペースなし or 分割後も超過 → 強制的にcompanyに前方、nameに後方
  var company = fullName.substring(0, Math.min(fullName.length, companyMaxLen));
  var name = fullName.substring(company.length);
  if (name.length > nameMaxLen) {
    name = name.substring(0, nameMaxLen); // 最終手段: 切り詰め
  }
  return { company: company, name: name };
}

/**
 * ヤマト運輸B2用CSVデータを「ヤマトCSV」シートに登録
 *
 * @param {string} sheetName - シート名（通常 'ヤマトCSV'）
 * @param {Array} records - 受注レコード配列
 * @param {Object} e - オリジナルのリクエストパラメータ
 * @param {string} deliveryId - 受注ID
 * @param {Array} orderHeaders - 受注シートのヘッダー配列
 */
function addRecordYamato(sheetName, records, e, deliveryId, orderHeaders) {
  Logger.log('addRecordYamato: sheet=' + sheetName + ', deliveryId=' + (deliveryId || '') + ', recordsCount=' + (records ? records.length : 0) + ', orderHeadersCount=' + (orderHeaders ? orderHeaders.length : 0));
  const adds = [];
  var record = [];
  // 受注データから値を取得（ヘッダー名でアクセス）
  const orderData = records[0];
  record['発送日'] = Utilities.formatDate(new Date(getOrderValue(orderData, '発送日', orderHeaders)), 'JST', 'yyyy/MM/dd');
  record['お客様管理番号'] = deliveryId || "";
  record['送り状種別'] = e.parameter.invoiceType ? e.parameter.invoiceType.split(':')[0] : "";
  record['クール区分'] = e.parameter.coolCls ? e.parameter.coolCls.split(':')[0] : "";
  record['伝票番号'] = "";
  record['出荷予定日'] = Utilities.formatDate(new Date(getOrderValue(orderData, '発送日', orderHeaders)), 'JST', 'yyyy/MM/dd');
  record['お届け予定（指定）日'] = Utilities.formatDate(new Date(getOrderValue(orderData, '納品日', orderHeaders)), 'JST', 'yyyy/MM/dd');
  record['配達時間帯'] = e.parameter.deliveryTime ? e.parameter.deliveryTime.split(":")[0] : "";
  record['お届け先コード'] = "";
  record['お届け先電話番号'] = getOrderValue(orderData, '発送先電話番号', orderHeaders);
  record['お届け先電話番号枝番'] = "";
  record['お届け先郵便番号'] = getOrderValue(orderData, '発送先郵便番号', orderHeaders);
  // ヤマトB2: お届け先住所・お届け先住所（アパートマンション名）は各16文字まで。超える場合は16文字目以降を次カラムへ
  (function () {
    const ADDRESS_MAX_LEN = 16;
    var toAddr = (getOrderValue(orderData, '発送先住所', orderHeaders) || '').toString();
    record['お届け先住所'] = toAddr.length > ADDRESS_MAX_LEN ? toAddr.substring(0, ADDRESS_MAX_LEN) : toAddr;
    record['お届け先住所（アパートマンション名）'] = toAddr.length > ADDRESS_MAX_LEN ? toAddr.substring(ADDRESS_MAX_LEN) : "";
  })();
  // ヤマトB2: お届け先名(16文字), 会社・部門名１(25文字) — 長い名前を自動分割
  (function () {
    var toName = splitNameForCarrier(
      getOrderValue(orderData, '発送先名', orderHeaders), 16, 25
    );
    record['お届け先会社・部門名１'] = toName.company;
    record['お届け先会社・部門名２'] = "";
    record['お届け先名'] = toName.name;
  })();
  record['お届け先名略称カナ'] = "";
  record['敬称'] = "";
  record['ご依頼主コード'] = "";
  record['ご依頼主電話番号'] = getOrderValue(orderData, '発送元電話番号', orderHeaders);
  record['ご依頼主電話番号枝番'] = "";
  record['ご依頼主郵便番号'] = getOrderValue(orderData, '発送元郵便番号', orderHeaders);
  // ヤマトB2: ご依頼主住所・ご依頼主住所（アパートマンション名）は各16文字まで。超える場合は16文字目以降を次カラムへ
  (function () {
    const ADDRESS_MAX_LEN = 16;
    var fromAddr = (getOrderValue(orderData, '発送元住所', orderHeaders) || '').toString();
    record['ご依頼主住所'] = fromAddr.length > ADDRESS_MAX_LEN ? fromAddr.substring(0, ADDRESS_MAX_LEN) : fromAddr;
    record['ご依頼主住所（アパートマンション名）'] = fromAddr.length > ADDRESS_MAX_LEN ? fromAddr.substring(ADDRESS_MAX_LEN) : "";
  })();
  // ヤマトB2: ご依頼主名(16文字), ご依頼主名（カナ）は使用しないため会社名分割不要だが
  // ご依頼主名も16文字制限があるため同様に分割（会社・部門名フィールドがないので切り詰め）
  (function () {
    var fromName = (getOrderValue(orderData, '発送元名', orderHeaders) || '').toString();
    record['ご依頼主名'] = fromName.length > 16 ? fromName.substring(0, 16) : fromName;
  })();
  record['ご依頼主略称カナ'] = "";
  record['品名コード１'] = "";
  record['品名１'] = getOrderValue(orderData, '品名', orderHeaders);
  record['品名コード２'] = "";
  record['品名２'] = "";
  record['荷扱い１'] = e.parameter.cargo1 ? e.parameter.cargo1.split(':')[0] : "";
  record['荷扱い２'] = e.parameter.cargo2 ? e.parameter.cargo2.split(':')[0] : "";
  record['荷扱い３'] = e.parameter.cargo3 ? e.parameter.cargo3.split(':')[0] : "";
  record['記事'] = getOrderValue(orderData, '送り状備考欄', orderHeaders);
  record['コレクト代金引換額（税込）'] = getOrderValue(orderData, '代引総額', orderHeaders);
  record['コレクト内消費税額等'] = getOrderValue(orderData, '代引内税', orderHeaders);
  record['営業所止置き'] = "";
  record['営業所コード'] = "";
  record['発行枚数'] = getOrderValue(orderData, '発行枚数', orderHeaders);
  record['個数口枠の印字'] = "";
  record['ご請求先顧客コード'] = "019543385101";
  record['ご請求先分類コード'] = "";
  record['運賃管理番号'] = "01";
  record['クロネコwebコレクトデータ登録'] = "";
  record['クロネコwebコレクト加盟店番号'] = "";
  record['クロネコwebコレクト申込受付番号１'] = "";
  record['クロネコwebコレクト申込受付番号２'] = "";
  record['クロネコwebコレクト申込受付番号３'] = "";
  record['お届け予定ｅメール利用区分'] = "";
  record['お届け予定ｅメールe-mailアドレス'] = "";
  record['入力機種'] = "";
  record['お届け予定eメールメッセージ'] = "";
  const addList = [
    record['発送日'],
    record['お客様管理番号'],
    record['送り状種別'],
    record['クール区分'],
    record['伝票番号'],
    record['出荷予定日'],
    record['お届け予定（指定）日'],
    record['配達時間帯'],
    record['お届け先コード'],
    record['お届け先電話番号'],
    record['お届け先電話番号枝番'],
    record['お届け先郵便番号'],
    record['お届け先住所'],
    record['お届け先住所（アパートマンション名）'],
    record['お届け先会社・部門名１'],
    record['お届け先会社・部門名２'],
    record['お届け先名'],
    record['お届け先名略称カナ'],
    record['敬称'],
    record['ご依頼主コード'],
    record['ご依頼主電話番号'],
    record['ご依頼主電話番号枝番'],
    record['ご依頼主郵便番号'],
    record['ご依頼主住所'],
    record['ご依頼主住所（アパートマンション名）'],
    record['ご依頼主名'],
    record['ご依頼主略称カナ'],
    record['品名コード１'],
    record['品名１'],
    record['品名コード２'],
    record['品名２'],
    record['荷扱い１'],
    record['荷扱い２'],
    record['記事'],
    record['コレクト代金引換額（税込）'],
    record['コレクト内消費税額等'],
    record['営業所止置き'],
    record['営業所コード'],
    record['発行枚数'],
    record['個数口枠の印字'],
    record['ご請求先顧客コード'],
    record['ご請求先分類コード'],
    record['運賃管理番号'],
    record['クロネコwebコレクトデータ登録'],
    record['クロネコwebコレクト加盟店番号'],
    record['クロネコwebコレクト申込受付番号１'],
    record['クロネコwebコレクト申込受付番号２'],
    record['クロネコwebコレクト申込受付番号３'],
    record['お届け予定ｅメール利用区分'],
    record['お届け予定ｅメールe-mailアドレス'],
    record['入力機種'],
    record['お届け予定eメールメッセージ']
  ];
  adds.push(addList);
  addRecords(sheetName, adds);
}
/**
 * 佐川急便e飛伝用CSVデータを「佐川CSV」シートに登録
 * 
 * @param {string} sheetName - シート名（通常 '佐川CSV'）
 * @param {Array} records - 受注レコード配列
 * @param {Object} e - オリジナルのリクエストパラメータ
 * @param {string} deliveryId - 受注ID
 * @param {Array} orderHeaders - 受注シートのヘッダー配列
 */
function addRecordSagawa(sheetName, records, e, deliveryId, orderHeaders) {
  Logger.log('addRecordSagawa: sheet=' + sheetName + ', deliveryId=' + (deliveryId || '') + ', recordsCount=' + (records ? records.length : 0) + ', orderHeadersCount=' + (orderHeaders ? orderHeaders.length : 0));
  const adds = [];
  var record = [];
  // 受注データから値を取得（ヘッダー名でアクセス）
  const orderData = records[0];
  record['発送日'] = Utilities.formatDate(new Date(getOrderValue(orderData, '発送日', orderHeaders)), 'JST', 'yyyy/MM/dd');
  record['お届け先コード取得区分'] = "";
  record['お届け先コード'] = "";
  record['お届け先電話番号'] = getOrderValue(orderData, '発送先電話番号', orderHeaders);
  record['お届け先郵便番号'] = getOrderValue(orderData, '発送先郵便番号', orderHeaders);
  record['お届け先住所１'] = getOrderValue(orderData, '発送先住所', orderHeaders);
  record['お届け先住所２'] = "";
  record['お届け先住所３'] = "";
  // 佐川e飛伝: お届け先名称１(16文字), お届け先名称２(16文字) — 長い名前を自動分割
  (function () {
    var toName = splitNameForCarrier(
      getOrderValue(orderData, '発送先名', orderHeaders), 16, 16
    );
    record['お届け先名称１'] = toName.company || toName.name;
    record['お届け先名称２'] = toName.company ? toName.name : '';
  })();
  record['お客様管理番号'] = deliveryId || "";
  record['お客様コード'] = "";
  record['部署ご担当者コード取得区分'] = "";
  record['部署ご担当者コード'] = "";
  record['部署ご担当者名称'] = "";
  record['荷送人電話番号'] = "";
  record['ご依頼主コード取得区分'] = "";
  record['ご依頼主コード'] = "";
  record['ご依頼主電話番号'] = getOrderValue(orderData, '発送元電話番号', orderHeaders);
  record['ご依頼主郵便番号'] = getOrderValue(orderData, '発送元郵便番号', orderHeaders);
  record['ご依頼主住所１'] = getOrderValue(orderData, '発送元住所', orderHeaders);
  record['ご依頼主住所２'] = "";
  // 佐川e飛伝: ご依頼主名称１(16文字), ご依頼主名称２(16文字) — 長い名前を自動分割
  (function () {
    var fromName = splitNameForCarrier(
      getOrderValue(orderData, '発送元名', orderHeaders), 16, 16
    );
    record['ご依頼主名称１'] = fromName.company || fromName.name;
    record['ご依頼主名称２'] = fromName.company ? fromName.name : '';
  })();
  record['荷姿'] = "";
  const productName = getOrderValue(orderData, '品名', orderHeaders) || '';
  if (productName.length > 16) {
    record['品名１'] = productName.substring(0, 16);
  }
  else {
    record['品名１'] = productName;
  }
  record['品名２'] = "";
  record['品名３'] = "";
  record['品名４'] = "";
  record['品名５'] = "";
  record['荷札荷姿'] = "";
  record['荷札品名１'] = "";
  record['荷札品名２'] = "";
  record['荷札品名３'] = "";
  record['荷札品名４'] = "";
  record['荷札品名５'] = "";
  record['荷札品名６'] = "";
  record['荷札品名７'] = "";
  record['荷札品名８'] = "";
  record['荷札品名９'] = "";
  record['荷札品名１０'] = "";
  record['荷札品名１１'] = "";
  record['出荷個数'] = getOrderValue(orderData, '発行枚数', orderHeaders);
  record['スピード指定'] = "000";
  record['クール便指定'] = e.parameter.coolCls ? e.parameter.coolCls.split(':')[0] : "";
  record['配達日'] = Utilities.formatDate(new Date(getOrderValue(orderData, '納品日', orderHeaders)), 'JST', 'yyyyMMdd');
  record['配達指定時間帯'] = e.parameter.deliveryTime ? e.parameter.deliveryTime.split(":")[0] : "";
  record['配達指定時間（時分）'] = "";
  record['代引金額'] = getOrderValue(orderData, '代引総額', orderHeaders);
  record['消費税'] = getOrderValue(orderData, '代引内税', orderHeaders);
  record['決済種別'] = "";
  record['保険金額'] = "";
  record['指定シール１'] = e.parameter.cargo1 ? e.parameter.cargo1.split(':')[0] : "";
  record['指定シール２'] = e.parameter.cargo2 ? e.parameter.cargo2.split(':')[0] : "";
  record['指定シール３'] = e.parameter.cargo3 ? e.parameter.cargo3.split(':')[0] : "";
  record['営業所受取'] = "";
  record['SRC区分'] = "";
  record['営業所受取営業所コード'] = "";
  record['元着区分'] = e.parameter.invoiceType ? e.parameter.invoiceType.split(':')[0] : "";
  record['メールアドレス'] = "";
  record['ご不在時連絡先'] = getOrderValue(orderData, '発送先電話番号', orderHeaders);
  record['出荷予定日'] = "";
  record['セット数'] = "";
  record['お問い合せ送り状No.'] = "";
  record['出荷場印字区分'] = "";
  record['集約解除指定'] = "";
  record['編集０１'] = "";
  record['編集０２'] = "";
  record['編集０３'] = "";
  record['編集０４'] = "";
  record['編集０５'] = "";
  record['編集０６'] = "";
  record['編集０７'] = "";
  record['編集０８'] = "";
  record['編集０９'] = "";
  record['編集１０'] = "";
  const addList = [
    record['発送日'],
    record['お届け先コード取得区分'],
    record['お届け先コード'],
    record['お届け先電話番号'],
    record['お届け先郵便番号'],
    record['お届け先住所１'],
    record['お届け先住所２'],
    record['お届け先住所３'],
    record['お届け先名称１'],
    record['お届け先名称２'],
    record['お客様管理番号'],
    record['お客様コード'],
    record['部署ご担当者コード取得区分'],
    record['部署ご担当者コード'],
    record['部署ご担当者名称'],
    record['荷送人電話番号'],
    record['ご依頼主コード取得区分'],
    record['ご依頼主コード'],
    record['ご依頼主電話番号'],
    record['ご依頼主郵便番号'],
    record['ご依頼主住所１'],
    record['ご依頼主住所２'],
    record['ご依頼主名称１'],
    record['ご依頼主名称２'],
    record['荷姿'],
    record['品名１'],
    record['品名２'],
    record['品名３'],
    record['品名４'],
    record['品名５'],
    record['荷札荷姿'],
    record['荷札品名１'],
    record['荷札品名２'],
    record['荷札品名３'],
    record['荷札品名４'],
    record['荷札品名５'],
    record['荷札品名６'],
    record['荷札品名７'],
    record['荷札品名８'],
    record['荷札品名９'],
    record['荷札品名１０'],
    record['荷札品名１１'],
    record['出荷個数'],
    record['スピード指定'],
    record['クール便指定'],
    record['配達日'],
    record['配達指定時間帯'],
    record['配達指定時間（時分）'],
    record['代引金額'],
    record['消費税'],
    record['決済種別'],
    record['保険金額'],
    record['指定シール１'],
    record['指定シール２'],
    record['指定シール３'],
    record['営業所受取'],
    record['SRC区分'],
    record['営業所受取営業所コード'],
    record['元着区分'],
    record['メールアドレス'],
    record['ご不在時連絡先'],
    record['出荷予定日'],
    record['セット数'],
    record['お問い合せ送り状No.'],
    record['出荷場印字区分'],
    record['集約解除指定'],
    record['編集０１'],
    record['編集０２'],
    record['編集０３'],
    record['編集０４'],
    record['編集０５'],
    record['編集０６'],
    record['編集０７'],
    record['編集０８'],
    record['編集０９'],
    record['編集１０']
  ];
  adds.push(addList);
  addRecords(sheetName, adds);
}
// レコード登録
function addRecords(sheetName, records) {
  Logger.log(sheetName);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  const lastRow = sheet.getLastRow();
  sheet.getRange(lastRow + 1, 1, records.length, records[0].length).setNumberFormat('@').setValues(records).setBorder(true, true, true, true, true, true);
}
// 在庫更新
function updateZaiko(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('商品');
  var rowNum = 0;
  for (let i = 0; i < 10; i++) {
    rowNum++;
    var product = "product" + rowNum;
    var quantity = "quantity" + rowNum;
    const productVal = e.parameter[product];
    const count = Number(e.parameter[quantity]);
    var zaiko = 0;
    if (count > 0) {
      const targetRow = sheet.getRange('B:B').createTextFinder(productVal).matchEntireCell(true).findNext().getRow();
      const targetCol = sheet.getRange('1:1').createTextFinder('在庫数').matchEntireCell(true).findNext().getColumn();
      zaiko = Number(sheet.getRange(targetRow, targetCol).getValue());
      sheet.getRange(targetRow, targetCol).setValue(zaiko - count);
    }
  }
}

// 納品書・領収書用のテンプレートとフォルダを取得する関数
// グローバルスコープでの初期化を避け、必要時に取得する
function getDeliveredTemplate() {
  const templateId = getDeliveredTemplateId();
  if (!templateId) {
    throw new Error('DELIVERED_TEMPLATE_ID がスクリプトプロパティに設定されていません');
  }
  try {
    return DriveApp.getFileById(templateId);
  } catch (e) {
    throw new Error(`納品書テンプレート(ID: ${templateId})へのアクセスに失敗しました: ${e.message}`);
  }
}

function getDeliveredPdfOutdir() {
  const folderId = getDeliveredPdfFolderId();
  if (!folderId) {
    throw new Error('DELIVERED_PDF_FOLDER_ID がスクリプトプロパティに設定されていません');
  }
  try {
    return DriveApp.getFolderById(folderId);
  } catch (e) {
    throw new Error(`納品書PDF出力フォルダ(ID: ${folderId})へのアクセスに失敗しました: ${e.message}`);
  }
}

function getReceiptTemplate() {
  const templateId = getReceiptTemplateId();
  if (!templateId) {
    throw new Error('RECEIPT_TEMPLATE_ID がスクリプトプロパティに設定されていません');
  }
  try {
    return DriveApp.getFileById(templateId);
  } catch (e) {
    throw new Error(`領収書テンプレート(ID: ${templateId})へのアクセスに失敗しました: ${e.message}`);
  }
}

function getReceiptPdfOutdir() {
  const folderId = getReceiptPdfFolderId();
  if (!folderId) {
    throw new Error('RECEIPT_PDF_FOLDER_ID がスクリプトプロパティに設定されていません');
  }
  try {
    return DriveApp.getFolderById(folderId);
  } catch (e) {
    throw new Error(`領収書PDF出力フォルダ(ID: ${folderId})へのアクセスに失敗しました: ${e.message}`);
  }
}

// 納品書ファイル生成
function createFile(records) {
  // PDF変換する元ファイルを作成する
  let wFileRtn = createGDoc(records);
  // PDF変換
  createDeliveredPdf(wFileRtn[0], wFileRtn[1], wFileRtn[2]);
  // PDF変換したあとは元ファイルを削除する
  DriveApp.getFileById(wFileRtn[0]).setTrashed(true);
}

// 納品書のドキュメントの中身を置換
function createGDoc(rowVal) {
  // 顧客情報シートを取得する
  const customerItems = getAllRecords('顧客情報');
  const shippingToItems = getAllRecords('発送先情報');
  const productItems = getAllRecords('商品');
  var customerItem = [];
  Logger.log(rowVal);
  Logger.log(rowVal[0]);
  Logger.log(rowVal[0]['顧客名']);
  const shippingName = rowVal[0]['顧客名'].split('　');
  Logger.log(shippingName);
  // データ走査
  customerItems.forEach(function (wVal) {
    if (shippingName.length > 1) {
      // 会社名と同じ
      if (shippingName[1] == wVal['氏名'] && shippingName[0] == wVal['会社名']) {
        customerItem = wVal;
      }
    }
    else {
      // 会社名と同じ
      if (shippingName[0] == wVal['氏名']) {
        customerItem = wVal;
      }
      if (shippingName[0] == wVal['会社名']) {
        customerItem = wVal;
      }
    }
  });
  if (customerItem.length == 0) {
    customerItem['会社名'] = rowVal[0]['顧客名'].split('　')[0];
    customerItem['住所１'] = rowVal[0]['顧客住所'];
    customerItem['郵便番号'] = rowVal[0]['顧客郵便番号'];
    customerItem['住所２'] = "";
  }
  // テンプレートファイルをコピーする
  const wCopyFile = getDeliveredTemplate().makeCopy()
    , wCopyFileId = wCopyFile.getId()
    , wCopyDoc = DocumentApp.openById(wCopyFileId); // コピーしたファイルをGoogleドキュメントとして開く
  let wCopyDocBody = wCopyDoc.getBody(); // Googleドキュメント内の本文を取得する
  var post = String(customerItem['郵便番号']);
  post = post.substring(0, 3).concat("-").concat(post.substring(3, 7));

  // 注文書ファイル内の可変文字部（として用意していた箇所）を変更する
  wCopyDocBody = wCopyDocBody.replaceText('{{company_name}}', customerItem['会社名'] ? customerItem['会社名'] : (customerItem['氏名'] || ''));
  wCopyDocBody = wCopyDocBody.replaceText('{{post}}', post || '');
  wCopyDocBody = wCopyDocBody.replaceText('{{address1}}', customerItem['住所１'] || '');
  wCopyDocBody = wCopyDocBody.replaceText('{{address2}}', customerItem['住所２'] || '');
  wCopyDocBody = wCopyDocBody.replaceText('{{delivery_num}}', rowVal[0]['受注ID'] || '');
  wCopyDocBody = wCopyDocBody.replaceText('{{delivery_date}}', Utilities.formatDate(new Date(rowVal[0]['納品日']), 'JST', 'yyyy年MM月dd日'));
  wCopyDocBody = wCopyDocBody.replaceText('{{deliveryMemo}}', rowVal[0]['納品書備考欄'] || '');
  let totals = 0;
  let tentax = 0;
  let eigtax = 0;
  let tentax_t = 0;
  let eigtax_t = 0;
  let tax = 0;
  let amount = 0;

  for (let i = 0; i < 10; i++) {
    var total = 0;
    if (i < rowVal.length) {
      var productData;
      // 商品分類
      var changeText = '{{bunrui' + (i + 1) + '}}';
      wCopyDocBody = wCopyDocBody.replaceText(changeText, rowVal[i]['商品分類']);
      // 商品名
      changeText = '{{product' + (i + 1) + '}}';
      wCopyDocBody = wCopyDocBody.replaceText(changeText, rowVal[i]['商品名']);
      // データ走査
      productItems.forEach(function (wVal) {
        // 商品名と同じ
        if (wVal['商品名'] == rowVal[i]['商品名']) {
          productData = wVal;
        }
      });
      // 価格（P)
      changeText = '{{price' + (i + 1) + '}}';
      wCopyDocBody = wCopyDocBody.replaceText(changeText, '￥ ' + rowVal[i]['販売価格']);
      // 数量
      changeText = '{{c' + (i + 1) + '}}';
      wCopyDocBody = wCopyDocBody.replaceText(changeText, rowVal[i]['受注数']);
      // 金額
      changeText = '{{amount' + (i + 1) + '}}';
      total = rowVal[i]['受注数'] * rowVal[i]['販売価格'];
      if (Number(productData['税率']) > 8) {
        var taxValTotal = Math.round(Number(total * 1.1));
        var taxVal = taxValTotal - Number(total);
        tentax += taxVal;
        tentax_t += total;
        tax += taxVal;
        amount += total;
        totals += taxValTotal;
      }
      else {
        var taxValTotal = Math.round(Number(total * 1.08));
        var taxVal = taxValTotal - Number(total);
        eigtax += taxVal;
        eigtax_t += total;
        tax += taxVal;
        amount += total;
        totals += taxValTotal;
      }
      wCopyDocBody = wCopyDocBody.replaceText(changeText, '￥ ' + total.toLocaleString());
    }
    else {
      var changeText = '{{bunrui' + (i + 1) + '}}';
      wCopyDocBody = wCopyDocBody.replaceText(changeText, '');
      changeText = '{{product' + (i + 1) + '}}';
      wCopyDocBody = wCopyDocBody.replaceText(changeText, '');
      changeText = '{{price' + (i + 1) + '}}';
      wCopyDocBody = wCopyDocBody.replaceText(changeText, '');
      changeText = '{{c' + (i + 1) + '}}';
      wCopyDocBody = wCopyDocBody.replaceText(changeText, '');
      changeText = '{{amount' + (i + 1) + '}}';
      wCopyDocBody = wCopyDocBody.replaceText(changeText, '');
    }
  }
  wCopyDocBody = wCopyDocBody.replaceText('{{amount}}', amount.toLocaleString());
  wCopyDocBody = wCopyDocBody.replaceText('{{tax}}', tax.toLocaleString());
  wCopyDocBody = wCopyDocBody.replaceText('{{10tax_t}}', tentax_t.toLocaleString());
  wCopyDocBody = wCopyDocBody.replaceText('{{10tax}}', tentax.toLocaleString());
  wCopyDocBody = wCopyDocBody.replaceText('{{8tax_t}}', eigtax_t.toLocaleString());
  wCopyDocBody = wCopyDocBody.replaceText('{{8tax}}', eigtax.toLocaleString());
  wCopyDocBody = wCopyDocBody.replaceText('{{amount_delivered}}', totals.toLocaleString());
  wCopyDocBody = wCopyDocBody.replaceText('{{total}}', '￥ ' + totals.toLocaleString());
  wCopyDoc.saveAndClose();

  // ファイル名を変更する
  let fileName = Utilities.formatDate(new Date(), 'JST', 'yyyyMMdd') + "_" + customerItem['会社名'] + ' 御中';
  wCopyFile.setName(fileName);
  // コピーしたファイルIDとファイル名を返却する（あとでこのIDをもとにPDFに変換するため）
  return [wCopyFileId, fileName, customerItem['会社名']];
}
// PDF生成
function createDeliveredPdf(docId, fileName, targetFolderName) {
  // PDF変換するためのベースURLを作成する
  let wUrl = `https://docs.google.com/document/d/${docId}/export?exportFormat=pdf`;

  // headersにアクセストークンを格納する
  let wOtions = {
    headers: {
      'Authorization': `Bearer ${ScriptApp.getOAuthToken()}`
    }
  };
  // PDFを作成する
  let wBlob = UrlFetchApp.fetch(wUrl, wOtions).getBlob().setName(fileName + '.pdf');
  // 保存先のフォルダが存在するか確認
  const deliveredPdfOutdir = getDeliveredPdfOutdir();
  var targetFolder = null;
  var folders = deliveredPdfOutdir.getFoldersByName(targetFolderName);

  if (folders.hasNext()) {
    // 既存のフォルダが見つかった場合
    targetFolder = folders.next();
  } else {
    // フォルダが存在しない場合は新規作成
    targetFolder = deliveredPdfOutdir.createFolder(targetFolderName);
  }
  //PDFを指定したフォルダに保存する
  return targetFolder.createFile(wBlob).getId();
}
// PDF生成
function createReceiptPdf(docId, fileName) {
  // PDF変換するためのベースURLを作成する
  let wUrl = `https://docs.google.com/document/d/${docId}/export?exportFormat=pdf`;

  // headersにアクセストークンを格納する
  let wOtions = {
    headers: {
      'Authorization': `Bearer ${ScriptApp.getOAuthToken()}`
    }
  };
  // PDFを作成する
  let wBlob = UrlFetchApp.fetch(wUrl, wOtions).getBlob().setName(fileName + '.pdf');

  //PDFを指定したフォルダに保存する
  return getReceiptPdfOutdir().createFile(wBlob).getId();
}
// 領収書のドキュメントを置換
function createReceiptGDoc(rowVal) {
  // 顧客情報シートを取得する
  const customerItems = getAllRecords('顧客情報');
  const shippingToItems = getAllRecords('発送先情報');
  const productItems = getAllRecords('商品');
  var customerItem = [];
  Logger.log(rowVal);
  Logger.log(rowVal[0]);
  Logger.log(rowVal[0]['顧客名']);
  const shippingName = rowVal[0]['顧客名'].split('　')[0];
  Logger.log(shippingName);
  // データ走査
  shippingToItems.forEach(function (wVal) {
    if (shippingName.length > 1) {
      // 会社名と同じ
      if (shippingName[1] == wVal['氏名'] && shippingName[0] == wVal['会社名']) {
        customerItem = wVal;
      }
    }
    else {
      // 会社名と同じ
      if (shippingName == wVal['氏名']) {
        customerItem = wVal;
      }
    }
  });
  if (customerItem.length == 0) {
    customerItem['会社名'] = rowVal[0]['顧客名'].split('　')[0];
    customerItem['住所１'] = rowVal[0]['顧客住所'];
    customerItem['郵便番号'] = rowVal[0]['顧客郵便番号'];
    customerItem['住所２'] = "";
  }
  // テンプレートファイルをコピーする
  const wCopyFile = getReceiptTemplate().makeCopy()
    , wCopyFileId = wCopyFile.getId()
    , wCopyDoc = DocumentApp.openById(wCopyFileId); // コピーしたファイルをGoogleドキュメントとして開く
  let wCopyDocBody = wCopyDoc.getBody(); // Googleドキュメント内の本文を取得する
  var post = String(customerItem['郵便番号']);
  post = post.substring(0, 3).concat("-").concat(post.substring(3, 7));

  // 注文書ファイル内の可変文字部（として用意していた箇所）を変更する
  wCopyDocBody = wCopyDocBody.replaceText('{{company_name}}', customerItem['会社名'] ? customerItem['会社名'] : customerItem['氏名']);
  wCopyDocBody = wCopyDocBody.replaceText('{{post}}', post);
  wCopyDocBody = wCopyDocBody.replaceText('{{address1}}', customerItem['住所１']);
  wCopyDocBody = wCopyDocBody.replaceText('{{address2}}', customerItem['住所２']);
  wCopyDocBody = wCopyDocBody.replaceText('{{delivery_num}}', rowVal[0]['受注ID']);
  wCopyDocBody = wCopyDocBody.replaceText('{{delivery_date}}', Utilities.formatDate(new Date(rowVal[0]['納品日']), 'JST', 'yyyy年MM月dd日'));
  let totals = 0;
  let tentax = 0;
  let eigtax = 0;
  let tentax_t = 0;
  let eigtax_t = 0;
  let tax = 0;
  let amount = 0;

  for (let i = 0; i < 10; i++) {
    var total = 0;
    if (i < rowVal.length) {
      var productData;
      // 商品分類
      var changeText = '{{bunrui' + (i + 1) + '}}';
      wCopyDocBody = wCopyDocBody.replaceText(changeText, rowVal[i]['商品分類']);
      // 商品名
      changeText = '{{product' + (i + 1) + '}}';
      wCopyDocBody = wCopyDocBody.replaceText(changeText, rowVal[i]['商品名']);
      // データ走査
      productItems.forEach(function (wVal) {
        // 商品名と同じ
        if (wVal['商品名'] == rowVal[i]['商品名']) {
          productData = wVal;
        }
      });
      // 価格（P)
      changeText = '{{price' + (i + 1) + '}}';
      wCopyDocBody = wCopyDocBody.replaceText(changeText, '￥ ' + rowVal[i]['販売価格']);
      // 数量
      changeText = '{{c' + (i + 1) + '}}';
      wCopyDocBody = wCopyDocBody.replaceText(changeText, rowVal[i]['受注数']);
      // 金額
      changeText = '{{amount' + (i + 1) + '}}';
      total = rowVal[i]['受注数'] * rowVal[i]['販売価格'];
      if (Number(productData['税率']) > 8) {
        var taxValTotal = Math.round(Number(total * 1.1));
        var taxVal = taxValTotal - Number(total);
        tentax += taxVal;
        tentax_t += total;
        tax += taxVal;
        amount += total;
        totals += taxValTotal;
      }
      else {
        var taxValTotal = Math.round(Number(total * 1.08));
        var taxVal = taxValTotal - Number(total);
        eigtax += taxVal;
        eigtax_t += total;
        tax += taxVal;
        amount += total;
        totals += taxValTotal;
      }
      wCopyDocBody = wCopyDocBody.replaceText(changeText, '￥ ' + total.toLocaleString());
    }
    else {
      var changeText = '{{bunrui' + (i + 1) + '}}';
      wCopyDocBody = wCopyDocBody.replaceText(changeText, '');
      changeText = '{{product' + (i + 1) + '}}';
      wCopyDocBody = wCopyDocBody.replaceText(changeText, '');
      changeText = '{{price' + (i + 1) + '}}';
      wCopyDocBody = wCopyDocBody.replaceText(changeText, '');
      changeText = '{{c' + (i + 1) + '}}';
      wCopyDocBody = wCopyDocBody.replaceText(changeText, '');
      changeText = '{{amount' + (i + 1) + '}}';
      wCopyDocBody = wCopyDocBody.replaceText(changeText, '');
    }
  }
  wCopyDocBody = wCopyDocBody.replaceText('{{amount}}', amount.toLocaleString());
  wCopyDocBody = wCopyDocBody.replaceText('{{tax}}', tax.toLocaleString());
  wCopyDocBody = wCopyDocBody.replaceText('{{10tax_t}}', tentax_t.toLocaleString());
  wCopyDocBody = wCopyDocBody.replaceText('{{10tax}}', tentax.toLocaleString());
  wCopyDocBody = wCopyDocBody.replaceText('{{8tax_t}}', eigtax_t.toLocaleString());
  wCopyDocBody = wCopyDocBody.replaceText('{{8tax}}', eigtax.toLocaleString());
  wCopyDocBody = wCopyDocBody.replaceText('{{amount_delivered}}', totals.toLocaleString());
  wCopyDocBody = wCopyDocBody.replaceText('{{total}}', '￥ ' + totals.toLocaleString());
  wCopyDoc.saveAndClose();

  // ファイル名を変更する
  let fileName = customerItem['会社名'] + '領収文書_' + Utilities.formatDate(new Date(), 'JST', 'yyyyMMdd');
  wCopyFile.setName(fileName);
  // コピーしたファイルIDとファイル名を返却する（あとでこのIDをもとにPDFに変換するため）
  return [wCopyFileId, fileName];
}
// 領収書のファイル生成
function createReceiptFile(records) {
  // PDF変換する元ファイルを作成する
  let wFileRtn = createReceiptGDoc(records);
  // PDF変換
  createReceiptPdf(wFileRtn[0], wFileRtn[1]);
  // PDF変換したあとは元ファイルを削除する
  DriveApp.getFileById(wFileRtn[0]).setTrashed(true);
}

/**
 * 仮受注データ削除（AI取込一覧からの遷移完了時のクリーンアップ）
 *
 * AI取込一覧から受注入力画面に遷移して受注登録が完了した際、
 * 元の仮受注データ（'仮受注'シート）を削除します。
 *
 * 処理フロー:
 * 1. LINE Bot用スプレッドシート取得
 * 2. '仮受注'シート取得
 * 3. シートなし: return（何もしない）
 * 4. 全データを取得してループ
 * 5. A列（仮受注ID）が一致する行を発見
 * 6. sheet.deleteRow() で行削除
 * 7. Logger.log で削除ログ出力 & break
 *
 * 使用シーン:
 * - AI取込一覧画面で「受注入力」ボタンクリック
 * - 受注入力画面で内容確認・修正
 * - 「受注する」ボタンで受注完了
 * - createOrder() 内で本関数を自動呼び出し → 仮受注データ削除
 *
 * @param {string} tempOrderId - 仮受注ID（AI取込一覧から引き継がれる）
 *
 * @see createOrder() - 受注登録完了時に本関数を呼び出し
 *
 * 呼び出し元: orderCode.js の createOrder() 関数（約2062行目）
 * データソース: LINE Bot用スプレッドシートの '仮受注' シート
 */
function deleteTempOrder(tempOrderId) {
  const ss = SpreadsheetApp.openById(getLineBotSpreadsheetId());
  const sheet = ss.getSheetByName('仮受注');

  if (!sheet) return;

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === tempOrderId) {
      sheet.deleteRow(i + 1);
      Logger.log(`仮受注データ削除: ${tempOrderId}`);
      break;
    }
  }
}

// ============================================
// 既存データ移行スクリプト
// ヤマトCSV/佐川CSVの「お客様管理番号」に受注IDを設定
// ============================================

/**
 * 既存ヤマト/佐川CSVデータに受注IDを紐付ける移行スクリプト
 * @returns {Object} - マッチング結果 {yamato: {success: N, failed: N}, sagawa: {success: N, failed: N}}
 */
function migrateExistingCSVData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 受注シートから全データ取得
  const orderSheet = ss.getSheetByName('受注');
  const orderData = orderSheet.getDataRange().getValues();
  const orderHeaders = orderData[0];

  // 受注シートのカラムインデックス取得
  const orderIdCol = orderHeaders.indexOf('受注ID');
  const shippingDateCol = orderHeaders.indexOf('発送日');
  const shippingToNameCol = orderHeaders.indexOf('発送先名');
  const shippingToTelCol = orderHeaders.indexOf('発送先電話番号');
  const deliveryMethodCol = orderHeaders.indexOf('納品方法');

  // 受注データをMap化（発送日+発送先名+電話番号 → 受注ID）
  const orderMap = new Map();
  for (let i = 1; i < orderData.length; i++) {
    const row = orderData[i];
    const orderId = row[orderIdCol];
    const shippingDate = formatDateKey(row[shippingDateCol]);
    const shippingToName = normalizeString(row[shippingToNameCol]);
    const shippingToTel = normalizeString(row[shippingToTelCol]);
    const deliveryMethod = row[deliveryMethodCol];

    if (!orderId || !shippingDate || !shippingToName || !shippingToTel) continue;

    const key = `${shippingDate}|${shippingToName}|${shippingToTel}`;

    // 同一キーに複数の受注IDがある場合は配列で保持
    if (!orderMap.has(key)) {
      orderMap.set(key, []);
    }
    orderMap.get(key).push({ orderId, deliveryMethod });
  }

  Logger.log(`受注データ: ${orderMap.size}件のユニークキー作成`);

  // ヤマトCSV移行
  const yamatoResult = migrateCSVSheet('ヤマトCSV', orderMap, 'ヤマト');

  // 佐川CSV移行
  const sagawaResult = migrateCSVSheet('佐川CSV', orderMap, '佐川');

  const result = {
    yamato: yamatoResult,
    sagawa: sagawaResult,
    totalSuccess: yamatoResult.success + sagawaResult.success,
    totalFailed: yamatoResult.failed + sagawaResult.failed
  };

  Logger.log('=== 移行完了 ===');
  Logger.log(`ヤマトCSV: 成功=${yamatoResult.success}件, 失敗=${yamatoResult.failed}件`);
  Logger.log(`佐川CSV: 成功=${sagawaResult.success}件, 失敗=${sagawaResult.failed}件`);
  Logger.log(`合計: 成功=${result.totalSuccess}件, 失敗=${result.totalFailed}件`);

  return result;
}

/**
 * 個別CSVシートの移行処理
 * @param {string} sheetName - シート名（ヤマトCSV or 佐川CSV）
 * @param {Map} orderMap - 受注データのMap
 * @param {string} deliveryMethod - 納品方法（ヤマト or 佐川）
 * @returns {Object} - {success: N, failed: N, details: [...]}
 */
function migrateCSVSheet(sheetName, orderMap, deliveryMethod) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    Logger.log(`${sheetName}シートが見つかりません`);
    return { success: 0, failed: 0, details: [] };
  }

  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  // CSVシートのカラムインデックス取得
  const shippingDateCol = headers.indexOf('発送日');
  const customerMgmtNoCol = headers.indexOf('お客様管理番号');

  // ヤマト: お届け先名, お届け先電話番号
  // 佐川: お届け先名称１, お届け先電話番号
  const shippingToNameCol = sheetName === 'ヤマトCSV'
    ? headers.indexOf('お届け先名')
    : headers.indexOf('お届け先名称１');
  const shippingToTelCol = headers.indexOf('お届け先電話番号');

  let successCount = 0;
  let failedCount = 0;
  const failedDetails = [];

  // データ行をループ（ヘッダー行をスキップ）
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const currentCustomerMgmtNo = row[customerMgmtNoCol];

    // 既にお客様管理番号が設定されている場合はスキップ
    if (currentCustomerMgmtNo && currentCustomerMgmtNo !== '') {
      continue;
    }

    const shippingDate = formatDateKey(row[shippingDateCol]);
    const shippingToName = normalizeString(row[shippingToNameCol]);
    const shippingToTel = normalizeString(row[shippingToTelCol]);

    if (!shippingDate || !shippingToName || !shippingToTel) {
      failedCount++;
      failedDetails.push({
        row: i + 1,
        reason: '必須項目が空',
        data: { shippingDate, shippingToName, shippingToTel }
      });
      continue;
    }

    const key = `${shippingDate}|${shippingToName}|${shippingToTel}`;
    const matchedOrders = orderMap.get(key);

    if (!matchedOrders || matchedOrders.length === 0) {
      failedCount++;
      failedDetails.push({
        row: i + 1,
        reason: '受注データに一致なし',
        key: key
      });
      continue;
    }

    // 納品方法でフィルタリング
    const matchedOrder = matchedOrders.find(o => o.deliveryMethod === deliveryMethod);

    if (!matchedOrder) {
      failedCount++;
      failedDetails.push({
        row: i + 1,
        reason: `納品方法不一致（期待: ${deliveryMethod}）`,
        key: key
      });
      continue;
    }

    // お客様管理番号に受注IDを設定
    sheet.getRange(i + 1, customerMgmtNoCol + 1).setValue(matchedOrder.orderId);
    successCount++;
    Logger.log(`${sheetName} 行${i + 1}: お客様管理番号=${matchedOrder.orderId} 設定完了`);
  }

  Logger.log(`${sheetName}移行完了: 成功=${successCount}件, 失敗=${failedCount}件`);

  if (failedCount > 0) {
    Logger.log(`${sheetName}失敗詳細:`);
    failedDetails.forEach(detail => {
      Logger.log(`  行${detail.row}: ${detail.reason} - ${JSON.stringify(detail.key || detail.data)}`);
    });
  }

  return { success: successCount, failed: failedCount, details: failedDetails };
}

/**
 * 既存の「ヤマトCSV」シートの住所を16文字制限に合わせて分割するパッチ。
 * お届け先住所・ご依頼主住所が16文字超の場合、先頭16文字をそのまま残し、
 * 17文字目以降を「お届け先住所（アパートマンション名）」または「ご依頼主住所（アパートマンション名）」に設定する。
 * スクリプトエディタで実行するか、メニューから実行してください。
 * @returns {Object} - { updatedRows: number, message: string }
 */
function patchYamatoCsvAddressColumns() {
  const ADDRESS_MAX_LEN = 16;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('ヤマトCSV');

  if (!sheet) {
    Logger.log('ヤマトCSVシートが見つかりません');
    return { updatedRows: 0, message: 'ヤマトCSVシートが見つかりません' };
  }

  const data = sheet.getDataRange().getValues();
  if (!data.length) {
    Logger.log('ヤマトCSVシートにデータがありません');
    return { updatedRows: 0, message: 'データがありません' };
  }

  const headers = data[0];
  const colToAddr = headers.indexOf('お届け先住所');
  const colToApt = headers.indexOf('お届け先住所（アパートマンション名）');
  const colFromAddr = headers.indexOf('ご依頼主住所');
  const colFromApt = headers.indexOf('ご依頼主住所（アパートマンション名）');

  if (colToAddr === -1 || colToApt === -1 || colFromAddr === -1 || colFromApt === -1) {
    Logger.log('必要なヘッダーが見つかりません: お届け先住所, お届け先住所（アパートマンション名）, ご依頼主住所, ご依頼主住所（アパートマンション名）');
    return { updatedRows: 0, message: '必要なヘッダーが見つかりません' };
  }

  let updatedRows = 0;
  for (let r = 1; r < data.length; r++) {
    const row = data[r];
    let changed = false;

    let toAddr = (row[colToAddr] != null && row[colToAddr] !== '') ? String(row[colToAddr]).trim() : '';
    if (toAddr.length > ADDRESS_MAX_LEN) {
      row[colToAddr] = toAddr.substring(0, ADDRESS_MAX_LEN);
      row[colToApt] = toAddr.substring(ADDRESS_MAX_LEN);
      changed = true;
    }

    let fromAddr = (row[colFromAddr] != null && row[colFromAddr] !== '') ? String(row[colFromAddr]).trim() : '';
    if (fromAddr.length > ADDRESS_MAX_LEN) {
      row[colFromAddr] = fromAddr.substring(0, ADDRESS_MAX_LEN);
      row[colFromApt] = fromAddr.substring(ADDRESS_MAX_LEN);
      changed = true;
    }

    if (changed) updatedRows++;
  }

  if (updatedRows > 0) {
    sheet.getRange(1, 1, data.length, headers.length).setValues(data);
  }

  const message = updatedRows > 0
    ? 'ヤマトCSV: ' + updatedRows + '行の住所を16文字で分割して更新しました。'
    : 'ヤマトCSV: 16文字を超える住所はありませんでした。';
  Logger.log(message);
  return { updatedRows: updatedRows, message: message };
}

/**
 * 日付を統一フォーマットのキーに変換（yyyy/MM/dd）
 */
function formatDateKey(dateValue) {
  if (!dateValue) return '';

  try {
    let date;
    if (dateValue instanceof Date) {
      date = dateValue;
    } else if (typeof dateValue === 'string') {
      date = new Date(dateValue.replace(/\//g, '-'));
    } else {
      return '';
    }

    if (isNaN(date.getTime())) return '';

    return Utilities.formatDate(date, 'JST', 'yyyy/MM/dd');
  } catch (e) {
    return '';
  }
}

/**
 * 文字列を正規化（空白除去、全角→半角変換）
 */
function normalizeString(str) {
  if (!str) return '';

  return String(str)
    .replace(/\s+/g, '')  // 全ての空白を除去
    .replace(/[Ａ-Ｚａ-ｚ０-９]/g, function (s) {
      // 全角英数字を半角に変換
      return String.fromCharCode(s.charCodeAt(0) - 0xFEE0);
    })
    .replace(/[‐－―−]/g, '-')  // 各種ハイフンを統一
    .toLowerCase();  // 小文字に統一
}
