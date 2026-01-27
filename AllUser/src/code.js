/**
 * AllUser版 - 発送件数確認・スキャン・出荷チェック機能
 * URLのみで共有可能な限定公開版
 *
 * 共通ロジックは sharedLib.js を使用
 * デプロイ前に sync-shared.sh で sharedLib.js をコピーすること
 *
 * 【必要なスクリプトプロパティ】
 * - MASTER_SPREADSHEET_ID: 受注管理スプレッドシートのID
 * - VISION_API_KEY: OCR機能用（任意）
 */

/**
 * Webアプリのエントリーポイント
 */
function doGet(e) {
  const template = HtmlService.createTemplateFromFile('index');
  template.deployURL = ScriptApp.getService().getUrl();
  const htmlOutput = template.evaluate();
  htmlOutput.addMetaTag('viewport', 'width=device-width, initial-scale=1');
  htmlOutput.setTitle('発送件数確認');
  return htmlOutput;
}

/**
 * スプレッドシートを取得（スクリプトプロパティから）
 */
function getSpreadsheet() {
  const spreadsheetId = PropertiesService.getScriptProperties().getProperty('MASTER_SPREADSHEET_ID');
  if (!spreadsheetId) {
    throw new Error('MASTER_SPREADSHEET_ID が未設定です。スクリプトプロパティに設定してください。');
  }
  return SpreadsheetApp.openById(spreadsheetId);
}

// ============================================
// 発送件数データ取得（sharedLib.jsを使用）
// ============================================

/**
 * 発送予定日ごとの受注サマリーデータを取得
 */
function selectShippingDate() {
  const ss = getSpreadsheet();
  const dates = calculateDatesShared();
  const dateStrings = createDateStringsShared(dates);
  const dataStructure = initializeDataStructureShared();

  // 商品単位マップの作成
  const unitMap = getProductUnitMapShared(ss);

  // 必要な日付範囲のみ取得
  const items = getOrdersByDateRangeShared(ss, dates.yesterday, dates.dayAfter3);
  processItemsShared(items, dates, dataStructure, unitMap);

  return JSON.stringify(buildResultShared(dateStrings, dataStructure));
}

// ============================================
// 出荷済ステータス更新（sharedLib.jsを使用）
// ============================================

/**
 * 受注の出荷済ステータスを更新
 */
function updateOrderShippedStatus(orderId, shippedValue) {
  const ss = getSpreadsheet();
  return updateOrderShippedStatusShared(ss, orderId, shippedValue);
}

// ============================================
// 追跡番号更新（sharedLib.jsを使用）
// ============================================

/**
 * 受注に追跡番号を紐付けし、出荷済に更新
 */
function updateOrderTrackingStatus(orderId, trackingNumber) {
  const ss = getSpreadsheet();
  return updateOrderTrackingStatusShared(ss, orderId, trackingNumber);
}

// ============================================
// OCR機能（sharedLib.jsを使用）
// ============================================

/**
 * Vision APIキーを取得
 */
function getVisionApiKey() {
  return getVisionApiKeyShared();
}

/**
 * 画像から追跡番号を認識（OCR）
 */
function recognizeTrackingNumber(base64Data) {
  return recognizeTrackingNumberShared(base64Data);
}
