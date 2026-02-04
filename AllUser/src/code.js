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

  // 昨日〜明々後日 + 発送日超過（未出荷・非キャンセルで発送日が今日より前）のみ取得（超過は当日分として表示・先頭表示）
  const items = getOrdersByDateRangeShared(ss, dates.yesterday, dates.dayAfter3, true);
  processItemsShared(items, dates, dataStructure, unitMap);

  return JSON.stringify(buildResultShared(dateStrings, dataStructure, dates));
}

// ============================================
// 出荷済ステータス更新（sharedLib.jsを使用）
// ============================================

/**
 * 受注の出荷済ステータスを更新
 * @param {string|number} orderId - 受注ID（非空文字列または正の数）
 * @param {boolean|string} shippedValue - 出荷済の値（booleanまたは'○'/'true'などの文字列）
 * @returns {Object} 更新結果
 * @throws {Error} バリデーションエラー時
 */
function updateOrderShippedStatus(orderId, shippedValue) {
  // orderIdのバリデーション
  if (orderId === null || orderId === undefined || orderId === '') {
    throw new Error('orderIdは必須です。非空文字列または正の数値を指定してください。');
  }

  // orderIdを文字列に正規化（数値の場合は文字列に変換）
  let normalizedOrderId;
  if (typeof orderId === 'number') {
    if (!isFinite(orderId) || orderId <= 0 || !Number.isInteger(orderId)) {
      throw new Error('orderIdは正の整数である必要があります。');
    }
    normalizedOrderId = String(orderId);
  } else if (typeof orderId === 'string') {
    if (orderId.trim() === '') {
      throw new Error('orderIdは空文字列ではいけません。');
    }
    normalizedOrderId = orderId;
  } else {
    throw new Error('orderIdは文字列または数値である必要があります。');
  }

  // shippedValueのバリデーションと正規化
  let normalizedShippedValue;
  if (typeof shippedValue === 'boolean') {
    normalizedShippedValue = shippedValue ? '○' : '';
  } else if (typeof shippedValue === 'string') {
    // 文字列の場合は truthy な値を '○' に、それ以外を '' に変換
    const lowerValue = shippedValue.toLowerCase().trim();
    normalizedShippedValue = (lowerValue === '○' || lowerValue === 'true' || lowerValue === '1' || lowerValue === 'yes') ? '○' : '';
  } else if (shippedValue === null || shippedValue === undefined) {
    normalizedShippedValue = '';
  } else if (typeof shippedValue === 'number') {
    normalizedShippedValue = (shippedValue !== 0 && isFinite(shippedValue)) ? '○' : '';
  } else {
    throw new Error('shippedValueはboolean、文字列、または数値である必要があります。');
  }

  // バリデーション通過後、スプレッドシートを取得して更新
  const ss = getSpreadsheet();
  return updateOrderShippedStatusShared(ss, normalizedOrderId, normalizedShippedValue);
}

/**
 * 注文のステータスを更新（発送前/収穫待ち/出荷済み）
 * @param {string} orderId - 受注ID
 * @param {string} newStatus - 新しいステータス（'発送前'/'収穫待ち'/'出荷済み'）
 * @returns {Object} 実行結果
 */
function updateOrderStatus(orderId, newStatus) {
  const ss = getSpreadsheet();
  return updateOrderStatusShared(ss, orderId, newStatus);
}

// ============================================
// 追跡番号更新（sharedLib.jsを使用）
// ============================================

/**
 * 受注に追跡番号を紐付けし、出荷済に更新
 * @param {string|number} orderId - 受注ID（非空文字列または正の数）
 * @param {string} trackingNumber - 追跡番号（非空文字列）
 * @returns {Object} 更新結果
 * @throws {Error} バリデーションエラー時
 */
function updateOrderTrackingStatus(orderId, trackingNumber) {
  // orderIdのバリデーション
  if (orderId === null || orderId === undefined || orderId === '') {
    throw new Error('orderIdは必須です。非空文字列または正の数値を指定してください。');
  }

  // orderIdを文字列に正規化（数値の場合は文字列に変換）
  let normalizedOrderId;
  if (typeof orderId === 'number') {
    if (!isFinite(orderId) || orderId <= 0 || !Number.isInteger(orderId)) {
      throw new Error('orderIdは正の整数である必要があります。');
    }
    normalizedOrderId = String(orderId);
  } else if (typeof orderId === 'string') {
    if (orderId.trim() === '') {
      throw new Error('orderIdは空文字列ではいけません。');
    }
    normalizedOrderId = orderId;
  } else {
    throw new Error('orderIdは文字列または数値である必要があります。');
  }

  // trackingNumberのバリデーション
  if (trackingNumber === null || trackingNumber === undefined) {
    throw new Error('trackingNumberは必須です。非空文字列を指定してください。');
  }

  if (typeof trackingNumber !== 'string') {
    throw new Error('trackingNumberは文字列である必要があります。');
  }

  if (trackingNumber.trim() === '') {
    throw new Error('trackingNumberは空文字列ではいけません。');
  }

  // バリデーション通過後、スプレッドシートを取得して更新
  const ss = getSpreadsheet();
  return updateOrderTrackingStatusShared(ss, normalizedOrderId, trackingNumber.trim());
}

// ============================================
// OCR機能（sharedLib.jsを使用・サーバー側のみ）
// ============================================

/**
 * 画像から追跡番号を認識（OCR）。サーバー側エンドポイント。
 * クライアントは base64 画像のみ送信し、APIキーはサーバー内で getVisionApiKeyShared により参照される。
 * @param {string} base64Data - 画像のBase64文字列
 * @returns {Object} 抽出結果 {success: boolean, code?: string, message: string}
 */
function recognizeTrackingNumber(base64Data) {
  return recognizeTrackingNumberShared(base64Data);
}
