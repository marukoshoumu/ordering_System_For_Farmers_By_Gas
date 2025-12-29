/**
 * Google Drive監視トリガー Bot
 *
 * 指定フォルダにPDF/画像が追加されたら自動でAI解析して仮受注シートに登録
 */

// ========== 設定 ==========
// 機密情報は GAS の ScriptProperties で管理してください

// 監視対象のGoogle DriveフォルダID
const WATCH_FOLDER_ID = PropertiesService.getScriptProperties().getProperty('WATCH_FOLDER_ID') || 'YOUR_FOLDER_ID_HERE';

// LINE Botプロジェクトのスプレッドシート
const LINE_BOT_SPREADSHEET_ID = PropertiesService.getScriptProperties().getProperty('LINE_BOT_SPREADSHEET_ID') || '';

// Gemini API設定
const GEMINI_API_KEY = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
const GEMINI_VISION_API_URL = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent';

// ========== トリガー設定用関数 ==========

/**
 * 時間ベーストリガーを設定（手動実行）
 * 5分ごとに checkNewFiles を実行
 */
function setupDriveTrigger() {
  // 既存のトリガーを削除
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'checkNewFiles') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // 新しいトリガーを作成（5分ごと）
  ScriptApp.newTrigger('checkNewFiles')
    .timeBased()
    .everyMinutes(5)
    .create();

  Logger.log('トリガーを設定しました: checkNewFiles を5分ごとに実行');
}

/**
 * トリガーを削除（手動実行）
 */
function removeDriveTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'checkNewFiles') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  Logger.log('トリガーを削除しました');
}

// ========== メイン処理 ==========

/**
 * 新規ファイルをチェックして処理
 * トリガーから5分ごとに実行される
 */
function checkNewFiles() {
  try {
    const folder = DriveApp.getFolderById(WATCH_FOLDER_ID);
    const lastCheckTime = getLastCheckTime();
    const now = new Date();

    // 最後のチェック以降に追加されたファイルを取得
    const files = folder.getFiles();
    let processedCount = 0;

    while (files.hasNext()) {
      const file = files.next();
      const createdTime = file.getDateCreated();

      // 最後のチェック以降に作成されたファイルのみ処理
      if (createdTime > lastCheckTime) {
        const mimeType = file.getMimeType();

        // PDF or 画像ファイルのみ処理
        if (isSupportedFileType(mimeType)) {
          Logger.log(`新規ファイル検出: ${file.getName()}`);
          processFile(file);
          processedCount++;
        }
      }
    }

    // 最終チェック時刻を更新
    saveLastCheckTime(now);

    if (processedCount > 0) {
      Logger.log(`${processedCount}件のファイルを処理しました`);
    }

  } catch (error) {
    Logger.log('エラー: ' + error.toString());
  }
}

/**
 * ファイルを処理（AI解析 → 仮受注登録）
 */
function processFile(file) {
  try {
    const fileName = file.getName();
    const mimeType = file.getMimeType();

    // ファイルをBase64エンコード
    const blob = file.getBlob();
    const base64 = Utilities.base64Encode(blob.getBytes());

    // AI解析実行
    const analysisResult = analyzeFileWithGemini(base64, mimeType, fileName);

    if (!analysisResult || !analysisResult.success) {
      Logger.log(`AI解析失敗: ${fileName}`);
      return;
    }

    // 仮受注シートに登録
    registerTempOrder(analysisResult, fileName, file.getUrl());

    Logger.log(`処理完了: ${fileName}`);

  } catch (error) {
    Logger.log(`ファイル処理エラー(${file.getName()}): ${error.toString()}`);
  }
}

/**
 * Gemini AIでファイルを解析
 */
function analyzeFileWithGemini(base64, mimeType, fileName) {
  try {
    const prompt = `
以下のFAX/発注書画像から受注情報を抽出してJSONで返してください。

抽出項目:
- 顧客情報（会社名、担当者名、電話番号、FAX番号、郵便番号、住所）
- 発送先情報（会社名、担当者名、電話番号、郵便番号、住所）
- 商品リスト（商品名、数量、単位、単価）
- 日付（発送日、納品日、希望配達時間）
- 備考・メモ

JSON形式:
{
  "customer": {"companyName": "", "personName": "", "tel": "", "fax": "", "zipcode": "", "address": ""},
  "shippingTo": {"companyName": "", "personName": "", "tel": "", "zipcode": "", "address": ""},
  "items": [{"productName": "", "quantity": 0, "unit": "", "price": 0}],
  "shippingDate": "",
  "deliveryDate": "",
  "deliveryTime": "",
  "memo": ""
}
`;

    const payload = {
      contents: [{
        parts: [
          { text: prompt },
          {
            inline_data: {
              mime_type: mimeType,
              data: base64
            }
          }
        ]
      }]
    };

    const options = {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(
      `${GEMINI_VISION_API_URL}?key=${GEMINI_API_KEY}`,
      options
    );

    const result = JSON.parse(response.getContentText());

    if (!result.candidates || !result.candidates[0]) {
      Logger.log('AI解析レスポンスが不正: ' + JSON.stringify(result));
      return { success: false };
    }

    const text = result.candidates[0].content.parts[0].text;

    // JSON部分を抽出（マークダウンコードブロックを除去）
    let jsonText = text.replace(/```json\n?/g, '').replace(/```\n?/g, '').trim();
    const analysisData = JSON.parse(jsonText);

    return {
      success: true,
      data: analysisData,
      rawResponse: text
    };

  } catch (error) {
    Logger.log('Gemini API エラー: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

/**
 * 仮受注シートに登録
 */
function registerTempOrder(analysisResult, fileName, fileUrl) {
  try {
    const ss = SpreadsheetApp.openById(LINE_BOT_SPREADSHEET_ID);
    const sheet = ss.getSheetByName('仮受注');

    if (!sheet) {
      Logger.log('仮受注シートが見つかりません');
      return;
    }

    const data = analysisResult.data;
    const tempOrderId = 'TMP' + Utilities.formatDate(new Date(), 'JST', 'yyyyMMddHHmmss');

    // 商品データをJSON化
    const itemsJson = JSON.stringify(data.items || []);

    // 顧客名抽出
    const customerName = data.customer?.companyName || '';

    // 発送先名抽出
    const shippingToName = data.shippingTo?.companyName || '';

    // 解析結果全体をJSON化
    const analysisJson = JSON.stringify(analysisResult);

    // 新しい行を追加
    sheet.appendRow([
      tempOrderId,                    // A: 仮受注ID
      new Date(),                     // B: 登録日時
      '未処理',                       // C: ステータス
      customerName,                   // D: 顧客名
      shippingToName,                 // E: 発送先
      itemsJson,                      // F: 商品データ
      fileName,                       // G: 原文（ファイル名）
      analysisJson,                   // H: 解析結果
      fileUrl,                        // I: ファイルURL
      '',                             // J: 処理日時
    ]);

    Logger.log(`仮受注登録完了: ${tempOrderId}`);

  } catch (error) {
    Logger.log('仮受注登録エラー: ' + error.toString());
  }
}

// ========== ユーティリティ関数 ==========

/**
 * 対応ファイルタイプか確認
 */
function isSupportedFileType(mimeType) {
  const supported = [
    'application/pdf',
    'image/png',
    'image/jpeg',
    'image/jpg',
    'image/gif',
    'image/webp'
  ];
  return supported.includes(mimeType);
}

/**
 * 最終チェック時刻を取得
 */
function getLastCheckTime() {
  const props = PropertiesService.getScriptProperties();
  const lastCheck = props.getProperty('LAST_CHECK_TIME');

  if (lastCheck) {
    return new Date(lastCheck);
  } else {
    // 初回は現在時刻の10分前
    return new Date(Date.now() - 10 * 60 * 1000);
  }
}

/**
 * 最終チェック時刻を保存
 */
function saveLastCheckTime(time) {
  const props = PropertiesService.getScriptProperties();
  props.setProperty('LAST_CHECK_TIME', time.toISOString());
}

/**
 * 手動テスト用: 指定ファイルを処理
 */
function testProcessFile(fileId) {
  const file = DriveApp.getFileById(fileId);
  processFile(file);
}

// ========== データ移行用関数 ==========

/**
 * 既存の「未確定」ステータスを「未処理」に一括変換
 * ※初回セットアップ時に1回だけ実行してください
 */
function migrateStatusToUnprocessed() {
  try {
    const ss = SpreadsheetApp.openById(LINE_BOT_SPREADSHEET_ID);
    const sheet = ss.getSheetByName('仮受注');

    if (!sheet) {
      Logger.log('仮受注シートが見つかりません');
      return;
    }

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const idxStatus = headers.indexOf('ステータス');

    if (idxStatus === -1) {
      Logger.log('ステータス列が見つかりません');
      return;
    }

    let updateCount = 0;

    // データ行をチェック（ヘッダー行をスキップ）
    for (let i = 1; i < data.length; i++) {
      const currentStatus = data[i][idxStatus];

      // 「未確定」または空白を「未処理」に変換
      if (currentStatus === '未確定' || currentStatus === '' || !currentStatus) {
        sheet.getRange(i + 1, idxStatus + 1).setValue('未処理');
        updateCount++;
      }
    }

    Logger.log(`ステータス移行完了: ${updateCount}件を「未処理」に更新しました`);
    return { success: true, count: updateCount };

  } catch (error) {
    Logger.log('ステータス移行エラー: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}
