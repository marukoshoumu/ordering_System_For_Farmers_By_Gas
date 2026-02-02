/**
 * スプレッドシート初期化モジュール
 * 
 * 初回セットアップ時に必要なシートとマスタデータを自動作成します。
 * GASエディタから実行: initializeMasterSpreadsheet() または initializeLineBotSpreadsheet()
 */

/**
 * マスタスプレッドシートを初期化
 * 
 * 必要なシートとヘッダー、固定マスタデータを作成します。
 * 既にシートが存在する場合はスキップします。
 * 
 * 実行方法:
 * GASエディタで initializeMasterSpreadsheet() を実行
 * 
 * @returns {Object} 実行結果 { success: boolean, message: string, createdSheets: Array<string> }
 */
function initializeMasterSpreadsheet() {
  try {
    const masterSpreadsheetId = getMasterSpreadsheetId();
    if (!masterSpreadsheetId) {
      return {
        success: false,
        message: 'MASTER_SPREADSHEET_ID が設定されていません。スクリプトプロパティを確認してください。'
      };
    }

    const ss = SpreadsheetApp.openById(masterSpreadsheetId);
    const createdSheets = [];

    Logger.log('マスタスプレッドシートの初期化を開始します...');

    // 1. トランザクション・CSV出力シートの作成
    createSheetIfNotExists(ss, '受注', getOrderSheetHeadersForInit(), createdSheets);
    createSheetIfNotExists(ss, '定期受注', getRecurringOrderHeaders(), createdSheets);
    createSheetIfNotExists(ss, '見積書データ', getQuotationDataHeaders(), createdSheets);
    createSheetIfNotExists(ss, 'ヤマトCSV', getYamatoCsvHeaders(), createdSheets);
    createSheetIfNotExists(ss, '佐川CSV', getSagawaCsvHeaders(), createdSheets);
    createSheetIfNotExists(ss, '発送先マッピング', getShippingMappingHeaders(), createdSheets);
    createSheetIfNotExists(ss, 'pass', ['ID', 'Pass'], createdSheets);
    
    // 2. CSV取込用商品マッピングシートの作成
    createSheetIfNotExists(ss, '食べチョク商品', getTabechokuProductHeaders(), createdSheets);
    createSheetIfNotExists(ss, 'ポケマル商品', getPokemaruProductHeaders(), createdSheets);
    createSheetIfNotExists(ss, 'ふるさと納税商品', getFurusatoProductHeaders(), createdSheets);
    
    // 3. その他マスタシートの作成
    createSheetIfNotExists(ss, '担当者', getRecipientHeaders(), createdSheets);
    
    // 4. 会社名シートの作成（COMPANY_DISPLAY_NAMEプロパティの値と同じ名前）
    const companyDisplayName = getCompanyDisplayName();
    createSheetIfNotExists(ss, companyDisplayName, getCompanySheetHeaders(), createdSheets);

    // 5. マスタデータシートの作成
    createSheetIfNotExists(ss, '納品方法', ['納品方法ＩＤ', '納品方法'], createdSheets);
    createSheetIfNotExists(ss, '保存温度帯', ['保存温度帯ＩＤ', '保存温度帯'], createdSheets);
    createSheetIfNotExists(ss, '見積納品方法', ['納品方法ＩＤ', '納品方法'], createdSheets);
    createSheetIfNotExists(ss, '単位', ['単位ＩＤ', '単位'], createdSheets);
    createSheetIfNotExists(ss, '受付方法', ['受付方法ＩＤ', '受付方法'], createdSheets);
    createSheetIfNotExists(ss, '配送時間帯', ['納品方法', '時間指定値', '時間指定'], createdSheets);
    createSheetIfNotExists(ss, '送り状種別', ['納品方法', '種別値', '種別'], createdSheets);
    createSheetIfNotExists(ss, 'クール区分', ['納品方法', '種別値', '種別'], createdSheets);
    createSheetIfNotExists(ss, '荷扱い', ['納品方法', '種別値', '種別'], createdSheets);
    createSheetIfNotExists(ss, '箱サイズ', ['箱サイズ'], createdSheets);

    // 6. 固定マスタデータの投入
    insertMasterDataIfEmpty(ss, '納品方法', getDeliveryMethodData());
    insertMasterDataIfEmpty(ss, '保存温度帯', getPreservationData());
    insertMasterDataIfEmpty(ss, '見積納品方法', getQuotationDeliveryMethodData());
    insertMasterDataIfEmpty(ss, '単位', getUnitData());
    insertMasterDataIfEmpty(ss, '受付方法', getReceiptMethodData());
    insertMasterDataIfEmpty(ss, '配送時間帯', getDeliveryTimeData());
    insertMasterDataIfEmpty(ss, '送り状種別', getInvoiceTypeData());
    insertMasterDataIfEmpty(ss, 'クール区分', getCoolClassData());
    insertMasterDataIfEmpty(ss, '荷扱い', getCargoData());
    insertMasterDataIfEmpty(ss, '箱サイズ', getBoxSizeData());

    Logger.log('マスタスプレッドシートの初期化が完了しました。作成されたシート数: ' + createdSheets.length);

    return {
      success: true,
      message: 'マスタスプレッドシートの初期化が完了しました。',
      createdSheets: createdSheets
    };
  } catch (error) {
    Logger.log('初期化エラー: ' + error.toString());
    return {
      success: false,
      message: '初期化中にエラーが発生しました: ' + error.toString()
    };
  }
}

/**
 * LINE Botスプレッドシートを初期化
 * 
 * 仮受注シートを作成します。
 * 
 * 実行方法:
 * GASエディタで initializeLineBotSpreadsheet() を実行
 * 
 * @returns {Object} 実行結果 { success: boolean, message: string, createdSheets: Array<string> }
 */
function initializeLineBotSpreadsheet() {
  try {
    const lineBotSpreadsheetId = getLineBotSpreadsheetId();
    if (!lineBotSpreadsheetId) {
      return {
        success: false,
        message: 'LINE_BOT_SPREADSHEET_ID が設定されていません。スクリプトプロパティを確認してください。'
      };
    }

    const ss = SpreadsheetApp.openById(lineBotSpreadsheetId);
    const createdSheets = [];

    Logger.log('LINE Botスプレッドシートの初期化を開始します...');

    // 仮受注シートの作成
    createSheetIfNotExists(ss, '仮受注', getTempOrderHeaders(), createdSheets);

    Logger.log('LINE Botスプレッドシートの初期化が完了しました。作成されたシート数: ' + createdSheets.length);

    return {
      success: true,
      message: 'LINE Botスプレッドシートの初期化が完了しました。',
      createdSheets: createdSheets
    };
  } catch (error) {
    Logger.log('初期化エラー: ' + error.toString());
    return {
      success: false,
      message: '初期化中にエラーが発生しました: ' + error.toString()
    };
  }
}

// ========================================
// ヘッダー定義関数
// ========================================

/**
 * 商品マッピングシートのヘッダー定義
 * @returns {Array<string>} ヘッダー配列
 */
function getProductMappingHeaders() {
  return [
    '顧客表記',
    '商品名',
    '商品分類',
    '規格',
    '発送先名',
    '顧客名',
    '登録日',
    '使用回数',
    '最終使用日',
    '登録者',
    '備考'
  ];
}

/**
 * 受注シートのヘッダー定義（初期化用）
 * @returns {Array<string>} ヘッダー配列
 */
function getOrderSheetHeadersForInit() {
  return [
    '受注ID',
    '受注日',
    '発送日',
    '納品日',
    '顧客名',
    '発送先名',
    '発送先郵便番号',
    '発送先住所',
    '発送先電話番号',
    '発送元名',
    '発送元郵便番号',
    '発送元住所',
    '発送元電話番号',
    '受付方法',
    '受付者',
    '納品方法',
    '配達時間帯',
    '納品書',
    '請求書',
    '領収書',
    'パンフ',
    'レシピ',
    'その他添付',
    '品名',
    '送り状種別',
    'クール区分',
    '荷扱い１',
    '荷扱い２',
    '荷扱い３',
    '代引総額',
    '代引内税',
    '発行枚数',
    '社内メモ',
    '送り状備考欄',
    '納品書備考欄',
    'メモ',
    '納品書テキスト',
    '商品データJSON',
    '出荷済',
    '追跡番号',
    '紐付け受注ID',
    '小計'
  ];
}

/**
 * 見積書データシートのヘッダー定義
 * @returns {Array<string>} ヘッダー配列
 */
function getQuotationDataHeaders() {
  return [
    '発注ID',
    '作成日',
    '顧客名',
    '納品先名',
    '納品方法',
    'リードタイム',
    '商品分類',
    '商品名',
    '規格',
    '出荷ロット',
    '価格',
    '単位',
    '産地',
    'JANコード',
    '保存温度帯',
    '賞味期限延長',
    'その他',
    'メモ'
  ];
}

/**
 * ヤマトCSVシートのヘッダー定義
 * @returns {Array<string>} ヘッダー配列
 */
function getYamatoCsvHeaders() {
  return [
    '発送日',
    'お客様管理番号',
    '送り状種別',
    'クール区分',
    '伝票番号',
    '出荷予定日',
    'お届け予定（指定）日',
    '配達時間帯',
    'お届け先コード',
    'お届け先電話番号',
    'お届け先電話番号枝番',
    'お届け先郵便番号',
    'お届け先住所',
    'お届け先住所（アパートマンション名）',
    'お届け先会社・部門名１',
    'お届け先会社・部門名２',
    'お届け先名',
    'お届け先名略称カナ',
    '敬称',
    'ご依頼主コード',
    'ご依頼主電話番号',
    'ご依頼主電話番号枝番',
    'ご依頼主郵便番号',
    'ご依頼主住所',
    'ご依頼主住所（アパートマンション名）',
    'ご依頼主名',
    'ご依頼主略称カナ',
    '品名コード１',
    '品名１',
    '品名コード２',
    '品名２',
    '荷扱い１',
    '荷扱い２',
    '記事',
    'コレクト代金引換額（税込）',
    'コレクト内消費税額等',
    '営業所止置き',
    '営業所コード',
    '発行枚数',
    '個数口枠の印字',
    'ご請求先顧客コード',
    'ご請求先分類コード',
    '運賃管理番号',
    'クロネコwebコレクトデータ登録',
    'クロネコwebコレクト加盟店番号',
    'クロネコwebコレクト申込受付番号１'
  ];
}

/**
 * 佐川CSVシートのヘッダー定義
 * @returns {Array<string>} ヘッダー配列
 */
function getSagawaCsvHeaders() {
  return [
    '発送日',
    'お届け先コード取得区分',
    'お届け先コード',
    'お届け先電話番号',
    'お届け先郵便番号',
    'お届け先住所１',
    'お届け先住所２',
    'お届け先住所３',
    'お届け先名称１',
    'お届け先名称２',
    'お客様管理番号',
    'お客様コード',
    '部署ご担当者コード取得区分',
    '部署ご担当者コード',
    '部署ご担当者名称',
    '荷送人電話番号',
    'ご依頼主コード取得区分',
    'ご依頼主コード',
    'ご依頼主電話番号',
    'ご依頼主郵便番号',
    'ご依頼主住所',
    'ご依頼主名称',
    '品名コード',
    '品名',
    '送り状種別',
    'クール区分',
    '荷扱い１',
    '荷扱い２',
    '荷扱い３',
    '記事',
    'コレクト代金引換額（税込）',
    'コレクト内消費税額等',
    '営業所止置き',
    '営業所コード',
    '発行枚数',
    '個数口枠の印字',
    'ご請求先顧客コード',
    'ご請求先分類コード',
    '運賃管理番号'
  ];
}

/**
 * 発送先マッピングシートのヘッダー定義
 * @returns {Array<string>} ヘッダー配列
 */
function getShippingMappingHeaders() {
  return [
    'AI認識発送先名',
    'マスタ発送先名',
    '顧客名',
    '信頼度',
    '最終使用日',
    '作成日時'
  ];
}

/**
 * 仮受注シートのヘッダー定義
 * @returns {Array<string>} ヘッダー配列
 */
function getTempOrderHeaders() {
  return [
    '仮受注ID',
    '登録日時',
    'ステータス',
    '顧客名',
    '発送先',
    '商品データ',
    '原文',
    '解析結果',
    'ファイルURL',
    '処理日時',
    '通知済み',
    '通知日時'
  ];
}

/**
 * 食べチョク商品シートのヘッダー定義
 * @returns {Array<string>} ヘッダー配列
 */
function getTabechokuProductHeaders() {
  return [
    '商品名',
    '変換値',
    '変換数量'
  ];
}

/**
 * ポケマル商品シートのヘッダー定義
 * @returns {Array<string>} ヘッダー配列
 */
function getPokemaruProductHeaders() {
  return [
    '商品名',
    '分量',
    '変換値',
    '変換数量'
  ];
}

/**
 * ふるさと納税商品シートのヘッダー定義
 * @returns {Array<string>} ヘッダー配列
 */
function getFurusatoProductHeaders() {
  return [
    '商品名',
    '変換値',
    '変換数量'
  ];
}

/**
 * 担当者シートのヘッダー定義
 * @returns {Array<string>} ヘッダー配列
 */
function getRecipientHeaders() {
  return [
    '担当者名',
    '部署',
    'TEL',
    'メールアドレス',
    '備考'
  ];
}

/**
 * 会社名シートのヘッダー定義（COMPANY_DISPLAY_NAMEプロパティの値と同じ名前のシート）
 * @returns {Array<string>} ヘッダー配列
 */
function getCompanySheetHeaders() {
  return [
    '名前',
    '電話番号',
    '郵便番号',
    '住所'
  ];
}

// ========================================
// 固定マスタデータ定義
// ========================================

/**
 * 納品方法の固定データ
 * @returns {Array<Array<string>>} データ行の配列 [ID, 納品方法]
 */
function getDeliveryMethodData() {
  return [
    ['1', 'ヤマト'],
    ['2', '佐川'],
    ['3', '西濃運輸'],
    ['4', '配達'],
    ['5', '店舗受取'],
    ['6', 'ヤマト伝票受領'],
    ['9', '佐川伝票受領'],
    ['7', 'レターパック'],
    ['8', 'その他']
  ];
}

/**
 * 保存温度帯の固定データ
 * @returns {Array<Array<string>>} データ行の配列 [ID, 保存温度帯]
 */
function getPreservationData() {
  return [
    ['1', '常温'],
    ['2', '冷蔵'],
    ['3', '冷凍']
  ];
}

/**
 * 見積納品方法の固定データ
 * @returns {Array<Array<string>>} データ行の配列 [ID, 納品方法]
 */
function getQuotationDeliveryMethodData() {
  return [
    ['1', 'ヤマト'],
    ['2', '佐川'],
    ['3', '西濃運輸'],
    ['4', '自社配送'],
    ['5', 'その他']
  ];
}

/**
 * 単位の固定データ
 * @returns {Array<Array<string>>} データ行の配列 [ID, 単位]
 */
function getUnitData() {
  return [
    ['1', 'g'],
    ['2', 'kg'],
    ['3', '個'],
    ['4', '箱'],
    ['5', '本'],
    ['6', 'ケース'],
    ['7', '袋']
  ];
}

/**
 * 受付方法の固定データ
 * @returns {Array<Array<string>>} データ行の配列 [ID, 受付方法]
 */
function getReceiptMethodData() {
  return [
    ['1', 'FAX'],
    ['2', 'メール'],
    ['3', '電話'],
    ['4', 'LINE'],
    ['5', 'ふるさと納税'],
    ['6', 'Messenger'],
    ['7', '食べチョク'],
    ['8', 'その他']
  ];
}

/**
 * 配送時間帯の固定データ（納品方法ごと）
 * @returns {Array<Array<string>>} データ行の配列 [納品方法, 時間指定値, 時間指定]
 */
function getDeliveryTimeData() {
  return [
    ['ヤマト', ':指定なし', '指定なし'],
    ['ヤマト', '0812:午前中', '午前中'],
    ['ヤマト', '1416:14～16時', '14～16時'],
    ['ヤマト', '1618:16～18時', '16～18時'],
    ['ヤマト', '1820:18～20時', '18～20時'],
    ['ヤマト', '1921:19～21時', '19～21時'],
    ['佐川', ':指定なし', '指定なし'],
    ['佐川', '01:午前中', '午前中'],
    ['佐川', '12:12:00～14:00', '12:00～14:00'],
    ['佐川', '14:14:00～16:00', '14:00～16:00'],
    ['佐川', '16:16:00～18:00', '16:00～18:00'],
    ['佐川', '18:18:00～20:00', '18:00～20:00'],
    ['佐川', '04:18:00～21:00', '18:00～21:00'],
    ['佐川', '19:19:00～21:00', '19:00～21:00']
  ];
}

/**
 * 送り状種別の固定データ（納品方法ごと）
 * @returns {Array<Array<string>>} データ行の配列 [納品方法, 種別値, 種別]
 */
function getInvoiceTypeData() {
  return [
    ['ヤマト', '0:発払い', '発払い'],
    ['ヤマト', '2:コレクト', 'コレクト'],
    ['ヤマト', '5:着払い', '着払い'],
    ['佐川', '1:元払', '元払'],
    ['佐川', '2:着払', '着払']
  ];
}

/**
 * クール区分の固定データ（納品方法ごと）
 * @returns {Array<Array<string>>} データ行の配列 [納品方法, 種別値, 種別]
 */
function getCoolClassData() {
  return [
    ['ヤマト', '0:通常', '通常'],
    ['ヤマト', '1:冷凍', '冷凍'],
    ['ヤマト', '2:冷蔵', '冷蔵'],
    ['佐川', '001:通常', '通常'],
    ['佐川', '002:冷蔵', '冷蔵'],
    ['佐川', '003:冷凍', '冷凍'],
    ['ヤマト伝票受領', '0:通常', '通常'],
    ['ヤマト伝票受領', '1:冷凍', '冷凍'],
    ['ヤマト伝票受領', '2:冷蔵', '冷蔵'],
    ['佐川伝票受領', '001:通常', '通常'],
    ['佐川伝票受領', '002:冷蔵', '冷蔵'],
    ['佐川伝票受領', '003:冷凍', '冷凍']
  ];
}

/**
 * 荷扱いの固定データ（納品方法ごと）
 * @returns {Array<Array<string>>} データ行の配列 [納品方法, 種別値, 種別]
 */
function getCargoData() {
  return [
    ['ヤマト', 'ワレ物注意:ワレ物注意', 'ワレ物注意'],
    ['ヤマト', '下積厳禁:下積厳禁', '下積厳禁'],
    ['ヤマト', '天地無用:天地無用', '天地無用'],
    ['ヤマト', 'ナマモノ:ナマモノ', 'ナマモノ'],
    ['ヤマト', '水漏厳禁:水漏厳禁', '水漏厳禁'],
    ['佐川', '001:冷蔵', '飛脚クール便（冷蔵）'],
    ['佐川', '002:冷凍', '飛脚クール便（冷凍）'],
    ['佐川', '004:営業所受取サービス', '営業所受取サービス'],
    ['佐川', '005:指定日配達サービス', '指定日配達サービス'],
    ['佐川', '008:ｅコレクト（現金）', 'ｅコレクト（現金）'],
    ['佐川', '011:取扱注意', '取扱注意'],
    ['佐川', '012:貴重品', '貴重品'],
    ['佐川', '013:天地無用', '天地無用'],
    ['佐川', '020:時間帯指定サービス（午前中）', '時間帯指定サービス（午前中）'],
    ['佐川', '021:時間帯指定サービス（18時～21時）', '時間帯指定サービス（18時～21時）'],
    ['佐川', '022:時間帯指定サービス（12時～14時）', '時間帯指定サービス（12時～14時）'],
    ['佐川', '023:時間帯指定サービス（14時～16時）', '時間帯指定サービス（14時～16時）'],
    ['佐川', '024:時間帯指定サービス（16時～18時）', '時間帯指定サービス（16時～18時）'],
    ['佐川', '025:時間帯指定サービス（18時～20時）', '時間帯指定サービス（18時～20時）'],
    ['佐川', '026:時間帯指定サービス（19時～21時）', '時間帯指定サービス（19時～21時）']
  ];
}

/**
 * 箱サイズの固定データ
 * @returns {Array<Array<string>>} データ行の配列 [箱サイズ]
 */
function getBoxSizeData() {
  return [
    ['60サイズ'],
    ['80サイズ'],
    ['100サイズ'],
    ['120サイズ'],
    ['140サイズ'],
    ['160サイズ']
  ];
}

// ========================================
// ユーティリティ関数
// ========================================

/**
 * シートが存在しない場合に作成
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss - スプレッドシート
 * @param {string} sheetName - シート名
 * @param {Array<string>} headers - ヘッダー配列
 * @param {Array<string>} createdSheets - 作成されたシート名を追加する配列
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} シートオブジェクト
 */
function createSheetIfNotExists(ss, sheetName, headers, createdSheets) {
  let sheet = ss.getSheetByName(sheetName);
  
  if (sheet) {
    Logger.log('シート「' + sheetName + '」は既に存在します。スキップします。');
    return sheet;
  }

  // 新規シートを作成
  sheet = ss.insertSheet(sheetName);
  
  // ヘッダー行を設定
  if (headers && headers.length > 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    // ヘッダー行のスタイル設定
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#4285f4');
    headerRange.setFontColor('#ffffff');
    headerRange.setFontWeight('bold');
    headerRange.setHorizontalAlignment('center');
    
    // ヘッダー行を固定
    sheet.setFrozenRows(1);
  }
  
  createdSheets.push(sheetName);
  Logger.log('シート「' + sheetName + '」を作成しました。');
  
  return sheet;
}

/**
 * マスタデータが空の場合に初期データを投入
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss - スプレッドシート
 * @param {string} sheetName - シート名
 * @param {Array<Array<string>>} data - データ行の配列（ヘッダー行は含まない）
 */
function insertMasterDataIfEmpty(ss, sheetName, data) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    Logger.log('シート「' + sheetName + '」が見つかりません。スキップします。');
    return;
  }

  // データ行があるか確認（ヘッダー行を除く）
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    Logger.log('シート「' + sheetName + '」には既にデータが存在します。スキップします。');
    return;
  }

  // データを投入
  if (data && data.length > 0) {
    sheet.getRange(2, 1, data.length, data[0].length).setValues(data);
    Logger.log('シート「' + sheetName + '」に初期データを投入しました。行数: ' + data.length);
  }
}
