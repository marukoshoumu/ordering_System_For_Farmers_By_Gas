/**
 * テスト関数 - 送り状CSV作成のテスト実行
 *
 * createShip()関数のテストを実行します。
 * 2024-05-09の発送日でヤマト・佐川CSVを作成します。
 */
function test6() {
  var record = {};
  record['shippingDate'] = '2024-05-09';
  createShip(JSON.stringify(record));
}

/**
 * 送り状CSV作成（ヤマト・佐川の両方を発送日で抽出）
 *
 * 指定された発送日のヤマトCSVと佐川CSVを作成し、Google Driveに保存します。
 * ファイル名にはタイムスタンプが付加されます（yamato_20240509_1430.csv）。
 *
 * 処理フロー:
 * 1. JSON文字列をパースして発送日取得
 * 2. 発送日を yyyy/MM/dd 形式に変換
 * 3. 現在時刻から yyyyMMdd_HHmm 形式のタイムスタンプ生成
 * 4. yamato_TIMESTAMP.csv, sagawa_TIMESTAMP.csv のファイル名生成
 * 5. createCsv('ヤマトCSV', ...) でヤマトCSV作成 → Driveに保存
 * 6. createCsv('佐川CSV', ...) で佐川CSV作成 → Driveに保存
 * 7. 作成されたファイル名をJSONで返却
 *
 * @param {string} datas - JSON文字列 { shippingDate: "2024-05-09" }
 * @returns {string|null} JSON文字列 { yamato: "ファイル名", sagawa: "ファイル名" } または null（該当データなし）
 *
 * 返却例:
 * '{"yamato":"yamato_20240509_1430.csv","sagawa":"sagawa_20240509_1430.csv"}'
 *
 * @see createCsv() - CSV作成とDrive保存
 * @see 'ヤマトCSV'シート - ヤマト運輸送り状データ
 * @see '佐川CSV'シート - 佐川急便送り状データ
 *
 * 呼び出し元: createShippingSlips.html の送り状作成ボタン
 */
function createShip(datas) {
  var data = JSON.parse(datas);
  var target = Utilities.formatDate(new Date(data['shippingDate']), 'JST', 'yyyy/MM/dd');
  var dateNow = Utilities.formatDate(new Date(), 'JST', 'yyyyMMdd_HHmm');
  const yamatoFileName = 'yamato_' + dateNow + '.csv';
  const sagawaFileName = 'sagawa_' + dateNow + '.csv';
  var record = {};
  var yamatoCsv = getCsvContent('ヤマトCSV', target);
  var sagawaCsv = getCsvContent('佐川CSV', target);
  if (yamatoCsv) {
    const blob = createBlob(yamatoCsv, yamatoFileName);
    writeDrive(blob, 1);
    record['yamato'] = yamatoFileName;
  }
  if (sagawaCsv) {
    const blob = createBlob(sagawaCsv, sagawaFileName);
    writeDrive(blob, 2);
    record['sagawa'] = sagawaFileName;
  }
  if (Object.keys(record).length) {
    return JSON.stringify({
      yamato: record['yamato'] || null,
      sagawa: record['sagawa'] || null,
      yamatoCsv: yamatoCsv || null,
      sagawaCsv: sagawaCsv || null,
    });
  }
  else {
    return null;
  }

}
/**
 * CSV作成（指定日のデータを抽出してCSVファイル生成）
 *
 * 指定されたシート名と発送日から該当するレコードを抽出し、
 * CSVファイルを作成してGoogle Driveに保存します。
 *
 * 処理フロー:
 * 1. アクティブスプレッドシート取得
 * 2. 指定シート（'ヤマトCSV' or '佐川CSV'）のデータ全取得
 * 3. 発送日を yyyy/MM/dd 形式に変換
 * 4. 全行をループ:
 *    - A列（発送日）が targetDate と一致する行を抽出
 *    - 1列目（発送日）を除外して、2列目以降をカンマ区切りで結合
 *    - csv文字列に追加（改行コード: \r\n）
 * 5. csv文字列が空でない場合:
 *    - createBlob() でBlobオブジェクト作成
 *    - writeDrive() でDriveに保存
 *    - return true
 * 6. csv文字列が空の場合:
 *    - return false（該当データなし）
 *
 * @param {string} sheetName - シート名（'ヤマトCSV' or '佐川CSV'）
 * @param {string} dateVal - 発送日（Date文字列またはDateオブジェクト）
 * @param {string} fileName - 作成するCSVファイル名
 * @param {number} type - 出力先タイプ（1=ヤマトフォルダ, 2=佐川フォルダ）
 * @returns {boolean} true=CSV作成成功, false=該当データなし
 *
 * @see createBlob() - CSV文字列からBlobオブジェクト作成
 * @see writeDrive() - BlobをGoogle Driveに保存
 */
function createCsv(sheetName, dateVal, fileName, type) {
  const csv = getCsvContent(sheetName, dateVal);
  if (csv) {
    const blob = createBlob(csv, fileName);
    writeDrive(blob, type);
    return true;
  }
  return false;
}

/**
 * CSV BlobオブジェクトDownload Data:image/svg+xml,%3csvg%20xmlns=%27http://www.w3.org/2000/svg%27%20version=%271.1%27%20width=%2730%27%20height=%2730%27/%3e作成
 *
 * CSV文字列からBlobオブジェクトを作成します。
 * UTF-8エンコーディングで text/csv 形式のBlobを生成します。
 *
 * @param {string} csv - CSV文字列（カンマ区切り、改行区切り）
 * @param {string} fileName - ファイル名（例: "yamato_20240509_1430.csv"）
 * @returns {Blob} CSVのBlobオブジェクト
 *
 * @see Utilities.newBlob() - Google Apps Script の Blob生成メソッド
 * @see writeDrive() - 生成したBlobをDriveに保存
 */
function createBlob(csv, fileName) {
  const contentType = 'text/csv';
  const charset = 'utf-8';
  const blob = Utilities.newBlob('', contentType, fileName).setDataFromString(csv, charset);
  return blob;
}

/**
 * Google DriveにCSVファイルを保存（ヤマト/佐川フォルダに振り分け）
 *
 * BlobオブジェクトをGoogle Driveの指定フォルダに保存します。
 * type パラメータによってヤマトフォルダまたは佐川フォルダに振り分けます。
 *
 * 処理フロー:
 * 1. getYamatoFolderId(), getSagawaFolderId() でフォルダID取得
 * 2. DriveApp.getFolderById() でフォルダオブジェクト取得
 * 3. type == 1: ヤマトフォルダに createFile(blob)
 * 4. type != 1: 佐川フォルダに createFile(blob)
 *
 * @param {Blob} blob - CSVのBlobオブジェクト
 * @param {number} type - 出力先タイプ（1=ヤマトフォルダ, 2=佐川フォルダ）
 *
 * @see getYamatoFolderId() - ヤマトフォルダID取得（config.js）
 * @see getSagawaFolderId() - 佐川フォルダID取得（config.js）
 * @see createBlob() - Blob作成
 */
function writeDrive(blob, type) {
  // CSV出力先（必要なフォルダのみ取得）
  if (type == 1) {
    const folderId = getYamatoFolderId();
    if (!folderId) {
      throw new Error('YAMATO_FOLDER_ID がスクリプトプロパティに設定されていません');
    }
    try {
      DriveApp.getFolderById(folderId).createFile(blob);
    } catch (e) {
      throw new Error(`ヤマトCSVフォルダ(ID: ${folderId})へのアクセスに失敗しました: ${e.message}`);
    }
  } else {
    const folderId = getSagawaFolderId();
    if (!folderId) {
      throw new Error('SAGAWA_FOLDER_ID がスクリプトプロパティに設定されていません');
    }
    try {
      DriveApp.getFolderById(folderId).createFile(blob);
    } catch (e) {
      throw new Error(`佐川CSVフォルダ(ID: ${folderId})へのアクセスに失敗しました: ${e.message}`);
    }
  }
}

// ========================================
// CSV バリデーション
// ========================================

/**
 * 発送日が過去日でないか検証
 * @param {string} shippingDate - 発送日文字列
 * @returns {string|null} エラーメッセージ（問題なければ null）
 */
function validateShippingDate(shippingDate) {
  var today = new Date();
  today.setHours(0, 0, 0, 0);
  var shipDate = new Date(shippingDate);
  if (isNaN(shipDate.getTime())) {
    return '発送日が無効です。正しい日付を指定してください。';
  }
  shipDate.setHours(0, 0, 0, 0);
  if (shipDate < today) {
    var formatted = Utilities.formatDate(shipDate, 'JST', 'yyyy/MM/dd');
    return '発送日 ' + formatted + ' は過去日です。本日以降の日付を指定してください。';
  }
  return null;
}

/**
 * 佐川CSVの内容をバリデーション
 *
 * 佐川CSV列配置（0始まり、発送日除外後）:
 *   2: お届け先電話番号, 3: お届け先郵便番号, 4: お届け先住所１,
 *   7: お届け先名称１, 9: お客様管理番号(受注ID),
 *   17: ご依頼主電話番号, 18: ご依頼主郵便番号, 19: ご依頼主住所１, 21: ご依頼主名称１,
 *   24: 品名１, 43: クール便指定, 44: 配達日, 45: 配達指定時間帯, 57: 元着区分
 *
 * @param {string} csvContent - CSV文字列（\r\n区切り）
 * @returns {Array<string>} エラーメッセージ配列（空なら問題なし）
 */
function validateSagawaCsv(csvContent) {
  var errors = [];
  var lines = csvContent.split('\r\n').filter(function(l) { return l.trim().length > 0; });
  var validCool = { '': true, '001': true, '002': true, '003': true };
  var validInvoice = { '': true, '1': true, '2': true };
  var validTime = { '': true, '01': true, '12': true, '14': true, '16': true, '18': true, '04': true, '19': true };

  for (var i = 0; i < lines.length; i++) {
    var cols = parseCsvLine(lines[i]);
    var orderId = (cols[9] || '').trim();
    var label = orderId ? '佐川 [' + orderId + ']: ' : '佐川CSV ' + (i + 1) + '行目: ';

    // お届け先 必須フィールド
    if (!cols[7] || !cols[7].trim()) {
      errors.push(label + 'お届け先名称１が空です');
    }
    if (!cols[4] || !cols[4].trim()) {
      errors.push(label + 'お届け先住所１が空です');
    }
    if (!cols[2] || !cols[2].trim()) {
      errors.push(label + 'お届け先電話番号が空です');
    }
    if (!cols[3] || !cols[3].trim()) {
      errors.push(label + 'お届け先郵便番号が空です');
    }

    // ご依頼主 必須フィールド
    if (!cols[21] || !cols[21].trim()) {
      errors.push(label + 'ご依頼主名称１が空です');
    }
    if (!cols[19] || !cols[19].trim()) {
      errors.push(label + 'ご依頼主住所１が空です');
    }
    if (!cols[17] || !cols[17].trim()) {
      errors.push(label + 'ご依頼主電話番号が空です');
    }

    // 品名
    if (!cols[24] || !cols[24].trim()) {
      errors.push(label + '品名１が空です');
    }

    // クール便指定チェック
    var cool = (cols[43] || '').trim();
    if (!validCool[cool]) {
      errors.push(label + 'クール便指定「' + cool + '」が不正です（001:通常/002:冷蔵/003:冷凍）');
    }

    // 配達日フォーマット（YYYYMMDD、8桁数字）
    var delivDate = (cols[44] || '').trim();
    if (delivDate && !/^\d{8}$/.test(delivDate)) {
      errors.push(label + '配達日「' + delivDate + '」の形式が不正です（YYYYMMDD）');
    }
    if (delivDate && /^\d{8}$/.test(delivDate)) {
      var dd = new Date(delivDate.substring(0, 4) + '-' + delivDate.substring(4, 6) + '-' + delivDate.substring(6, 8));
      var today = new Date();
      today.setHours(0, 0, 0, 0);
      if (dd < today) {
        errors.push(label + '配達日「' + delivDate + '」は過去日です');
      }
    }

    // 配達指定時間帯
    var time = (cols[45] || '').trim();
    if (!validTime[time]) {
      errors.push(label + '配達指定時間帯「' + time + '」が不正です');
    }

    // 元着区分（送り状種別）
    var invoice = (cols[57] || '').trim();
    if (!validInvoice[invoice]) {
      errors.push(label + '元着区分「' + invoice + '」が不正です（1:元払/2:着払）');
    }
  }
  return errors;
}

/**
 * ヤマトCSVの内容をバリデーション（B2クラウド公式レイアウト準拠）
 *
 * ヤマトCSV列配置（0始まり、発送日除外後）:
 *   0: お客様管理番号(受注ID), 1: 送り状種別, 2: クール区分,
 *   4: 出荷予定日, 5: お届け予定日, 6: 配達時間帯,
 *   8: お届け先電話番号, 10: お届け先郵便番号, 11: お届け先住所,
 *   15: お届け先名, 19: ご依頼主電話番号, 21: ご依頼主郵便番号,
 *   22: ご依頼主住所, 24: ご依頼主名, 27: 品名1,
 *   33: コレクト代金引換額, 34: コレクト内消費税額等,
 *   37: 発行枚数, 39: ご請求先顧客コード, 41: 運賃管理番号
 *
 * @param {string} csvContent - CSV文字列（\r\n区切り）
 * @returns {Array<string>} エラーメッセージ配列
 */
function validateYamatoCsv(csvContent) {
  var errors = [];
  var lines = csvContent.split('\r\n').filter(function(l) { return l.trim().length > 0; });
  // B2公式: 0:発払い,2:コレクト,3:クロネコゆうメール,4:タイム,5:着払い,6:発払い(複数口),7:ゆうパケット,8:宅急便コンパクト,9:コンパクトコレクト,A:ネコポス
  var validInvoice = { '0':1, '2':1, '3':1, '4':1, '5':1, '6':1, '7':1, '8':1, '9':1, 'A':1 };
  var validCool = { '':1, '0':1, '1':1, '2':1 };
  var validTime = { '':1, '0812':1, '1416':1, '1618':1, '1820':1, '1921':1, '0010':1, '0017':1 };

  for (var i = 0; i < lines.length; i++) {
    var cols = parseCsvLine(lines[i]);
    var orderId = (cols[0] || '').trim();
    var label = orderId ? 'ヤマト [' + orderId + ']: ' : 'ヤマトCSV ' + (i + 1) + '行目: ';

    // 送り状種別（必須）
    var invoiceType = (cols[1] || '').trim();
    if (!invoiceType) {
      errors.push(label + '送り状種別が空です');
    } else if (!validInvoice[invoiceType]) {
      errors.push(label + '送り状種別「' + invoiceType + '」が不正です（0:発払い/2:コレクト/5:着払い等）');
    }

    // クール区分（発払・コレクト・着払のみ指定可能）
    var cool = (cols[2] || '').trim();
    if (!validCool[cool]) {
      errors.push(label + 'クール区分「' + cool + '」が不正です（0:通常/1:冷凍/2:冷蔵）');
    }

    // 出荷予定日（必須、YYYY/MM/DD形式）
    var shipDate = (cols[4] || '').trim();
    if (!shipDate) {
      errors.push(label + '出荷予定日が空です');
    } else {
      var sd = new Date(shipDate);
      var today = new Date();
      today.setHours(0, 0, 0, 0);
      if (isNaN(sd.getTime())) {
        errors.push(label + '出荷予定日「' + shipDate + '」の形式が不正です（YYYY/MM/DD）');
      } else if (sd < today) {
        errors.push(label + '出荷予定日「' + shipDate + '」は過去日です');
      }
    }

    // お届け予定日（任意だが、設定時は過去日チェック）
    var delivDate = (cols[5] || '').trim();
    if (delivDate) {
      var dd = new Date(delivDate);
      var today2 = new Date();
      today2.setHours(0, 0, 0, 0);
      if (isNaN(dd.getTime())) {
        errors.push(label + 'お届け予定日「' + delivDate + '」の形式が不正です（YYYY/MM/DD）');
      } else if (dd < today2) {
        errors.push(label + 'お届け予定日「' + delivDate + '」は過去日です');
      }
    }

    // 配達時間帯
    var time = (cols[6] || '').trim();
    if (!validTime[time]) {
      errors.push(label + '配達時間帯「' + time + '」が不正です（0812/1416/1618/1820/1921）');
    }

    // お届け先 必須フィールド
    if (!cols[8] || !cols[8].trim()) {
      errors.push(label + 'お届け先電話番号が空です');
    }
    if (!cols[10] || !cols[10].trim()) {
      errors.push(label + 'お届け先郵便番号が空です');
    }
    if (!cols[11] || !cols[11].trim()) {
      errors.push(label + 'お届け先住所が空です');
    }
    if (!cols[15] || !cols[15].trim()) {
      errors.push(label + 'お届け先名が空です');
    }

    // ご依頼主 必須フィールド
    if (!cols[19] || !cols[19].trim()) {
      errors.push(label + 'ご依頼主電話番号が空です');
    }
    if (!cols[21] || !cols[21].trim()) {
      errors.push(label + 'ご依頼主郵便番号が空です');
    }
    if (!cols[22] || !cols[22].trim()) {
      errors.push(label + 'ご依頼主住所が空です');
    }
    if (!cols[24] || !cols[24].trim()) {
      errors.push(label + 'ご依頼主名が空です');
    }

    // 品名1（必須）
    if (!cols[27] || !cols[27].trim()) {
      errors.push(label + '品名１が空です');
    }

    // コレクトの場合、代金引換額が必須
    if (invoiceType === '2') {
      if (!cols[33] || !cols[33].trim()) {
        errors.push(label + 'コレクト代金引換額が空です（送り状種別がコレクトの場合は必須）');
      }
    }

    // ご請求先顧客コード（着払い・クロネコゆうメール以外は必須）
    if (invoiceType !== '5' && invoiceType !== '3') {
      if (!cols[39] || !cols[39].trim()) {
        errors.push(label + 'ご請求先顧客コードが空です');
      }
    }

    // 運賃管理番号（必須）
    if (!cols[41] || !cols[41].trim()) {
      errors.push(label + '運賃管理番号が空です');
    }
  }
  return errors;
}

// ========================================
// 配送伝票ワーカー連携
// ========================================

/**
 * 送り状CSV作成 + 自動印刷（統合関数）
 *
 * 処理フロー:
 * 1. createShip() でCSVをDriveに保存
 * 2. sendToWorker() でワーカーに送信 → Playwright自動化 → PDF取得 → Drive保存
 * 3. DriveにPDFが保存されると、Google Drive for Desktop経由で
 *    ローカルMacのprint-agent（chokidar + lp）が自動検出・印刷する
 *
 * @param {string} datas - JSON文字列 { shippingDate: "2024-05-09" }
 * @returns {string|null} JSON文字列 { yamato, sagawa, workerResults }
 *
 * @see createShip() - CSV作成とDrive保存
 * @see sendToWorker() - ワーカーへのWebhook送信
 *
 * 呼び出し元: createShippingSlips.html の送り状作成＋自動印刷ボタン
 */
function createShipAndPrint(datas) {
  const data = JSON.parse(datas);
  const shippingDate = data['shippingDate'];

  // 0. 発送日バリデーション（過去日チェック）
  var dateError = validateShippingDate(shippingDate);
  if (dateError) {
    writeSlipLog(shippingDate, '共通', '', 'validation_error', dateError);
    return JSON.stringify({ validationErrors: [dateError] });
  }

  // 1. 既存の CSV 作成処理を実行
  const result = createShip(datas);
  if (!result) {
    return null;
  }

  const record = JSON.parse(result);

  // 2. CSVコンテンツのバリデーション
  var validationErrors = [];

  if (record['yamatoCsv']) {
    var yamatoErrors = validateYamatoCsv(record['yamatoCsv']);
    validationErrors = validationErrors.concat(yamatoErrors);
  }
  if (record['sagawaCsv']) {
    var sagawaErrors = validateSagawaCsv(record['sagawaCsv']);
    validationErrors = validationErrors.concat(sagawaErrors);
  }

  if (validationErrors.length > 0) {
    record['validationErrors'] = validationErrors;
    var errorDetail = validationErrors.join('\n');
    Logger.log('バリデーションエラー: ' + errorDetail);
    // キャリアごとにログ記録
    if (record['yamato']) {
      var yErrors = validationErrors.filter(function(e) { return e.indexOf('ヤマト') === 0; });
      if (yErrors.length > 0) writeSlipLog(shippingDate, 'ヤマト', '', 'validation_error', yErrors.join('\n'));
    }
    if (record['sagawa']) {
      var sErrors = validationErrors.filter(function(e) { return e.indexOf('佐川') === 0; });
      if (sErrors.length > 0) writeSlipLog(shippingDate, '佐川', '', 'validation_error', sErrors.join('\n'));
    }
    return JSON.stringify(record);
  }

  // 3. ワーカー用ジョブをキューに登録（ローカル Worker がポーリングで取得）
  const workerResults = {};

  // 4. ヤマトCSVがある場合、ジョブをキューに登録
  if (record['yamato']) {
    try {
      const wResult = queueSlipJob('yamato', record['yamatoCsv'], shippingDate);
      workerResults['yamato'] = wResult;
    } catch (e) {
      Logger.log('ヤマトジョブ登録エラー: ' + e.message);
      workerResults['yamato'] = { success: false, message: e.message };
      writeSlipLog(shippingDate, 'ヤマト', '', 'error', e.message);
    }
  }

  // 5. 佐川CSVがある場合、ジョブをキューに登録
  if (record['sagawa']) {
    try {
      const wResult = queueSlipJob('sagawa', record['sagawaCsv'], shippingDate);
      workerResults['sagawa'] = wResult;
    } catch (e) {
      Logger.log('佐川ジョブ登録エラー: ' + e.message);
      workerResults['sagawa'] = { success: false, message: e.message };
      writeSlipLog(shippingDate, '佐川', '', 'error', e.message);
    }
  }

  record['workerResults'] = workerResults;
  return JSON.stringify(record);
}

// sendSlipToPrinter() は削除済み
// Brother MFC-J990DN にはメール受信機能がないため、
// 印刷は print-agent（ローカルMac上の chokidar + lp コマンド）が担当する。
// Drive同期フォルダにPDFが保存されると print-agent が自動検出・印刷する。

/**
 * 指定されたシート・発送日のCSVテキスト内容を取得
 *
 * @param {string} sheetName - シート名（'ヤマトCSV' or '佐川CSV'）
 * @param {string} dateVal - 発送日
 * @returns {string|null} CSVテキスト（該当データなしの場合は null）
 */
/**
 * RFC 4180準拠のCSV行パーサー（クォートされたフィールド対応）
 */
function parseCsvLine(line) {
  var fields = [];
  var field = '';
  var inQuotes = false;
  for (var i = 0; i < line.length; i++) {
    var ch = line[i];
    if (inQuotes) {
      if (ch === '"') {
        if (i + 1 < line.length && line[i + 1] === '"') {
          field += '"';
          i++;
        } else {
          inQuotes = false;
        }
      } else {
        field += ch;
      }
    } else {
      if (ch === '"') {
        inQuotes = true;
      } else if (ch === ',') {
        fields.push(field);
        field = '';
      } else {
        field += ch;
      }
    }
  }
  fields.push(field);
  return fields;
}

function escapeCsvField(field) {
  var str = (field == null) ? '' : String(field);
  if (str.indexOf(',') >= 0 || str.indexOf('"') >= 0 || str.indexOf('\n') >= 0 || str.indexOf('\r') >= 0) {
    return '"' + str.replace(/"/g, '""') + '"';
  }
  return str;
}

function getCsvContent(sheetName, dateVal) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    Logger.log('シートが見つかりません: ' + sheetName);
    return null;
  }
  const values = sheet.getDataRange().getValues();
  let csv = '';
  var targetDate = Utilities.formatDate(new Date(dateVal), 'JST', 'yyyy/MM/dd');

  for (var value of values) {
    if (value[0] == targetDate) {
      var val = value.slice(1);
      csv += val.map(escapeCsvField).join(',') + "\r\n";
    }
  }
  return csv.length > 0 ? csv : null;
}

/**
 * 伝票処理ジョブをキューに登録（ローカル Worker がポーリングで取得）
 *
 * 伝票処理ログに "queued" ステータスで書き込み、
 * CacheService に CSV 内容を保存する。
 *
 * @param {string} carrier - 'yamato' or 'sagawa'
 * @param {string} csvContent - CSV テキスト
 * @param {string} shippingDate - 発送日
 * @returns {{ success: boolean, jobId: string }}
 */
function queueSlipJob(carrier, csvContent, shippingDate) {
  var jobId = Utilities.getUuid();
  var carrierLabel = carrier === 'yamato' ? 'ヤマト' : '佐川';

  // CacheService に CSV 内容を保存（最大6時間、1キー100KB制限）
  var cache = CacheService.getScriptCache();
  var cacheKey = 'slip_csv_' + jobId;
  cache.put(cacheKey, csvContent, 21600);

  // 保存を検証（100KB超過等で静かに失敗する場合がある）
  var verify = cache.get(cacheKey);
  if (!verify) {
    throw new Error('CSVキャッシュ保存に失敗しました（サイズ超過の可能性: ' + csvContent.length + '文字）');
  }

  // CSVの件数をカウント（空行を除く）
  var recordCount = csvContent.split('\r\n').filter(function(l) { return l.trim().length > 0; }).length;

  // 伝票処理ログに queued で書き込み
  writeSlipLog(shippingDate, carrierLabel, jobId, 'queued', '', recordCount);

  Logger.log('ジョブ登録完了: ' + carrier + ' jobId=' + jobId);
  return { success: true, jobId: jobId };
}

/**
 * 伝票処理ログシートに1行追記
 *
 * @param {string} shippingDate - 発送日
 * @param {string} carrier - 業者名（'ヤマト' or '佐川'）
 * @param {string} jobId - ワーカーのジョブID
 * @param {string} status - ステータス ('processing', 'pdf_ready', 'printed', 'error')
 * @param {string} [errorDetail] - エラー詳細（任意）
 */
function writeSlipLog(shippingDate, carrier, jobId, status, errorDetail, recordCount) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('伝票処理ログ');
    if (!sheet) {
      sheet = ss.insertSheet('伝票処理ログ');
      sheet.appendRow(['発送日', '業者', 'jobId', 'ステータス', 'CSV作成時刻', 'PDF完了時刻', '印刷時刻', 'Driveリンク', 'エラー詳細', '件数']);
    }
    const now = new Date();
    sheet.appendRow([shippingDate, carrier, jobId, status, now, '', '', '', errorDetail || '', recordCount || '']);
  } catch (e) {
    Logger.log('ログシート書き込みエラー: ' + e.message);
  }
}