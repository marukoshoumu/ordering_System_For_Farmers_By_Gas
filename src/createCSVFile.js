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
  if (createCsv('ヤマトCSV', target, yamatoFileName, 1)) {
    record['yamato'] = yamatoFileName;
  }
  if (createCsv('佐川CSV', target, sagawaFileName, 2)) {
    record['sagawa'] = sagawaFileName;
  }
  if (Object.keys(record).length) {
    return JSON.stringify(record);
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
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  const values = sheet.getDataRange().getValues();
  let csv = '';
  var targetDate = Utilities.formatDate(new Date(dateVal), 'JST', 'yyyy/MM/dd');

  for (var value of values) {
    if (value[0] == targetDate) {
      var val = value.slice(1);
      csv += val.join(',') + "\r\n";
    }
  }
  if (csv.length != 0) {
    const blob = createBlob(csv, fileName);
    writeDrive(blob, type);
    return true;
  }
  else {
    return false;
  }
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