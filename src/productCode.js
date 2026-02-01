/**
 * スプレッドシートのシート直接リンクURLを取得
 *
 * 指定されたシート名のスプレッドシート直接リンクURLを生成します。
 * スプレッドシートURLに #gid=シートID を付加することで、特定のシートに直接遷移できるURLを作成します。
 *
 * 処理フロー:
 * 1. アクティブスプレッドシート取得
 * 2. 指定されたシート名のシートオブジェクト取得
 * 3. スプレッドシートのURL取得
 * 4. シートID取得
 * 5. URL + "#gid=" + シートID を返却
 *
 * @param {string} sheetName - シート名（例: "受注", "顧客情報", "商品"）
 * @returns {string} シート直接リンクURL（例: "https://docs.google.com/spreadsheets/d/.../edit#gid=123456"）
 *
 * 使用例:
 * const url = getOpenUrl("受注");
 * // 返却: "https://docs.google.com/spreadsheets/d/ABC.../edit#gid=0"
 *
 * @see getOpenUrlDrive() - Drive フォルダURLを取得（ヤマト/佐川/見積書等）
 */
function getOpenUrl(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  var ss_url = ss.getUrl();
  var sh_id = sheet.getSheetId();

  return ss_url + "#gid=" + sh_id;
}

/**
 * Driveフォルダの直接リンクURLを取得（シート名に応じたフォルダ）
 *
 * シート名に応じて、対応するGoogle DriveフォルダのURLを返します。
 * 見積書、請求書、納品書、ヤマトCSV、佐川CSV、freee納品書CSVの各フォルダURLを取得できます。
 *
 * シート名とフォルダの対応:
 * - '見積書' → getQuotationFolderUrl()
 * - '請求書' → getBillFolderUrl()
 * - '納品書' → getDeliveredFolderUrl()
 * - 'ヤマト' → getYamatoFolderUrl()
 * - '佐川' → getSagawaFolderUrl()
 * - 'freee納品書CSV' → getFreeeCSVFolderUrl()
 *
 * @param {string} sheetName - シート名（'見積書', '請求書', '納品書', 'ヤマト', '佐川', 'freee納品書CSV'）
 * @returns {string|undefined} DriveフォルダURL、該当なしの場合はundefined
 *
 * @see getQuotationFolderUrl() - 見積書フォルダURL (config.js)
 * @see getBillFolderUrl() - 請求書フォルダURL (config.js)
 * @see getDeliveredFolderUrl() - 納品書フォルダURL (config.js)
 * @see getYamatoFolderUrl() - ヤマトCSVフォルダURL (config.js)
 * @see getSagawaFolderUrl() - 佐川CSVフォルダURL (config.js)
 * @see getFreeeCSVFolderUrl() - freee納品書CSVフォルダURL (config.js)
 *
 * 使用例:
 * const url = getOpenUrlDrive('freee納品書CSV');
 * // 返却: "https://drive.google.com/drive/folders/ABC..." (getFreeeCSVFolderUrl()の戻り値)
 *
 * const url2 = getOpenUrlDrive('見積書');
 * // 返却: "https://drive.google.com/drive/folders/XYZ..." (getQuotationFolderUrl()の戻り値)
 *
 * 呼び出し元: フロントエンドの「フォルダを開く」ボタン
 */
function getOpenUrlDrive(sheetName) {
  if (sheetName == '見積書') {
    return getQuotationFolderUrl();
  }
  if (sheetName == '請求書') {
    return getBillFolderUrl();
  }
  if (sheetName == '納品書') {
    return getDeliveredFolderUrl();
  }
  if (sheetName == 'ヤマト') {
    return getYamatoFolderUrl();
  }
  if (sheetName == '佐川') {
    return getSagawaFolderUrl();
  }
  if (sheetName == 'freee納品書CSV') {
    return getFreeeCSVFolderUrl();
  }
}
