/**
 * 見積書作成に必要なマスタデータを取得する
 *
 * 処理フロー:
 * 1. 各マスタシートから全レコードを取得
 *    - 商品分類
 *    - 商品
 *    - 単位
 *    - 保存温度帯
 *    - 見積納品方法
 * 2. 取得したデータを1つのオブジェクトにまとめる
 * 3. JSON文字列に変換して返却
 *
 * @returns {string} マスタデータを含むJSON文字列。以下の構造を持つ:
 *   {
 *     bunruis: Array,        // 商品分類マスタ
 *     products: Array,       // 商品マスタ
 *     units: Array,          // 単位マスタ
 *     preservations: Array,  // 保存温度帯マスタ
 *     deliveryMethods: Array // 見積納品方法マスタ
 *   }
 *
 * @see getAllRecords - 各シートからレコードを取得する関数
 * @see createQuoutation - このデータを使用して見積書を作成する関数
 *
 * @example
 * // Web側から呼び出す場合
 * const data = getQuotations();
 * const masterData = JSON.parse(data);
 * console.log(masterData.products); // 商品マスタ配列
 */
function getQuotations() {
  const bunrui = getAllRecords('商品分類');
  const product = getAllRecords('商品');
  const unit = getAllRecords('単位');
  const preservation = getAllRecords('保存温度帯');
  const deliveryMethod = getAllRecords('見積納品方法');
  const records = [];
  const rec = {};
  rec.bunruis = bunrui;
  rec.products = product;
  rec.units = unit;
  rec.preservations = preservation;
  rec.deliveryMethods = deliveryMethod;
  records.push(rec);
  Logger.log(records);
  Logger.log(JSON.stringify(records));
  return JSON.stringify(records);
}

/**
 * 見積書データを作成し、PDF出力を行う
 *
 * 処理フロー:
 * 1. 受信したJSON文字列をパースして見積書データを取得
 * 2. 一意の発注IDを生成
 * 3. 顧客情報と納品情報を抽出
 * 4. 商品リストを1つずつレコード形式に変換
 *    - 各商品に共通情報（発注ID、日付、顧客名等）を付与
 *    - 商品固有情報（分類、商品名、規格、価格等）を設定
 * 5. 変換したレコードを「見積書データ」シートに追加
 * 6. 見積書PDFを生成
 * 7. 生成したPDFファイル名を返却
 *
 * @param {string} datas - 見積書情報を含むJSON文字列。以下の構造を持つ:
 *   {
 *     customerName: string,      // 顧客名
 *     shippingToName: string,    // 納品先名
 *     deliveryMethod: string,    // 納品方法
 *     leadTime: string,          // リードタイム
 *     memo: string,              // メモ
 *     products: Array<{          // 商品リスト
 *       bunrui: string,          // 商品分類
 *       product: string,         // 商品名
 *       spec: string,            // 規格
 *       shipLot: string,         // 出荷ロット
 *       price: number,           // 価格
 *       unit: string,            // 単位
 *       place: string,           // 産地
 *       jancd: string,           // JANコード
 *       preservation: string,    // 保存温度帯
 *       bestEten: string,        // 賞味期限延長
 *       etc: string              // その他
 *     }>
 *   }
 *
 * @returns {string} 生成されたPDFファイル名（例: "顧客名_見積書_20260112_1430"）
 *
 * @see generateId - 発注IDを生成する関数
 * @see addRecords - 見積書データシートにレコードを追加する関数
 * @see makeQuoutation - 見積書PDFを作成する関数
 *
 * @example
 * // Web側から呼び出す場合
 * const quotationData = {
 *   customerName: "株式会社サンプル",
 *   shippingToName: "本社",
 *   deliveryMethod: "宅配便",
 *   leadTime: "3営業日",
 *   memo: "備考欄",
 *   products: [
 *     { bunrui: "野菜", product: "トマト", spec: "1kg", shipLot: "10", price: 500, ... }
 *   ]
 * };
 * const fileName = createQuoutation(JSON.stringify(quotationData));
 * console.log(fileName); // "株式会社サンプル_見積書_20260112_1430"
 */
function createQuoutation(datas) {
  var data = JSON.parse(datas);
  Logger.log(data);
  const adds = [];

  // 発注ID
  const deliveryId = generateId();
  const customerName = data['customerName'];
  const shippingToName = data['shippingToName'];
  const deliveryMethod = data['deliveryMethod'];
  const leadTime = data['leadTime'];
  const memo = data['memo'];
  const products = data['products'];

  products.forEach(function (wVal) {
    const record = [
      deliveryId,
      Utilities.formatDate(new Date(), 'JST', 'yyyy/MM/dd'),
      customerName,
      shippingToName,
      deliveryMethod,
      leadTime,
      wVal['bunrui'],
      wVal['product'],
      wVal['spec'],
      wVal['shipLot'],
      wVal['price'],
      wVal['unit'],
      wVal['place'],
      wVal['jancd'],
      wVal['preservation'],
      wVal['bestEten'],
      wVal['etc'],
      memo,];
    adds.push(record);
  });
  addRecords('見積書データ', adds);
  let filenm = makeQuoutation(adds);
  return filenm;
}

/**
 * 見積書シートにデータを設定し、PDF化を実行する
 *
 * 処理フロー:
 * 1. アクティブスプレッドシートと「見積書」シートを取得
 * 2. 明細行（20-27行目）を初期化してクリア
 * 3. ヘッダー情報を設定
 *    - 作成日（L3）
 *    - 顧客名（B5）※全角スペースで分割して最初の部分を使用
 *    - 納品場所（B12）
 *    - 納入方法（B13）
 *    - 出荷リードタイム（B14）
 * 4. 商品明細を設定（20行目から開始、最大8行）
 *    - B列: 商品名
 *    - D列: 規格
 *    - E列: 出荷ロット
 *    - F列: 価格
 *    - G列: 単位
 *    - H列: 産地
 *    - I列: JANコード
 *    - J列: 保存温度帯
 *    - K列: 賞味期限延長
 *    - L列: その他
 * 5. メモ情報を設定（B30）
 * 6. 商品が8行未満の場合、「以下余白」を表示
 * 7. シートの変更を確定（flush）
 * 8. PDF生成に必要な情報を取得してPDF作成を実行
 *
 * @param {Array<Array>} fileData - 見積書データの配列。各要素は以下の構造:
 *   [
 *     [0] 発注ID,
 *     [1] 作成日（yyyy/MM/dd形式）,
 *     [2] 顧客名,
 *     [3] 納品先名,
 *     [4] 納品方法,
 *     [5] リードタイム,
 *     [6] 商品分類,
 *     [7] 商品名,
 *     [8] 規格,
 *     [9] 出荷ロット,
 *     [10] 価格,
 *     [11] 単位,
 *     [12] 産地,
 *     [13] JANコード,
 *     [14] 保存温度帯,
 *     [15] 賞味期限延長,
 *     [16] その他,
 *     [17] メモ
 *   ]
 *
 * @returns {string} 生成されたPDFファイル名
 *
 * @see createQuoutation - この関数を呼び出す親関数
 * @see createQuoutationPdf - PDF生成を実行する関数
 * @see getQuotationFolderId - PDFの保存先フォルダIDを取得する関数
 *
 * 注意事項:
 * - 商品明細は最大8行まで対応
 * - 顧客名は全角スペースで分割した最初の部分のみ使用
 * - 見積書シートのレイアウトに依存（行番号固定）
 *
 * @example
 * const quotationData = [
 *   ["Q20260112001", "2026/01/12", "株式会社サンプル　本社", "本社倉庫", "宅配便", "3営業日",
 *    "野菜", "トマト", "1kg", "10", 500, "個", "熊本県", "1234567890", "冷蔵", "なし", "特記なし", "備考"]
 * ];
 * const fileName = makeQuoutation(quotationData);
 */
function makeQuoutation(fileData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const pdfSheet = ss.getSheetByName('見積書');
  // 初期化
  for (let i = 0; i < 8; i++) {
    let line = 20 + i;
    pdfSheet.getRange(line, 2).setValue("");
    pdfSheet.getRange(line, 4).setValue("");
    pdfSheet.getRange(line,5).setValue("");
    pdfSheet.getRange(line,6).setValue("");
    pdfSheet.getRange(line,7).setValue("");
    pdfSheet.getRange(line,8).setValue("");
    pdfSheet.getRange(line,9).setValue("");
    pdfSheet.getRange(line,10).setValue("");
    pdfSheet.getRange(line,11).setValue("");
    pdfSheet.getRange(line,12).setValue("");
  }

  // 見積書データ設定
  pdfSheet.getRange(3, 12).setValue(fileData[0][1]);
  pdfSheet.getRange(5, 2).setValue(fileData[0][2].split('　')[0]);
  pdfSheet.getRange(12, 2).setValue('納品場所：' + fileData[0][3]);
  pdfSheet.getRange(13, 2).setValue('納入方法：' + fileData[0][4]);
  pdfSheet.getRange(14 , 2).setValue('出荷リードタイム：' + fileData[0][5]);
  for (let i = 0; i < fileData.length; i++) {
    let line = 20 + i;
    pdfSheet.getRange(line, 2).setValue(fileData[i][7]);
    pdfSheet.getRange(line, 4).setValue(fileData[i][8]);
    pdfSheet.getRange(line,5).setValue(fileData[i][9]);
    pdfSheet.getRange(line,6).setValue(fileData[i][10]);
    pdfSheet.getRange(line,7).setValue(fileData[i][11]);
    pdfSheet.getRange(line,8).setValue(fileData[i][12]);
    pdfSheet.getRange(line,9).setValue(fileData[i][13]);
    pdfSheet.getRange(line,10).setValue(fileData[i][14]);
    pdfSheet.getRange(line,11).setValue(fileData[i][15]);
    pdfSheet.getRange(line,12).setValue(fileData[i][16]);
  }
  pdfSheet.getRange(30, 2).setValue(fileData[0][17]);
  if (fileData.length < 8) {
    let line = 20 + fileData.length;
    pdfSheet.getRange(line, 2).setValue("以下余白");
  }
  SpreadsheetApp.flush();
  const fileName = fileData[0][2].split('　')[0];
  const ssId = ss.getId();
  const shId = pdfSheet.getSheetId();
  const folderId = getQuotationFolderId();
  return createQuoutationPdf(folderId, ssId, shId, fileName);
}

/**
 * 見積書シートをPDF化してGoogle Driveに保存
 *
 * Googleスプレッドシートの「見積書」シートをPDF形式でエクスポートし、
 * 指定されたGoogle Driveフォルダに保存します。
 * PDF出力オプション（用紙サイズ、向き、余白など）を細かく指定できます。
 *
 * 処理フロー:
 * 1. Google Sheets Export API URLを構築
 *    - スプレッドシートID（ssId）とシートID（shId）を指定
 * 2. PDFエクスポートオプションを設定:
 *    - 用紙サイズ: A4
 *    - 向き: 横向き（portrait=false）
 *    - ページ幅フィット: 有効
 *    - 余白: 上下0.50インチ、左右1.00インチ
 *    - 配置: 水平中央、垂直上部
 *    - 表示設定: タイトル・シート名・グリッド線なし
 * 3. OAuthトークンを取得してヘッダーに設定
 * 4. UrlFetchApp.fetch()でPDFを取得
 * 5. ファイル名を「顧客名_見積書_yyyyMMdd_HHmm.pdf」形式で設定
 * 6. 指定されたフォルダにPDFファイルを保存
 * 7. ファイル名（拡張子なし）を返却
 *
 * @param {string} folderId - PDF保存先のGoogle DriveフォルダID
 * @param {string} ssId - 見積書シートを含むスプレッドシートのID
 * @param {string} shId - 見積書シートのシートID
 * @param {string} billingNumber - 顧客名（ファイル名のプレフィックスとして使用）
 * @returns {string} 生成されたPDFファイル名（拡張子なし、例: "株式会社サンプル_見積書_20260112_1430"）
 *
 * @see makeQuoutation() - この関数を呼び出す親関数
 * @see getQuotationFolderId() - フォルダIDを取得する関数（config.js）
 *
 * PDFエクスポートオプション詳細:
 * - size: A4（用紙サイズ）
 * - portrait: false（横向き）
 * - fitw: true（ページ幅を用紙にフィット）
 * - top_margin: 0.50（上余白 インチ）
 * - right_margin: 1.00（右余白 インチ）
 * - bottom_margin: 0.50（下余白 インチ）
 * - left_margin: 1.00（左余白 インチ）
 * - horizontal_alignment: CENTER（水平中央揃え）
 * - vertical_alignment: TOP（垂直上部揃え）
 * - printtitle: false（スプレッドシート名を非表示）
 * - sheetnames: false（シート名を非表示）
 * - gridlines: false（グリッド線を非表示）
 * - fzr: false（固定行を非表示）
 * - fzc: false（固定列を非表示）
 *
 * @example
 * const fileName = createQuoutationPdf(
 *   "1abc2def3ghi...",  // フォルダID
 *   "4jkl5mno6pqr...",  // スプレッドシートID
 *   "123456789",        // シートID
 *   "株式会社サンプル"  // 顧客名
 * );
 * // returns: "株式会社サンプル_見積書_20260112_1430"
 *
 * Google Sheets Export API:
 * https://developers.google.com/drive/api/guides/ref-export-formats
 */
function createQuoutationPdf(folderId, ssId, shId, billingNumber) {

  //PDFを作成するためのベースとなるURL
  let baseUrl = "https://docs.google.com/spreadsheets/d/"
    + ssId
    + "/export?gid="
    + shId;

  //PDFのオプションを指定
  let pdfOptions = "&exportFormat=pdf&format=pdf"
    + "&size=A4" //用紙サイズ (A4)
    + "&portrait=false"  //用紙の向き true: 縦向き / false: 横向き
    + "&fitw=true"  //ページ幅を用紙にフィットさせるか true: フィットさせる / false: 原寸大
    + "&top_margin=0.50" //上の余白
    + "&right_margin=1.00" //右の余白
    + "&bottom_margin=0.50" //下の余白
    + "&left_margin=1.00" //左の余白
    + "&horizontal_alignment=CENTER" //水平方向の位置
    + "&vertical_alignment=TOP" //垂直方向の位置
    + "&printtitle=false" //スプレッドシート名の表示有無
    + "&sheetnames=false" //シート名の表示有無
    + "&gridlines=false" //グリッドラインの表示有無
    + "&fzr=false" //固定行の表示有無
    + "&fzc=false" //固定列の表示有無;

  //PDFを作成するためのURL
  let url = baseUrl + pdfOptions;

  //アクセストークンを取得する
  let token = ScriptApp.getOAuthToken();

  //headersにアクセストークンを格納する
  let options = {
    headers: {
      'Authorization': 'Bearer ' + token
    }
  };
  const filenm = billingNumber + '_見積書_' + Utilities.formatDate(new Date(), 'JST', 'yyyyMMdd_HHmm');
  //PDFを作成する
  let blob = UrlFetchApp.fetch(url, options).getBlob().setName(filenm + '.pdf');

  //PDFの保存先フォルダー
  //フォルダーIDは引数のfolderIdを使用します
  let folder = DriveApp.getFolderById(folderId);

  //PDFを指定したフォルダに保存する
  folder.createFile(blob);
  return filenm;
}