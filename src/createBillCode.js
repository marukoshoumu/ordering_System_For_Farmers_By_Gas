// PDF出力先フォルダとテンプレートファイルを取得する関数
// グローバルスコープでの初期化を避け、必要時に取得する
function getBillPdfOutdir() {
  const folderId = getBillPdfFolderId();
  if (!folderId) {
    throw new Error('BILL_PDF_FOLDER_ID がスクリプトプロパティに設定されていません');
  }
  try {
    return DriveApp.getFolderById(folderId);
  } catch (e) {
    throw new Error(`請求書PDF出力フォルダ(ID: ${folderId})へのアクセスに失敗しました: ${e.message}`);
  }
}

function getBillTemplate() {
  const templateId = getBillTemplateId();
  if (!templateId) {
    throw new Error('BILL_TEMPLATE_ID がスクリプトプロパティに設定されていません');
  }
  try {
    return DriveApp.getFileById(templateId);
  } catch (e) {
    throw new Error(`請求書テンプレート(ID: ${templateId})へのアクセスに失敗しました: ${e.message}`);
  }
}

/**
 * 請求書を作成するメイン関数
 *
 * 指定された期間と顧客条件に基づいて受注データを集計し、請求書PDFを生成します。
 *
 * 処理フロー:
 * 1. JSON形式の入力データをパースして、顧客名・期間・請求日を取得
 * 2. 「受注」シートから全レコードを取得
 * 3. 期間条件でフィルタリング（納品日が指定期間内）
 * 4. 顧客名が指定されている場合は該当顧客のみ、未指定の場合は全顧客を対象
 * 5. 同一顧客の同一商品を集約（受注数と販売価格を合算）
 * 6. 各顧客ごとにcreateBillFile()を呼び出して請求書PDFを生成
 * 7. 生成されたファイル名リストをJSON形式で返却
 *
 * @param {string} datas - JSON文字列形式の入力データ
 * @param {string} datas.customerName - 対象顧客名（空文字列の場合は全顧客）
 * @param {string} datas.targetDateFrom - 集計期間の開始日（ISO形式）
 * @param {string} datas.targetDateTo - 集計期間の終了日（ISO形式）
 * @param {string} datas.billDate - 請求日（ISO形式）
 *
 * @returns {string} 生成された請求書ファイル名の配列をJSON形式で返却
 *
 * @see createBillFile - 請求書ファイル生成関数
 * @see getAllRecords - 受注シートデータ取得関数（「受注」シート）
 *
 * @example
 * // 特定顧客の2024年1月分の請求書を作成
 * const input = JSON.stringify({
 *   customerName: "株式会社サンプル　田中太郎",
 *   targetDateFrom: "2024-01-01",
 *   targetDateTo: "2024-01-31",
 *   billDate: "2024-02-01"
 * });
 * const result = createBill(input);
 * // returns: '["株式会社サンプル_請求文書_20240201_1030"]'
 *
 * @example
 * // 全顧客の2024年1月分の請求書を一括作成
 * const input = JSON.stringify({
 *   customerName: "",
 *   targetDateFrom: "2024-01-01",
 *   targetDateTo: "2024-01-31",
 *   billDate: "2024-02-01"
 * });
 * const result = createBill(input);
 */
function createBill(datas) {
  var data = JSON.parse(datas);
  var customerName = data['customerName'];
  var targetFrom = Utilities.formatDate(new Date(data['targetDateFrom']), 'JST', 'yyyy/MM/dd');
  var targetTo = Utilities.formatDate(new Date(data['targetDateTo']), 'JST', 'yyyy/MM/dd');
  var dateNow = Utilities.formatDate(new Date(), 'JST', 'yyyyMMdd_HHmm');
  var billDate = Utilities.formatDate(new Date(data['billDate']), 'JST', 'yyyy/MM/dd');
  const items = getAllRecords('受注');
  var targetLists = [];
  if (customerName != '') {
    targetLists = items.reduce(function (result, target) {
      if (Utilities.formatDate(new Date(target['納品日']), 'JST', 'yyyy/MM/dd') >= targetFrom
        && Utilities.formatDate(new Date(target['納品日']), 'JST', 'yyyy/MM/dd') <= targetTo
        && target['顧客名'] === customerName) {
        if (result.length != 0) {
          const ele = result[0]['商品'].find(value => value['商品名'] === target['商品名']);
          if (ele) {
            ele['受注数'] += Number(target['受注数']);
            ele['販売価格'] += Number(target['販売価格']);
          }
          else {
            var product = {};
            var products = result[0]['商品'].concat();
            product['商品分類'] = target['商品分類'];
            product['商品名'] = target['商品名'];
            product['受注数'] = Number(target['受注数']);
            product['販売価格'] = Number(target['販売価格']);
            products.push(product);
            result[0]['商品'] = products;
          }
        }
        else {
          var record = {};
          var product = {};
          var products = [];
          product['商品分類'] = target['商品分類'];
          product['商品名'] = target['商品名'];
          product['受注数'] = Number(target['受注数']);
          product['販売価格'] = Number(target['販売価格']);
          products.push(product);
          record['顧客名'] = target['顧客名'];
          record['顧客住所'] = target['顧客住所'];
          record['顧客郵便番号'] = target['顧客郵便番号'];
          record['受注ID'] = target['受注ID'];
          record['納品日'] = target['納品日'];
          record['商品'] = products;
          result.push(record);
        }
      }
      return result;
    }, []);
  }
  else {
    targetLists = items.reduce(function (result, target) {
      if (Utilities.formatDate(new Date(target['納品日']), 'JST', 'yyyy/MM/dd') >= targetFrom
        && Utilities.formatDate(new Date(target['納品日']), 'JST', 'yyyy/MM/dd') <= targetTo) {
        const element = result.find(value => value['顧客名'] === target['顧客名']);
        if (element) {
          const ele = element['商品'].find(value => value['商品名'] === target['商品名']);
          if (ele) {
            ele['受注数'] += Number(target['受注数']);
            ele['販売価格'] += Number(target['販売価格']);
          }
          else {
            var product = {};
            var products = element['商品'].concat();
            product['商品分類'] = target['商品分類'];
            product['商品名'] = target['商品名'];
            product['受注数'] = Number(target['受注数']);
            product['販売価格'] = Number(target['販売価格']);
            products.push(product);
            element['商品'] = products;
          }
        } else {
          var record = {};
          var product = {};
          var products = [];
          product['商品分類'] = target['商品分類'];
          product['商品名'] = target['商品名'];
          product['受注数'] = Number(target['受注数']);
          product['販売価格'] = Number(target['販売価格']);
          products.push(product);
          record['顧客名'] = target['顧客名'];
          record['顧客住所'] = target['顧客住所'];
          record['顧客郵便番号'] = target['顧客郵便番号'];
          record['受注ID'] = target['受注ID'];
          record['納品日'] = target['納品日'];
          record['商品'] = products;
          result.push(record);
        }
      }
      return result;
    }, []);
  }
  Logger.log(targetLists);
  var datas = [];
  for (var target of targetLists) {
    var data = createBillFile(target, billDate);
    if (data) {
      datas.push(data);
    }
  }
  return JSON.stringify(datas);
}

/**
 * 請求書用Googleドキュメントを生成する関数
 *
 * テンプレートファイルを複製し、受注データと顧客情報を使って請求書を作成します。
 * 商品情報、税率計算、入金期日などを反映した請求書ドキュメントを生成します。
 *
 * 処理フロー:
 * 1. 「顧客情報」「発送先情報」「商品」シートからマスタデータを取得
 * 2. 顧客名から該当する顧客情報を検索（会社名・氏名でマッチング）
 * 3. 顧客情報が見つからない場合は受注データから最低限の情報を取得
 * 4. テンプレートファイルをコピーして新しいドキュメントを作成
 * 5. 顧客情報（会社名、郵便番号、住所など）をドキュメントに反映
 * 6. 入金期日を計算（翌末日、翌々中日、翌々末日、2週間後の4パターン）
 * 7. 商品情報（最大10品目）をループで処理
 *    - 商品分類、商品名、価格、数量、金額を設定
 *    - 税率（8%または10%）に基づいて税額を計算
 * 8. 合計金額、税額（8%・10%別）、総額を計算してドキュメントに反映
 * 9. ファイル名を「会社名_請求文書_yyyyMMdd_HHmm」形式で設定
 * 10. ドキュメントIDとファイル名を返却
 *
 * @param {Object} rowVal - 受注レコードデータ
 * @param {string} rowVal.顧客名 - 顧客名（「会社名　氏名」形式または「氏名」のみ）
 * @param {string} rowVal.顧客住所 - 顧客住所（顧客情報が見つからない場合に使用）
 * @param {string} rowVal.顧客郵便番号 - 顧客郵便番号（7桁数字）
 * @param {string} rowVal.受注ID - 受注ID（請求書に表示）
 * @param {Date} rowVal.納品日 - 納品日（入金期日計算で使用）
 * @param {Array<Object>} rowVal.商品 - 商品情報の配列（最大10品目）
 * @param {string} rowVal.商品[].商品分類 - 商品の分類
 * @param {string} rowVal.商品[].商品名 - 商品名
 * @param {number} rowVal.商品[].受注数 - 受注数量
 * @param {number} rowVal.商品[].販売価格 - 単価
 * @param {string} billDate - 請求日（ISO形式）
 *
 * @returns {Array<string>} [ドキュメントID, ファイル名] の配列
 * @returns {string} returns[0] - 生成されたGoogleドキュメントのID
 * @returns {string} returns[1] - 設定されたファイル名（会社名_請求文書_yyyyMMdd_HHmm）
 *
 * @see getAllRecords - マスタデータ取得関数（「顧客情報」「発送先情報」「商品」シート）
 * @see BILL_TEMPLATE - 請求書テンプレートファイル（グローバル定数）
 *
 * @example
 * const record = {
 *   顧客名: "株式会社サンプル　田中太郎",
 *   顧客住所: "東京都渋谷区1-2-3",
 *   顧客郵便番号: "1500001",
 *   受注ID: "ORD-2024-001",
 *   納品日: new Date("2024-01-15"),
 *   商品: [
 *     { 商品分類: "野菜", 商品名: "トマト", 受注数: 10, 販売価格: 300 },
 *     { 商品分類: "野菜", 商品名: "きゅうり", 受注数: 5, 販売価格: 200 }
 *   ]
 * };
 * const [docId, fileName] = createBillGDoc(record, "2024-02-01");
 * // returns: ["1a2b3c4d5e6f...", "株式会社サンプル_請求文書_20240201_1030"]
 *
 * @note 入金期日の計算パターン:
 *   - 翌末日: 翌月末日
 *   - 翌々中日: 翌々月15日
 *   - 2週間後: 納品日から14日後
 *   - 翌々末日: 翌々月末日
 *   - 未設定の場合: 翌末日として計算
 *
 * @note 税率計算:
 *   - 商品マスタの税率が8を超える場合は10%として計算
 *   - それ以外は8%として計算
 *   - 各税率ごとに対象金額と税額を集計
 */
function createBillGDoc(rowVal, billDate) {
  // 顧客情報シートを取得する
  const customerItems = getAllRecords('顧客情報');
  const shippingToItems = getAllRecords('発送先情報');
  const productItems = getAllRecords('商品');
  var customerItem = [];
  const customerName = rowVal['顧客名'].split('　');
  // データ走査
  customerItems.forEach(function (wVal) {
    if (customerName.length > 1) {
      // 会社名と同じ
      if (customerName[1] == wVal['氏名'] && customerName[0] == wVal['会社名']) {
        customerItem = wVal;
      }
    }
    else {
      if (customerName[0] == wVal['氏名']) {
        customerItem = wVal;
      }
      if (customerName[0] == wVal['会社名']) {
        customerItem = wVal;
      }
    }
  });
  if (customerItem.length == 0) {
    customerItem['会社名'] = rowVal['顧客名'].split('　')[0];
    customerItem['住所１'] = rowVal['顧客住所'];
    customerItem['郵便番号'] = rowVal['顧客郵便番号'];
    customerItem['住所２'] = "";
    customerItem['入金期日'] = "";
  }
  // テンプレートファイルをコピーする
  const wCopyFile = getBillTemplate().makeCopy()
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
  wCopyDocBody = wCopyDocBody.replaceText('{{delivery_num}}', rowVal['受注ID']);
  wCopyDocBody = wCopyDocBody.replaceText('{{delivery_date}}', Utilities.formatDate(new Date(billDate), 'JST', 'yyyy年MM月dd日'));
  const nowDate = new Date();
  var depositDate;
  if (!customerItem['入金期日'] || customerItem['入金期日'] === '翌末日') {
    var lastDate = new Date(nowDate.getFullYear(), nowDate.getMonth() + 2, 0);
    var yyyy = lastDate.getFullYear();
    var mm = ("0" + (lastDate.getMonth() + 1)).slice(-2);
    var dd = ("0" + lastDate.getDate()).slice(-2);
    depositDate = yyyy + '-' + mm + '-' + dd;
  }
  if (customerItem['入金期日'] === '翌々中日') {
    var lastDate = new Date(nowDate.getFullYear(), nowDate.getMonth() + 2, 15);
    var yyyy = lastDate.getFullYear();
    var mm = ("0" + (lastDate.getMonth() + 1)).slice(-2);
    var dd = ("0" + lastDate.getDate()).slice(-2);
    depositDate = yyyy + '-' + mm + '-' + dd;
  }
    if (customerItem['入金期日'] === '２週間後') {
    var nDate = new Date(rowVal['納品日']);
    nDate.setDate(nDate.getDate() + 14)
    var yyyy = nDate.getFullYear();
    var mm = ("0" + (nDate.getMonth() + 1)).slice(-2);
    var dd = ("0" + nDate.getDate()).slice(-2);
    depositDate = yyyy + '-' + mm + '-' + dd;
  }
  if (customerItem['入金期日'] === '翌々末日') {
    var lastDate = new Date(nowDate.getFullYear(), nowDate.getMonth() + 3, 0);
    var yyyy = lastDate.getFullYear();
    var mm = ("0" + (lastDate.getMonth() + 1)).slice(-2);
    var dd = ("0" + lastDate.getDate()).slice(-2);
    depositDate = yyyy + '-' + mm + '-' + dd;
  }

  wCopyDocBody = wCopyDocBody.replaceText('{{depositDate}}', Utilities.formatDate(new Date(depositDate), 'JST', 'yyyy年MM月dd日'));
  let totals = 0;
  let tentax = 0;
  let eigtax = 0;
  let tentax_t = 0;
  let eigtax_t = 0;
  let tax = 0;
  let amount = 0;

  for (let i = 0; i < 10; i++) {
    var total = 0;
    if (i < rowVal['商品'].length) {
      var productData;
      // 商品分類
      var changeText = '{{bunrui' + (i + 1) + '}}';
      wCopyDocBody = wCopyDocBody.replaceText(changeText, rowVal['商品'][i]['商品分類']);
      // 商品名
      changeText = '{{product' + (i + 1) + '}}';
      wCopyDocBody = wCopyDocBody.replaceText(changeText, rowVal['商品'][i]['商品名']);
      Logger.log(rowVal);
      Logger.log(rowVal['商品'][i]['商品名']);
      // データ走査
      productItems.forEach(function (wVal) {
        // 商品名と同じ
        if (wVal['商品名'] == rowVal['商品'][i]['商品名']) {
          productData = wVal;
        }
      });
      if (!productData) {
        productData['税率'] = 8;
      }
      // 価格（P)
      changeText = '{{price' + (i + 1) + '}}';
      wCopyDocBody = wCopyDocBody.replaceText(changeText, '￥ ' + rowVal['商品'][i]['販売価格']);
      // 数量
      changeText = '{{c' + (i + 1) + '}}';
      wCopyDocBody = wCopyDocBody.replaceText(changeText, rowVal['商品'][i]['受注数']);
      // 金額
      changeText = '{{amount' + (i + 1) + '}}';
      total = rowVal['商品'][i]['受注数'] * rowVal['商品'][i]['販売価格'];
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
  let fileName = customerItem['会社名'] + '_請求文書_' + Utilities.formatDate(new Date(), 'JST', 'yyyyMMdd_HHmm');
  wCopyFile.setName(fileName);
  // コピーしたファイルIDとファイル名を返却する（あとでこのIDをもとにPDFに変換するため）
  return [wCopyFileId, fileName];
}
/**
 * 請求書ファイル生成（Googleドキュメント作成 → PDF変換 → 元ファイル削除）
 *
 * 受注レコードと請求日から請求書Googleドキュメントを作成し、
 * PDFに変換してGoogle Driveに保存した後、元のドキュメントを削除します。
 *
 * 処理フロー:
 * 1. createBillGDoc()でGoogleドキュメント作成 → [ドキュメントID, ファイル名]を取得
 * 2. createBillPdf()でPDF変換 → PDF ID返却
 * 3. DriveApp.getFileById().setTrashed()で元ドキュメントを削除
 * 4. ファイル名を返却
 *
 * @param {Object} records - 受注レコードデータ（createBillGDocと同じ形式）
 * @param {string} billDate - 請求日（ISO形式）
 * @returns {string} 生成されたPDFファイル名（例: "株式会社サンプル_請求文書_20240201_1030"）
 *
 * @see createBillGDoc() - Googleドキュメント作成
 * @see createBillPdf() - PDF変換
 * @see createBill() - この関数を呼び出すメイン関数
 */
function createBillFile(records, billDate) {
  // PDF変換する元ファイルを作成する
  let wFileRtn = createBillGDoc(records, billDate);
  // PDF変換
  createBillPdf(wFileRtn[0], wFileRtn[1]);
  // PDF変換したあとは元ファイルを削除する
  DriveApp.getFileById(wFileRtn[0]).setTrashed(true);
  return wFileRtn[1];
}

/**
 * Google ドキュメントをPDFに変換してGoogle Driveに保存
 *
 * 指定されたGoogleドキュメントIDをPDF形式にエクスポートし、
 * 請求書PDF出力フォルダ（BILL_PDF_OUTDIR）に保存します。
 *
 * 処理フロー:
 * 1. Google Docs Export API URL構築（/export?exportFormat=pdf）
 * 2. OAuthトークンをヘッダーに設定
 * 3. UrlFetchApp.fetch()でPDF取得
 * 4. Blobオブジェクトに変換してファイル名設定（.pdf拡張子追加）
 * 5. BILL_PDF_OUTDIRフォルダにPDFファイル作成
 * 6. 作成したPDFのファイルIDを返却
 *
 * @param {string} docId - GoogleドキュメントのID
 * @param {string} fileName - 保存するファイル名（拡張子なし）
 * @returns {string} 作成されたPDFファイルのGoogle Drive ID
 *
 * @see BILL_PDF_OUTDIR - PDF出力先フォルダ（グローバル定数）
 * @see createBillFile() - この関数を呼び出す関数
 * @see getBillPdfFolderId() - PDF出力フォルダID取得（config.js）
 *
 * Google Docs Export API:
 * https://developers.google.com/drive/api/guides/ref-export-formats
 */
function createBillPdf(docId, fileName) {
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
  return getBillPdfOutdir().createFile(wBlob).getId();
}

