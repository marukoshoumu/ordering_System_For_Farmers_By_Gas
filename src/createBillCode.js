// PDF出力先
const BILL_PDF_OUTDIR = DriveApp.getFolderById(getBillPdfFolderId());
// 請求書のテンプレートファイル　
const BILL_TEMPLATE = DriveApp.getFileById(getBillTemplateId());
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

// 請求書のドキュメントを置換
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
  const wCopyFile = BILL_TEMPLATE.makeCopy()
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
// 請求書のファイル生成
function createBillFile(records, billDate) {
  // PDF変換する元ファイルを作成する
  let wFileRtn = createBillGDoc(records, billDate);
  // PDF変換
  createBillPdf(wFileRtn[0], wFileRtn[1]);
  // PDF変換したあとは元ファイルを削除する
  DriveApp.getFileById(wFileRtn[0]).setTrashed(true);
  return wFileRtn[1];
}
// PDF生成
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
  return BILL_PDF_OUTDIR.createFile(wBlob).getId();
}

