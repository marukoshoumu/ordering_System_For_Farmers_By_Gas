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