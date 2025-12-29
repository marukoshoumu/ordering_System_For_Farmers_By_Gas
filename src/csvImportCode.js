function csvFileImport(fileClass, datas) {
  var data = JSON.parse(datas);
  Logger.log(data);
  // 食べチョク
  if (fileClass == '1') {
    tabeAddRecord(data);
  }
  // ふるさと納税
  if (fileClass == '2') {
    furuAddRecord(data);
  }
  // ポケマル
  if (fileClass == '3') {
    maruSheAddRecord(data);
  }
}
function tabeDateFormat(str) {
  if (str == null || str == '') {
    return '';
  }
  if (str.length == 8) {
    return str.substring(0, 4) + '/' + str.substring(4, 6) + '/' + str.substring(6, 8);
  }
  else {
    return '';
  }
}
function tabeAddRecord(data) {
  const adds = [];
  const productItems = getAllRecords('食べチョク商品');
  data.forEach(function (wVal) {
    if (wVal['注文番号'] == '' || wVal['注文番号'] == null) {
      return;
    }
    const productData = productItems.find((ele) => ele['商品名'] == wVal['商品名']);

    const record = [
      wVal['注文番号'],
      Utilities.formatDate(new Date(tabeDateFormat(wVal['注文日'])), 'JST', 'yyyy/MM/dd'),
      '',
      productData['変換値'],
      productData['変換数量'],
      Number(wVal['商品代金(お客様に提示した額)']) + Number(wVal['生産者に支払われる送料']) + Number(wVal['生産者に支払われるクール代']),
      CONFIG.TABECHOKU.COMPANY,
      CONFIG.TABECHOKU.ZIPCODE,
      CONFIG.TABECHOKU.ADDRESS,
      '',
      wVal['お届け先名'],
      wVal['お届け先郵便番号'],
      wVal['お届け先住所'],
      wVal['お届け先電話番号'],
      CONFIG.COMPANY.NAME,
      CONFIG.COMPANY.ZIPCODE,
      CONFIG.COMPANY.ADDRESS,
      CONFIG.COMPANY.TEL,
      Utilities.formatDate(new Date(tabeDateFormat(wVal['発送日'])), 'JST', 'yyyy/MM/dd'),
      Utilities.formatDate(new Date(tabeDateFormat(wVal['お届け日'])), 'JST', 'yyyy/MM/dd'),
      '',
      '',
      wVal['配送業者'],
      wVal['お届け時間帯'],
      '',
      '',
      '',
      '',
      '',
      '',
      '',
      wVal['発払い'],
      wVal['クール種別'],
      '',
      '',
      '',
      '',
      '',
      '',
      '',
      wVal['特記事項'],
      Number(wVal['商品代金(お客様に提示した額)']) + Number(wVal['生産者に支払われる送料']) + Number(wVal['生産者に支払われるクール代'])
    ];
    adds.push(record);
  });
  Logger.log(adds);
  addRecords('受注', adds);
}
function furuAddRecord(data) {
  const adds = [];
  const productItems = getAllRecords('ふるさと納税商品');
  data.forEach(function (wVal) {
    if (wVal['配送管理ID'] == '' || wVal['配送管理ID'] == null) {
      return;
    }
    const productData = productItems.find((ele) => ele['商品名'] == wVal['返礼品']);
    const record = [
      wVal['配送管理ID'],
      Utilities.formatDate(new Date(), 'JST', 'yyyy/MM/dd'),
      '',
      productData['変換値'],
      Number(productData['変換数量'])*Number(wVal['個数']),
      wVal['提供価格'],
      CONFIG.FURUSATO.COMPANY,
      CONFIG.FURUSATO.ZIPCODE,
      CONFIG.FURUSATO.ADDRESS,
      CONFIG.FURUSATO.TEL,
      wVal['お届け先名'],
      wVal['届け先郵便番号'],
      wVal['届け先住所１'] + wVal['届け先住所２'],
      wVal['届け先電話番号'],
      CONFIG.FURUSATO.COMPANY,
      CONFIG.FURUSATO.ZIPCODE,
      CONFIG.FURUSATO.ADDRESS,
      CONFIG.FURUSATO.TEL,
      Utilities.formatDate(new Date(wVal['発送日']), 'JST', 'yyyy/MM/dd'),
      '',
      '',
      '',
      wVal['ヤマト産直伝票'],
      '',
      '',
      '',
      '',
      '',
      '',
      '',
      '',
      '',
      '',
      '',
      '',
      '',
      '',
      '',
      '',
      '',
      '',
      Number(productData['変換数量'])*Number(wVal['個数']) * Number(wVal['提供価格'])
    ];
    adds.push(record);
  });
  Logger.log(adds);
  addRecords('受注', adds);
}
function maruDateFormat(str) {
  if (str == null || str == '') {
    return new Date();
  }
  var maruDate = new Date();
  if (str.length > 5) {
    const dateStr = '2020/'+ str.substring(0, 5);
    let dateObject = new Date(dateStr);
    dateObject.setFullYear(maruDate.getFullYear());
    return dateObject;
  }
  else {
    return maruDate;
  }
}
function maruSheAddRecord(data) {
  const adds = [];
  const productItems = getAllRecords('ポケマル商品');
  data.forEach(function (wVal) {
    if (wVal['注文ID'] == '' || wVal['注文ID'] == null) {
      return;
    }
    const productData = productItems.find((ele) => ele['商品名'] == wVal['商品名'] && ele['分量'] == wVal['分量']);
    const record = [
      wVal['注文ID'],
      Utilities.formatDate(maruDateFormat(wVal['注文確定日']), 'JST', 'yyyy/MM/dd'),
      '',
      productData['変換値'],
      productData['変換数量'],
      '',
      CONFIG.POCKEMARU.COMPANY,
      CONFIG.POCKEMARU.ZIPCODE,
      CONFIG.POCKEMARU.ADDRESS,
      '',
      wVal['配送先宛名'],
      wVal['配送先〒'],
      wVal['配送先都道府県'] + wVal['配送先市区町村'] + wVal['配送先番地'] + wVal['配送先建物名・部屋番号'],
      wVal['配送先電話番号'],
      CONFIG.COMPANY.NAME,
      CONFIG.COMPANY.ZIPCODE,
      CONFIG.COMPANY.ADDRESS,
      CONFIG.COMPANY.TEL,
      Utilities.formatDate(maruDateFormat(wVal['発送日']), 'JST', 'yyyy/MM/dd'),
      wVal['着日指定'],
      '',
      '',
      wVal['ヤマト産直伝票'],
      '',
      '',
      '',
      '',
      '',
      '',
      '',
      '',
      '',
      '',
      '',
      '',
      '',
      '',
      '',
      '',
      '',
      wVal['申し送り事項'],
      ''
    ];
    adds.push(record);
  });
  Logger.log(adds);
  addRecords('受注', adds);
}