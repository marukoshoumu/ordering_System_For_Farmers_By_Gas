// 郵便番号から住所を取得する関数
function getAddressFromPostalCode(postalCodes) {
  var postalCode = JSON.parse(postalCodes);
  var url = "https://zipcloud.ibsnet.co.jp/api/search?zipcode=" + postalCode['postalCode'];
  var response = UrlFetchApp.fetch(url);
  var data = JSON.parse(response.getContentText());
  console.log(data);
  if (data.results && data.results.length > 0) {
    var address1 = data.results[0].address1;
    var address2 = data.results[0].address2;
    var address3 = data.results[0].address3;
    var res = {};
    res['address1'] = address1;
    res['address2'] = address2;
    res['address3'] = address3;
    return JSON.stringify(res);
  } else {
    return null;
  }
}
// 顧客情報登録
function addCustomer(record) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('顧客情報');
  const lastRow = sheet.getLastRow();
  const sheetBk = ss.getSheetByName('顧客情報BK');
  const lastBkRow = sheet.getLastRow();
  var data = JSON.parse(record);
  var rows = [];
  var row = [
    data['customerClassDg'],
    data['customerShowNameDg'],
    data['customerFuriganaDg'],
    data['customerNameDg'],
    data['customerDepDg'],
    data['customerPostDg'],
    data['customerIdentDg'],
    data['customerZipcodeDg'],
    data['customerAddress1Dg'],
    data['customerAddress2Dg'],
    data['customerTelDg'],
    data['customerPhoneDg'],
    data['customerFaxDg'],
    data['customerMailDg'],
    data['customerBillDg'],
    data['customerDepositDateDg'],
    data['customerMemoDg']
  ];
  rows.push(row);
  sheet.getRange(lastRow + 1, 1, 1, row.length).setNumberFormat('@').setValues(rows).setBorder(true, true, true, true, true, true);
  sheetBk.getRange(lastBkRow + 1, 1, 1, row.length).setNumberFormat('@').setValues(rows).setBorder(true, true, true, true, true, true);
}
// 顧客情報・発送先情報登録
function addCustomerShippingTo(record) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const customerSheet = ss.getSheetByName('顧客情報');
  const customerLastRow = customerSheet.getLastRow();
  const customerBkSheet = ss.getSheetByName('顧客情報BK');
  const customerBkLastRow = customerSheet.getLastRow();
  const shippingToSheet = ss.getSheetByName('発送先情報');
  const shippingToLastRow = shippingToSheet.getLastRow();
  const shippingToBkSheet = ss.getSheetByName('発送先情報BK');
  const shippingToBkLastRow = shippingToSheet.getLastRow();
  var data = JSON.parse(record);
  var customerRows = [];
  var customerRow = [
    data['customerClassDg'],
    data['customerShowNameDg'],
    data['customerFuriganaDg'],
    data['customerNameDg'],
    data['customerDepDg'],
    data['customerPostDg'],
    data['customerIdentDg'],
    data['customerZipcodeDg'],
    data['customerAddress1Dg'],
    data['customerAddress2Dg'],
    data['customerTelDg'],
    data['customerPhoneDg'],
    data['customerFaxDg'],
    data['customerMailDg'],
    data['customerBillDg'],
    data['customerDepositDateDg'],
    data['customerMemoDg']
  ];
  customerRows.push(customerRow);
  customerSheet.getRange(customerLastRow + 1, 1, 1, customerRow.length).setNumberFormat('@').setValues(customerRows).setBorder(true, true, true, true, true, true);
  customerBkSheet.getRange(customerBkLastRow + 1, 1, 1, customerRow.length).setNumberFormat('@').setValues(customerRows).setBorder(true, true, true, true, true, true);
  var shippingToRows = [];
  var shippingToRow = [
    data['customerNameDg'],
    data['customerDepDg'],
    data['customerIdentDg'],
    data['customerZipcodeDg'],
    data['customerAddress1Dg'],
    data['customerAddress2Dg'],
    data['customerTelDg'],
    data['customerMailDg'],
    data['customerMemoDg']
  ];
  shippingToRows.push(shippingToRow);
  shippingToSheet.getRange(shippingToLastRow + 1, 1, 1, shippingToRow.length).setNumberFormat('@').setValues(shippingToRows).setBorder(true, true, true, true, true, true);
  shippingToBkSheet.getRange(shippingToBkLastRow + 1, 1, 1, shippingToRow.length).setNumberFormat('@').setValues(shippingToRows).setBorder(true, true, true, true, true, true);
}
// 発送先情報登録
function addShippingTo(record) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shippingToSheet = ss.getSheetByName('発送先情報');
  const shippingToLastRow = shippingToSheet.getLastRow();
  const shippingToBkSheet = ss.getSheetByName('発送先情報BK');
  const shippingToBkLastRow = shippingToSheet.getLastRow();
  var data = JSON.parse(record);
  var shippingToRows = [];
  var shippingToRow = [
    data['shippingToNameDg'],
    data['shippingToDepDg'],
    data['shippingToIdentDg'],
    data['shippingToZipcodeDg'],
    data['shippingToAddress1Dg'],
    data['shippingToAddress2Dg'],
    data['shippingToTelDg'],
    data['shippingToMailDg'],
    data['shippingToMemoDg']
  ];
  shippingToRows.push(shippingToRow);
  shippingToSheet.getRange(shippingToLastRow + 1, 1, 1, shippingToRow.length).setNumberFormat('@').setValues(shippingToRows).setBorder(true, true, true, true, true, true);
  shippingToBkSheet.getRange(shippingToBkLastRow + 1, 1, 1, shippingToRow.length).setNumberFormat('@').setValues(shippingToRows).setBorder(true, true, true, true, true, true);
}
// 顧客情報検索
function customerSearch(record) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('顧客情報');
  var data = JSON.parse(record);
  const values = sheet.getDataRange().getValues();
  var bunrui = data['customerSearchBunruiDg'];
  var searchClass = data['customerSearchClassDg'];
  var key = data['customerSearchKeyDg'];
  const records = [];
  var targetCol = 3;
  if (bunrui == "2") {
    targetCol = 6;
  }
  if (bunrui == "3") {
    targetCol = 1;
  }
  if (bunrui == "4") {
    targetCol = 2;
  }
  for (const value of values) {
    if (searchClass != "") {
      if (searchClass == value[0] && String(value[targetCol].includes(key))) {
        const rec = {};
        rec['会社名'] = value[3];
        rec['氏名'] = value[6];
        rec['郵便番号'] = value[7];
        rec['住所１'] = value[8];
        rec['住所２'] = value[9];
        rec['TEL'] = value[10];
        records.push(rec);
      }
    }
    else {
      if (String(value[targetCol]).includes(key)) {
        const rec = {};
        rec['会社名'] = value[3];
        rec['氏名'] = value[6];
        rec['郵便番号'] = value[7];
        rec['住所１'] = value[8];
        rec['住所２'] = value[9];
        rec['TEL'] = value[10];
        records.push(rec);
      }
    }
  }
  return JSON.stringify(records);
}
// 発送先情報検索
function shippingToSearch(record) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('発送先情報');
  var data = JSON.parse(record);
  const values = sheet.getDataRange().getValues();
  Logger.log(data);
  Logger.log(values);
  var bunrui = data['shippingToSearchBunruiDg'];
  var key = data['shippingToSearchKeyDg'];
  const records = [];
  var targetCol = 0;
  if (bunrui == "2") {
    targetCol = 2;
  }
  for (const value of values) {
    if (String(value[targetCol]).includes(key)) {
      const rec = {};
      rec['会社名'] = value[0];
      rec['氏名'] = value[2];
      rec['郵便番号'] = value[3];
      rec['住所１'] = value[4];
      rec['住所２'] = value[5];
      rec['TEL'] = value[6];
      records.push(rec);
    }
  }
  return JSON.stringify(records);
}
// 発送先名で受注から最後のデータを取得
function getLastData(comp) {
  // 1. 全データを一度に取得（変更なし）
  const items = getAllRecords('受注');
  const cools = getAllRecords('クール区分');
  const cargos = getAllRecords('荷扱い');
  const invoiceTypes = getAllRecords('送り状種別');
  const deliveryTimes = getAllRecords('配送時間帯');
  
  // 2. マスタデータをMapオブジェクトに変換（高速検索用）
  const coolsMap = new Map();
  const cargosMap = new Map();
  const invoiceTypesMap = new Map();
  const deliveryTimesMap = new Map();
  
  cools.forEach(cool => {
    const key = `${cool['種別']}_${cool['納品方法']}`;
    coolsMap.set(key, cool['種別値']);
  });
  
  cargos.forEach(cargo => {
    const key = `${cargo['種別']}_${cargo['納品方法']}`;
    cargosMap.set(key, cargo['種別値']);
  });
  
  invoiceTypes.forEach(type => {
    const key = `${type['種別']}_${type['納品方法']}`;
    invoiceTypesMap.set(key, type['種別値']);
  });
  
  deliveryTimes.forEach(time => {
    const key = `${time['時間指定']}_${time['納品方法']}`;
    deliveryTimesMap.set(key, time['時間指定値']);
  });
  
  // 3. 発送先名で最後の受注データを検索（逆順で検索して最初に見つかったものを使用）
  let customerItem = null;
  for (let i = items.length - 1; i >= 0; i--) {
    if (items[i]['発送先名'] === comp) {
      customerItem = items[i];
      break;
    }
  }
  
  if (!customerItem) {
    return JSON.stringify([]);
  }
  
  // 4. 該当する受注IDのデータを一度のループで処理
  const targetOrderId = customerItem['受注ID'];
  const records = items
    .filter(item => item['受注ID'] === targetOrderId)
    .map(wVal => {
      // 5. 基本データのコピー（スプレッド構文を使用）
      const rec = {
        '商品分類': wVal['商品分類'],
        '商品名': wVal['商品名'],
        '受注数': wVal['受注数'],
        '販売価格': wVal['販売価格'],
        '発送先名': wVal['発送先名'],
        '発送先郵便番号': wVal['発送先郵便番号'],
        '発送先住所': wVal['発送先住所'],
        '発送先電話番号': wVal['発送先電話番号'],
        '顧客名': wVal['顧客名'],
        '顧客郵便番号': wVal['顧客郵便番号'],
        '顧客住所': wVal['顧客住所'],
        '顧客電話番号': wVal['顧客電話番号'],
        '発送元名': wVal['発送元名'],
        '発送元郵便番号': wVal['発送元郵便番号'],
        '発送元住所': wVal['発送元住所'],
        '発送元電話番号': wVal['発送元電話番号'],
        '納品方法': wVal['納品方法'],
        '納品書': wVal['納品書'],
        '請求書': wVal['請求書'],
        '領収書': wVal['領収書'],
        'パンフ': wVal['パンフ'],
        'レシピ': wVal['レシピ'],
        'その他添付': wVal['その他添付'],
        '代引総額': wVal['代引総額'],
        '代引内税': wVal['代引内税'],
        '発行枚数': wVal['発行枚数'],
        '送り状備考欄': wVal['送り状備考欄'] || '',
        '納品書備考欄': wVal['納品書備考欄'] || '',
        'メモ': wVal['メモ'] || ''
      };
      
      // 6. Mapを使った高速検索
      const deliveryMethod = wVal['納品方法'];
      
      // 配達時間帯
      rec['配達時間帯'] = wVal['配達時間帯'] 
        ? deliveryTimesMap.get(`${wVal['配達時間帯']}_${deliveryMethod}`) || ""
        : "";
      
      // 送り状種別
      rec['送り状種別'] = wVal['送り状種別']
        ? invoiceTypesMap.get(`${wVal['送り状種別']}_${deliveryMethod}`) || ""
        : "";
      
      // クール区分
      rec['クール区分'] = wVal['クール区分']
        ? coolsMap.get(`${wVal['クール区分']}_${deliveryMethod}`) || ""
        : "";
      
      // 荷扱い１
      rec['荷扱い１'] = wVal['荷扱い１']
        ? cargosMap.get(`${wVal['荷扱い１']}_${deliveryMethod}`) || ""
        : "";
      
      // 荷扱い２
      rec['荷扱い２'] = wVal['荷扱い２']
        ? cargosMap.get(`${wVal['荷扱い２']}_${deliveryMethod}`) || ""
        : "";
      
      return rec;
    });
  
  return JSON.stringify(records);
}
// function getLastData(comp) {
//   Logger.log(comp);
//   const searchString = comp;
//   var customerItem = '';
//   const items = getAllRecords('受注');
//   const cools = getAllRecords('クール区分');
//   const cargos = getAllRecords('荷扱い');
//   const invoiceTypes = getAllRecords('送り状種別');
//   const deliveryTimes = getAllRecords('配送時間帯');
//   items.forEach(function (wVal) {
//     // 会社名が同じ
//     if (searchString == wVal['発送先名']) {
//       customerItem = wVal;
//     }
//   });
//   const records = [];
//   if (!customerItem) {
//     return JSON.stringify(records);
//   }
//   items.forEach(function (wVal) {
//     // 受注IDが同じ
//     if (customerItem['受注ID'] == wVal['受注ID']) {
//       const rec = {};
//       rec['商品分類'] = wVal['商品分類'];
//       rec['商品名'] = wVal['商品名'];
//       rec['受注数'] = wVal['受注数'];
//       rec['販売価格'] = wVal['販売価格'];
//       rec['発送先名'] = wVal['発送先名'];
//       rec['発送先郵便番号'] = wVal['発送先郵便番号'];
//       rec['発送先住所'] = wVal['発送先住所'];
//       rec['発送先電話番号'] = wVal['発送先電話番号'];
//       rec['顧客名'] = wVal['顧客名'];
//       rec['顧客郵便番号'] = wVal['顧客郵便番号'];
//       rec['顧客住所'] = wVal['顧客住所'];
//       rec['顧客電話番号'] = wVal['顧客電話番号'];
//       rec['発送元名'] = wVal['発送元名'];
//       rec['発送元郵便番号'] = wVal['発送元郵便番号'];
//       rec['発送元住所'] = wVal['発送元住所'];
//       rec['発送元電話番号'] = wVal['発送元電話番号'];
//       rec['納品方法'] = wVal['納品方法'];
//       if (wVal['配達時間帯']) {
//         deliveryTimes.forEach(function (deliveryTimeVal) {
//           if (deliveryTimeVal['時間指定'] == wVal['配達時間帯'] && deliveryTimeVal['納品方法'] == wVal['納品方法']) {
//             rec['配達時間帯'] = deliveryTimeVal['時間指定値'];
//           }
//         });
//       }
//       else {
//         rec['配達時間帯'] = "";
//       }
//       rec['納品書'] = wVal['納品書'];
//       rec['請求書'] = wVal['請求書'];
//       rec['領収書'] = wVal['領収書'];
//       rec['パンフ'] = wVal['パンフ'];
//       rec['レシピ'] = wVal['レシピ'];
//       rec['その他添付'] = wVal['その他添付'];
//       if (wVal['送り状種別']) {
//         invoiceTypes.forEach(function (invoiceTypeVal) {
//           if (invoiceTypeVal['種別'] == wVal['送り状種別'] && invoiceTypeVal['納品方法'] == wVal['納品方法']) {
//             rec['送り状種別'] = invoiceTypeVal['種別値'];
//           }
//         });
//       }
//       else {
//         rec['送り状種別'] = "";
//       }
//       if (wVal['クール区分']) {
//         cools.forEach(function (coolVal) {
//           if (coolVal['種別'] == wVal['クール区分'] && coolVal['納品方法'] == wVal['納品方法']) {
//             rec['クール区分'] = coolVal['種別値'];
//           }
//         });
//       }
//       else {
//         rec['クール区分'] = "";
//       }
//       if (wVal['荷扱い１']) {
//         cargos.forEach(function (cargoVal) {
//           if (cargoVal['種別'] == wVal['荷扱い１'] && cargoVal['納品方法'] == wVal['納品方法']) {
//             rec['荷扱い１'] = cargoVal['種別値'];
//           }
//         });
//       }
//       else {
//         rec['荷扱い１'] = "";
//       }
//       if (wVal['荷扱い２']) {
//         cargos.forEach(function (cargoVal) {
//           if (cargoVal['種別'] == wVal['荷扱い２'] && cargoVal['納品方法'] == wVal['納品方法']) {
//             rec['荷扱い２'] = cargoVal['種別値'];
//           }
//         });
//       }
//       else {
//         rec['荷扱い２'] = "";
//       }
//       rec['代引総額'] = wVal['代引総額'];
//       rec['代引内税'] = wVal['代引内税'];
//       rec['発行枚数'] = wVal['発行枚数'];
//       records.push(rec);
//     }
//   });
//   Logger.log(records);
//   return JSON.stringify(records);
// }
// 顧客名で見積書データから最後のデータを取得
function getQuotationData(comp) {
  Logger.log(comp);
  const searchString = comp;
  var customerItem = '';
  const items = getAllRecords('見積書データ');
  items.forEach(function (wVal) {
    // 会社名が同じ
    if (searchString == wVal['顧客名']) {
      customerItem = wVal;
    }
  });
  const records = [];
  if (!customerItem) {
    return JSON.stringify(records);
  }
  items.forEach(function (wVal) {
    // 受注IDが同じ
    if (customerItem['見積ID'] == wVal['見積ID']) {
      const rec = {};
      rec['商品分類'] = wVal['商品分類'];
      rec['商品名'] = wVal['商品名'];
      rec['単価（外税）'] = wVal['単価（外税）'];
      records.push(rec);
    }
  });
  Logger.log(records);
  return JSON.stringify(records);
}
// ============================================
// 受注履歴取得関数（修正版）
// 
// 修正: 発送先名の比較時に前後の空白をtrimして比較
// ============================================

/**
 * 発送先名から過去の受注履歴を取得（発送日の降順）
 * @param {string} shippingToName - 発送先名
 * @returns {string} - JSON形式の受注履歴リスト [{orderId, shippingDate, orderDate, productSummary}]
 */
function getOrderHistoryByShippingTo(shippingToName) {
  if (!shippingToName) {
    return JSON.stringify([]);
  }
  
  // 検索キーワードをtrim
  const searchName = shippingToName.trim();
  
  // getAllRecordsを使用（前回商品反映と同じ方式）
  const items = getAllRecords('受注');
  
  // 発送先名が一致するデータを受注ID単位でグループ化
  const orderMap = new Map();
  
  for (let i = 0; i < items.length; i++) {
    const row = items[i];
    const rowName = row['発送先名'];
    
    // 両方trimして比較
    if (rowName && String(rowName).trim() === searchName) {
      const orderId = row['受注ID'];
      
      if (!orderMap.has(orderId)) {
        orderMap.set(orderId, {
          orderId: orderId,
          shippingDate: row['発送日'],
          orderDate: row['受注日'],
          products: []
        });
      }
      
      // 商品名を追加
      const productName = row['商品名'];
      if (productName) {
        orderMap.get(orderId).products.push(productName);
      }
    }
  }
  
  // Mapを配列に変換し、発送日の降順でソート
  const orders = Array.from(orderMap.values())
    .map(order => {
      // 日付フォーマット
      let shippingDateStr = '';
      if (order.shippingDate) {
        try {
          const d = new Date(order.shippingDate);
          if (!isNaN(d.getTime())) {
            shippingDateStr = Utilities.formatDate(d, 'JST', 'yyyy/MM/dd');
          }
        } catch (e) {
          shippingDateStr = String(order.shippingDate);
        }
      }
      
      // 商品サマリー（最大2つ + 他N件）
      let productSummary = '';
      if (order.products.length > 0) {
        const displayProducts = order.products.slice(0, 2);
        productSummary = displayProducts.join(', ');
        if (order.products.length > 2) {
          productSummary += ' 他' + (order.products.length - 2) + '件';
        }
      }
      
      return {
        orderId: order.orderId,
        shippingDate: shippingDateStr,
        shippingDateRaw: order.shippingDate,
        productSummary: productSummary,
        productCount: order.products.length
      };
    })
    .sort((a, b) => {
      // 発送日の降順（新しい順）
      const dateA = a.shippingDateRaw ? new Date(a.shippingDateRaw) : new Date(0);
      const dateB = b.shippingDateRaw ? new Date(b.shippingDateRaw) : new Date(0);
      return dateB - dateA;
    })
    .slice(0, 20); // 最大20件
  
  return JSON.stringify(orders);
}

/**
 * 受注IDから引継ぎ用データを取得（発送日・納品日・個数は除外）
 * @param {string} orderId - 受注ID
 * @returns {string} - JSON形式の受注データ
 */
function getOrderDataForInherit(orderId) {
  if (!orderId) {
    return JSON.stringify(null);
  }
  
  // 既存のgetOrderByOrderId関数を利用
  const data = getOrderByOrderId(orderId);
  
  if (!data) {
    return JSON.stringify(null);
  }
  
  // 引継ぎ用に発送日・納品日をクリア
  data.shippingDate = '';
  data.deliveryDate = '';
  
  // 商品の個数をクリア
  if (data.items) {
    data.items = data.items.map(item => ({
      ...item,
      quantity: ''
    }));
  }
  
  return JSON.stringify(data);
}
