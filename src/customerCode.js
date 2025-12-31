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

// ============================================
// 過去注文比較チェック機能
// 電話対応時の入力ミス防止用
// ============================================

/**
 * 発送先名から直近の注文履歴を取得し、現在の注文と比較して警告を生成
 * @param {string} shippingToName - 発送先名
 * @param {Array} currentItems - 現在の注文商品リスト [{productName, quantity}]
 * @returns {string} - JSON形式の比較結果
 */
function checkOrderAgainstHistory(shippingToName, currentItems) {
  if (!shippingToName || !currentItems || currentItems.length === 0) {
    return JSON.stringify({ hasWarnings: false, warnings: [] });
  }
  
  const searchName = shippingToName.trim();
  const items = getAllRecords('受注');
  
  // 発送先名が一致する過去注文を取得（受注ID単位）
  const orderMap = new Map();
  
  for (let i = 0; i < items.length; i++) {
    const row = items[i];
    const rowName = row['発送先名'];
    
    if (rowName && String(rowName).trim() === searchName) {
      const orderId = row['受注ID'];
      
      if (!orderMap.has(orderId)) {
        orderMap.set(orderId, {
          orderId: orderId,
          shippingDate: row['発送日'],
          items: []
        });
      }
      
      const productName = row['商品名'];
      const quantity = parseFloat(row['個数']) || 0;
      
      if (productName) {
        orderMap.get(orderId).items.push({
          productName: productName,
          quantity: quantity
        });
      }
    }
  }
  
  // 発送日でソートして直近3回を取得
  const recentOrders = Array.from(orderMap.values())
    .sort((a, b) => {
      const dateA = a.shippingDate ? new Date(a.shippingDate) : new Date(0);
      const dateB = b.shippingDate ? new Date(b.shippingDate) : new Date(0);
      return dateB - dateA;
    })
    .slice(0, 3);
  
  if (recentOrders.length === 0) {
    return JSON.stringify({ hasWarnings: false, warnings: [], isNewCustomer: true });
  }
  
  const warnings = [];
  
  // 1. 商品ごとの過去平均数量を計算
  const productStats = new Map();
  recentOrders.forEach(order => {
    order.items.forEach(item => {
      if (!productStats.has(item.productName)) {
        productStats.set(item.productName, { totalQty: 0, count: 0 });
      }
      const stats = productStats.get(item.productName);
      stats.totalQty += item.quantity;
      stats.count += 1;
    });
  });
  
  // 2. 数量異常チェック（±50%以上の差異）
  currentItems.forEach(currentItem => {
    if (!currentItem.productName) return;
    
    const stats = productStats.get(currentItem.productName);
    if (stats && stats.count > 0) {
      const avgQty = stats.totalQty / stats.count;
      const currentQty = parseFloat(currentItem.quantity) || 0;
      
      if (avgQty > 0 && currentQty > 0) {
        const diff = Math.abs(currentQty - avgQty) / avgQty;
        if (diff > 0.5) {
          const direction = currentQty > avgQty ? '多い' : '少ない';
          warnings.push({
            type: 'quantity',
            level: 'warning',
            productName: currentItem.productName,
            message: `${currentItem.productName}: 今回 ${currentQty} （過去平均: ${avgQty.toFixed(1)}）← ${direction}です`,
            currentQty: currentQty,
            avgQty: avgQty
          });
        }
      }
    }
  });
  
  // 3. 常連商品の欠落チェック（3回中2回以上注文している商品）
  productStats.forEach((stats, productName) => {
    if (stats.count >= 2) {  // 直近3回中2回以上注文
      const found = currentItems.some(item => item.productName === productName);
      if (!found) {
        warnings.push({
          type: 'missing',
          level: 'info',
          productName: productName,
          message: `${productName}: 今回なし（過去${stats.count}回連続で注文あり）← 漏れていませんか？`,
          orderCount: stats.count
        });
      }
    }
  });
  
  // 4. 初めての商品チェック（過去に注文したことがない商品）
  currentItems.forEach(currentItem => {
    if (!currentItem.productName) return;
    
    if (!productStats.has(currentItem.productName)) {
      warnings.push({
        type: 'new',
        level: 'info',
        productName: currentItem.productName,
        message: `${currentItem.productName}: この発送先では初めての注文です`,
        currentQty: parseFloat(currentItem.quantity) || 0
      });
    }
  });
  
  return JSON.stringify({
    hasWarnings: warnings.length > 0,
    warnings: warnings,
    recentOrderCount: recentOrders.length,
    isNewCustomer: false
  });
}

/**
 * 発送先名から直近の注文を取得（電話受注テンプレート用）
 * @param {string} shippingToName - 発送先名
 * @returns {string} - JSON形式の直近注文データ
 */
function getRecentOrderForTemplate(shippingToName) {
  if (!shippingToName) {
    return JSON.stringify(null);
  }
  
  const searchName = shippingToName.trim();
  const items = getAllRecords('受注');
  
  // 発送先名が一致する過去注文を取得
  const orderMap = new Map();
  
  for (let i = 0; i < items.length; i++) {
    const row = items[i];
    const rowName = row['発送先名'];
    
    if (rowName && String(rowName).trim() === searchName) {
      const orderId = row['受注ID'];
      
      if (!orderMap.has(orderId)) {
        orderMap.set(orderId, {
          orderId: orderId,
          shippingDate: row['発送日'],
          customerName: row['依頼人名'],
          items: []
        });
      }
      
      const productName = row['商品名'];
      const quantity = parseFloat(row['個数']) || 0;
      const price = parseFloat(row['価格']) || 0;
      const productCategory = row['商品分類'] || '';
      
      if (productName) {
        orderMap.get(orderId).items.push({
          productName: productName,
          quantity: quantity,
          price: price,
          productCategory: productCategory
        });
      }
    }
  }
  
  // 発送日でソートして最新を取得
  const recentOrder = Array.from(orderMap.values())
    .sort((a, b) => {
      const dateA = a.shippingDate ? new Date(a.shippingDate) : new Date(0);
      const dateB = b.shippingDate ? new Date(b.shippingDate) : new Date(0);
      return dateB - dateA;
    })[0];
  
  if (!recentOrder) {
    return JSON.stringify(null);
  }
  
  // 発送日をフォーマット
  let shippingDateStr = '';
  if (recentOrder.shippingDate) {
    try {
      const d = new Date(recentOrder.shippingDate);
      if (!isNaN(d.getTime())) {
        shippingDateStr = Utilities.formatDate(d, 'JST', 'yyyy/MM/dd');
      }
    } catch (e) {
      shippingDateStr = String(recentOrder.shippingDate);
    }
  }
  
  return JSON.stringify({
    orderId: recentOrder.orderId,
    shippingDate: shippingDateStr,
    customerName: recentOrder.customerName,
    items: recentOrder.items
  });
}
// ============================================
// 電話受注モード用関数
// ============================================

/**
 * 電話受注モード用の発送先リストを取得
 * 最終注文日でソートして、よく使う発送先を上位に表示
 * @returns {string} - JSON形式の発送先リスト
 */
function getShippingToListForPhoneOrder() {
  const items = getAllRecords('受注');
  
  // 発送先ごとに集計
  const shippingToMap = new Map();
  
  // 1. まず発送先情報マスタから全発送先を取得（新規顧客対応）
  const shippingToMaster = getAllRecords('発送先情報');
  for (let i = 0; i < shippingToMaster.length; i++) {
    const row = shippingToMaster[i];
    const shippingToName = row['会社名'] || row['氏名'];
    
    if (!shippingToName) continue;
    
    const name = String(shippingToName).trim();
    
    if (!shippingToMap.has(name)) {
      shippingToMap.set(name, {
        shippingToName: name,
        customerName: row['氏名'] || row['会社名'] || '',
        shippingToZipcode: row['郵便番号'] || '',
        shippingToAddress: (row['住所１'] || '') + (row['住所２'] || ''),
        shippingToTel: row['TEL'] || '',
        customerZipcode: row['郵便番号'] || '',
        customerAddress: (row['住所１'] || '') + (row['住所２'] || ''),
        customerTel: row['TEL'] || '',
        lastOrderDate: null,
        orderCount: 0,
        isNewCustomer: true  // 新規顧客フラグ
      });
    }
  }
  
  // 2. 受注履歴から発送先を取得・更新（既存顧客）
  for (let i = 0; i < items.length; i++) {
    const row = items[i];
    const shippingToName = row['発送先名'];
    
    if (!shippingToName) continue;
    
    const name = String(shippingToName).trim();
    
    if (!shippingToMap.has(name)) {
      shippingToMap.set(name, {
        shippingToName: name,
        customerName: row['依頼人名'] || '',
        shippingToZipcode: row['発送先郵便番号'] || '',
        shippingToAddress: row['発送先住所'] || '',
        shippingToTel: row['発送先TEL'] || '',
        customerZipcode: row['郵便番号'] || '',
        customerAddress: row['住所'] || '',
        customerTel: row['TEL'] || '',
        lastOrderDate: null,
        orderCount: 0,
        isNewCustomer: false
      });
    }
    
    const data = shippingToMap.get(name);
    data.orderCount++;
    data.isNewCustomer = false;  // 注文履歴があるので新規ではない
    
    // 最終注文日を更新
    const shippingDate = row['発送日'];
    if (shippingDate) {
      try {
        const d = new Date(shippingDate);
        if (!isNaN(d.getTime())) {
          if (!data.lastOrderDate || d > data.lastOrderDate) {
            data.lastOrderDate = d;
          }
        }
      } catch (e) {
        // 日付パースエラーは無視
      }
    }
  }
  
  // 配列に変換してソート（最終注文日降順、注文回数降順、新規顧客は最後）
  const result = Array.from(shippingToMap.values())
    .sort((a, b) => {
      // 新規顧客は最後に表示
      if (a.isNewCustomer && !b.isNewCustomer) return 1;
      if (!a.isNewCustomer && b.isNewCustomer) return -1;
      
      // まず最終注文日でソート
      const dateA = a.lastOrderDate ? a.lastOrderDate.getTime() : 0;
      const dateB = b.lastOrderDate ? b.lastOrderDate.getTime() : 0;
      if (dateB !== dateA) return dateB - dateA;
      
      // 同日の場合は注文回数でソート
      return b.orderCount - a.orderCount;
    });
  
  // 日付をフォーマット
  result.forEach(item => {
    if (item.lastOrderDate) {
      try {
        item.lastOrderDate = Utilities.formatDate(item.lastOrderDate, 'JST', 'yyyy/MM/dd');
      } catch (e) {
        item.lastOrderDate = '';
      }
    } else {
      item.lastOrderDate = '';
    }
  });
  
  return JSON.stringify(result);
}

/**
 * 電話受注モード用の顧客候補を取得（サジェスト用）
 * @param {string} query - 検索クエリ
 * @returns {string} - JSON形式の顧客リスト
 */
function getCustomerSuggestionsForPhoneOrder(query) {
  if (!query || query.trim().length < 1) {
    return JSON.stringify([]);
  }
  
  const searchQuery = query.trim().toLowerCase();
  const customerMap = new Map();
  
  // 1. 顧客情報マスタから取得
  const customers = getAllRecords('顧客情報');
  for (let i = 0; i < customers.length; i++) {
    const row = customers[i];
    const customerName = row['会社名'] || row['氏名'] || '';
    
    if (!customerName) continue;
    
    const name = String(customerName).trim();
    if (!name.toLowerCase().includes(searchQuery)) continue;
    
    if (!customerMap.has(name)) {
      customerMap.set(name, {
        customerName: name,
        customerZipcode: row['郵便番号'] || '',
        customerAddress: (row['住所１'] || '') + (row['住所２'] || ''),
        customerTel: row['TEL'] || '',
        source: 'master'
      });
    }
  }
  
  // 2. 受注履歴からも取得（マスタにない顧客がいる可能性）
  const orders = getAllRecords('受注');
  for (let i = 0; i < orders.length; i++) {
    const row = orders[i];
    const customerName = row['依頼人名'] || '';
    
    if (!customerName) continue;
    
    const name = String(customerName).trim();
    if (!name.toLowerCase().includes(searchQuery)) continue;
    
    if (!customerMap.has(name)) {
      customerMap.set(name, {
        customerName: name,
        customerZipcode: row['郵便番号'] || '',
        customerAddress: row['住所'] || '',
        customerTel: row['TEL'] || '',
        source: 'order'
      });
    }
  }
  
  // 配列に変換して上位10件を返す
  const result = Array.from(customerMap.values()).slice(0, 10);
  
  return JSON.stringify(result);
}

/**
 * 発送先名から紐づく顧客情報を取得（自動入力用）
 * @param {string} shippingToName - 発送先名
 * @returns {string} - JSON形式の顧客情報
 */
function getCustomerByShippingTo(shippingToName) {
  if (!shippingToName) {
    return JSON.stringify(null);
  }
  
  const searchName = shippingToName.trim();
  const orders = getAllRecords('受注');
  
  // 受注履歴から発送先に紐づく顧客を検索（最新を優先）
  for (let i = orders.length - 1; i >= 0; i--) {
    const row = orders[i];
    const rowShippingTo = row['発送先名'];
    
    if (rowShippingTo && String(rowShippingTo).trim() === searchName) {
      return JSON.stringify({
        customerName: row['依頼人名'] || '',
        customerZipcode: row['郵便番号'] || '',
        customerAddress: row['住所'] || '',
        customerTel: row['TEL'] || ''
      });
    }
  }
  
  // 受注履歴にない場合は発送先マスタから探す
  const shippingToMaster = getAllRecords('発送先情報');
  for (let i = 0; i < shippingToMaster.length; i++) {
    const row = shippingToMaster[i];
    const name = row['会社名'] || row['氏名'] || '';
    
    if (String(name).trim() === searchName) {
      // 発送先マスタの情報を顧客情報として返す
      return JSON.stringify({
        customerName: row['氏名'] || row['会社名'] || '',
        customerZipcode: row['郵便番号'] || '',
        customerAddress: (row['住所１'] || '') + (row['住所２'] || ''),
        customerTel: row['TEL'] || ''
      });
    }
  }
  
  return JSON.stringify(null);
}

/**
 * 電話受注の仮受注を登録（要確認フラグ付き）
 * @param {Object} orderData - 受注データ
 * @returns {string} - JSON形式の結果
 */
function createPhoneOrderDraft(orderData) {
  try {
    const data = typeof orderData === 'string' ? JSON.parse(orderData) : orderData;
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('仮受注');
    
    // 仮受注シートがなければ作成
    if (!sheet) {
      sheet = ss.insertSheet('仮受注');
      const headers = [
        '仮受注ID', 'ステータス', '登録日時', '要確認フラグ',
        '発送先名', '発送先TEL', '発送先郵便番号', '発送先住所',
        '顧客名', '顧客TEL', '顧客郵便番号', '顧客住所',
        '発送日', '納品日', '商品情報', 'メモ', '受付方法'
      ];
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    }
    
    const tempOrderId = 'phone_' + Utilities.getUuid().substring(0, 8);
    const now = Utilities.formatDate(new Date(), 'JST', 'yyyy/MM/dd HH:mm:ss');
    
    // 商品情報をJSON文字列化
    const itemsJson = JSON.stringify(data.items || []);
    
    const row = [
      tempOrderId,
      '未処理',
      now,
      data.needsConfirmation ? '要確認' : '',
      data.shippingToName || '',
      data.shippingToTel || '',
      data.shippingToZipcode || '',
      data.shippingToAddress || '',
      data.customerName || '',
      data.customerTel || '',
      data.customerZipcode || '',
      data.customerAddress || '',
      data.shippingDate || '',
      data.deliveryDate || '',
      itemsJson,
      data.memo || '',
      '電話'
    ];
    
    sheet.appendRow(row);
    
    return JSON.stringify({
      success: true,
      tempOrderId: tempOrderId,
      message: '仮受注を登録しました'
    });
  } catch (e) {
    return JSON.stringify({
      success: false,
      error: e.toString()
    });
  }
}
