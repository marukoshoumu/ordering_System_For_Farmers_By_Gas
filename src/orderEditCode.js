// ============================================
// 修正版: getOrderByOrderId関数
// 
// 修正内容:
// - 納品方法（ヤマト/佐川等）に応じた配達時間帯・送り状種別・クール区分・荷扱いの逆引き
// - マスタの「納品方法」列でフィルタリングして正しい種別値を取得
// ============================================

/**
 * 受注IDから受注データを取得する
 * @param {string} orderId - 受注ID
 * @returns {Object|null} - 受注データオブジェクト
 */
function getOrderByOrderId(orderId) {
  if (!orderId) return null;
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('受注');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  // ヘッダーからインデックスを取得
  const getColIndex = (name) => headers.indexOf(name);
  
  // 受注IDが一致する行を全て取得
  const matchingRows = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][getColIndex('受注ID')] === orderId) {
      matchingRows.push(data[i]);
    }
  }
  
  if (matchingRows.length === 0) return null;
  
  const firstRow = matchingRows[0];
  
  // 納品方法を先に取得（ヤマト/佐川の判定に使用）
  const deliveryMethod = firstRow[getColIndex('納品方法')] || '';
  
  // マスタデータを取得
  const master = getMasterDataCached();
  
  // ============================================
  // 納品方法でフィルタリングした逆引きマップを作成
  // 表示名（種別）→ 種別値 の変換
  // ============================================
  
  // 送り状種別の逆引き（種別 → 種別値）
  const invoiceTypeMap = {};
  master.invoiceTypes
    .filter(item => item['納品方法'] === deliveryMethod)
    .forEach(item => {
      invoiceTypeMap[item['種別']] = item['種別値'];
    });
  
  // クール区分の逆引き（種別 → 種別値）
  const coolClsMap = {};
  master.coolClss
    .filter(item => item['納品方法'] === deliveryMethod)
    .forEach(item => {
      coolClsMap[item['種別']] = item['種別値'];
    });
  
  // 荷扱いの逆引き（種別 → 種別値）
  const cargoMap = {};
  master.cargos
    .filter(item => item['納品方法'] === deliveryMethod)
    .forEach(item => {
      cargoMap[item['種別']] = item['種別値'];
    });
  
  // 配達時間帯の逆引き（時間指定 → 時間指定値）
  const deliveryTimeMap = {};
  master.deliveryTimes
    .filter(item => item['納品方法'] === deliveryMethod)
    .forEach(item => {
      deliveryTimeMap[item['時間指定']] = item['時間指定値'];
    });
  
  // ============================================
  // 受注データを取得
  // ============================================
  
  const result = {
    orderId: orderId,
    orderDate: firstRow[getColIndex('受注日')],
    
    // 発送先情報
    shippingToName: firstRow[getColIndex('発送先名')] || '',
    shippingToZipcode: firstRow[getColIndex('発送先郵便番号')] || '',
    shippingToAddress: firstRow[getColIndex('発送先住所')] || '',
    shippingToTel: firstRow[getColIndex('発送先電話番号')] || '',
    
    // 顧客情報
    customerName: firstRow[getColIndex('顧客名')] || '',
    customerZipcode: firstRow[getColIndex('顧客郵便番号')] || '',
    customerAddress: firstRow[getColIndex('顧客住所')] || '',
    customerTel: firstRow[getColIndex('顧客電話番号')] || '',
    
    // 発送元情報
    shippingFromName: firstRow[getColIndex('発送元名')] || '',
    shippingFromZipcode: firstRow[getColIndex('発送元郵便番号')] || '',
    shippingFromAddress: firstRow[getColIndex('発送元住所')] || '',
    shippingFromTel: firstRow[getColIndex('発送元電話番号')] || '',
    
    // 日程情報
    shippingDate: formatDateForInput(firstRow[getColIndex('発送日')]),
    deliveryDate: formatDateForInput(firstRow[getColIndex('納品日')]),
    
    // 受付情報
    receiptWay: firstRow[getColIndex('受付方法')] || '',
    recipient: firstRow[getColIndex('受付者')] || '',
    deliveryMethod: deliveryMethod,
    
    // 配達時間帯（表示名から時間指定値に変換）
    // 受注シートには「午前中」等の表示名が保存されている → 「0812:午前中」形式に変換
    deliveryTime: deliveryTimeMap[firstRow[getColIndex('配達時間帯')]] || '',
    
    // チェックリスト
    checklist: {
      deliverySlip: firstRow[getColIndex('納品書')] === '○',
      bill: firstRow[getColIndex('請求書')] === '○',
      receipt: firstRow[getColIndex('領収書')] === '○',
      pamphlet: firstRow[getColIndex('パンフ')] === '○',
      recipe: firstRow[getColIndex('レシピ')] === '○'
    },
    otherAttach: firstRow[getColIndex('その他添付')] || '',
    
    // 発送情報（表示名から種別値に変換）
    // 受注シートには「発払い」等の表示名が保存されている → 「0:発払い」形式に変換
    sendProduct: firstRow[getColIndex('品名')] || '',
    invoiceType: invoiceTypeMap[firstRow[getColIndex('送り状種別')]] || '',
    coolCls: coolClsMap[firstRow[getColIndex('クール区分')]] || '',
    cargo1: cargoMap[firstRow[getColIndex('荷扱い１')]] || '',
    cargo2: cargoMap[firstRow[getColIndex('荷扱い２')]] || '',
    
    // 代引・発行枚数
    cashOnDelivery: firstRow[getColIndex('代引総額')] || '',
    cashOnDeliTax: firstRow[getColIndex('代引内税')] || '',
    copiePrint: firstRow[getColIndex('発行枚数')] || '',
    
    // 備考
    csvmemo: firstRow[getColIndex('送り状備考欄')] || '',
    deliveryMemo: firstRow[getColIndex('納品書備考欄')] || '',
    memo: firstRow[getColIndex('メモ')] || '',
    
    // 商品情報（複数行）
    items: []
  };
  
  // 商品情報を取得
  matchingRows.forEach(row => {
    const bunrui = row[getColIndex('商品分類')];
    const product = row[getColIndex('商品名')];
    const quantity = row[getColIndex('受注数')];
    const price = row[getColIndex('販売価格')];
    
    if (product) {
      result.items.push({
        bunrui: bunrui || '',
        product: product || '',
        quantity: quantity || 0,
        price: price || 0
      });
    }
  });
  
  return result;
}

/**
 * 日付をinput[type="date"]用のフォーマットに変換
 */
function formatDateForInput(dateValue) {
  if (!dateValue) return '';
  
  try {
    let date;
    if (dateValue instanceof Date) {
      date = dateValue;
    } else if (typeof dateValue === 'string') {
      // yyyy/MM/dd または yyyy-MM-dd 形式を処理
      date = new Date(dateValue.replace(/\//g, '-'));
    } else {
      return '';
    }
    
    if (isNaN(date.getTime())) return '';
    
    return Utilities.formatDate(date, 'JST', 'yyyy-MM-dd');
  } catch (e) {
    Logger.log('日付変換エラー: ' + e.message);
    return '';
  }
}

/**
 * 受注IDで受注データを削除
 * @param {string} orderId - 受注ID
 * @returns {number} - 削除した行数
 */
function deleteOrderByOrderId(orderId) {
  if (!orderId) return 0;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('受注');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const orderIdCol = headers.indexOf('受注ID');

  if (orderIdCol === -1) {
    Logger.log('受注ID列が見つかりません');
    return 0;
  }

  // 削除対象の行を特定（後ろから削除するため逆順で収集）
  const rowsToDelete = [];
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][orderIdCol] === orderId) {
      rowsToDelete.push(i + 1); // シートの行番号（1-indexed）
    }
  }

  // 後ろから削除
  rowsToDelete.forEach(rowNum => {
    sheet.deleteRow(rowNum);
  });

  Logger.log('削除した行数: ' + rowsToDelete.length + ' (受注ID: ' + orderId + ')');
  return rowsToDelete.length;
}

/**
 * 受注IDでヤマトCSVデータを削除
 * @param {string} orderId - 受注ID
 * @returns {number} - 削除した行数
 */
function deleteYamatoCSVByOrderId(orderId) {
  if (!orderId) return 0;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('ヤマトCSV');

  if (!sheet) {
    Logger.log('ヤマトCSVシートが見つかりません');
    return 0;
  }

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const customerMgmtNoCol = headers.indexOf('お客様管理番号');

  if (customerMgmtNoCol === -1) {
    Logger.log('お客様管理番号列が見つかりません');
    return 0;
  }

  // 削除対象の行を特定（後ろから削除するため逆順で収集）
  const rowsToDelete = [];
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][customerMgmtNoCol] === orderId) {
      rowsToDelete.push(i + 1); // シートの行番号（1-indexed）
    }
  }

  // 後ろから削除
  rowsToDelete.forEach(rowNum => {
    sheet.deleteRow(rowNum);
  });

  Logger.log('ヤマトCSV削除: ' + rowsToDelete.length + '件 (受注ID: ' + orderId + ')');
  return rowsToDelete.length;
}

/**
 * 受注IDで佐川CSVデータを削除
 * @param {string} orderId - 受注ID
 * @returns {number} - 削除した行数
 */
function deleteSagawaCSVByOrderId(orderId) {
  if (!orderId) return 0;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('佐川CSV');

  if (!sheet) {
    Logger.log('佐川CSVシートが見つかりません');
    return 0;
  }

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const customerMgmtNoCol = headers.indexOf('お客様管理番号');

  if (customerMgmtNoCol === -1) {
    Logger.log('お客様管理番号列が見つかりません');
    return 0;
  }

  // 削除対象の行を特定（後ろから削除するため逆順で収集）
  const rowsToDelete = [];
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][customerMgmtNoCol] === orderId) {
      rowsToDelete.push(i + 1); // シートの行番号（1-indexed）
    }
  }

  // 後ろから削除
  rowsToDelete.forEach(rowNum => {
    sheet.deleteRow(rowNum);
  });

  Logger.log('佐川CSV削除: ' + rowsToDelete.length + '件 (受注ID: ' + orderId + ')');
  return rowsToDelete.length;
}
