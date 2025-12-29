/**
 * 受注一覧データを取得
 * @param {Object} params - 検索パラメータ {dateFrom, dateTo, destination, customer}
 * @returns {string} - JSON文字列
 */
function getOrderListData(params) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('受注');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  // 列インデックスを取得
  const getColIndex = (name) => headers.indexOf(name);
  
  const orderIdCol = getColIndex('受注ID');
  const orderDateCol = getColIndex('受注日');
  const shippingDateCol = getColIndex('発送日');
  const deliveryDateCol = getColIndex('納品日');
  const shippingToNameCol = getColIndex('発送先名');
  const customerNameCol = getColIndex('顧客名');
  const sendProductCol = getColIndex('品名');
  const productCol = getColIndex('商品名');
  const quantityCol = getColIndex('受注数');
  const priceCol = getColIndex('販売価格');
  const bunruiCol = getColIndex('商品分類');
  
  // パラメータ
  const startDate = params.dateFrom ? new Date(params.dateFrom) : null;
  const endDate = params.dateTo ? new Date(params.dateTo) : null;
  const searchDestination = params.destination || '';
  const searchCustomer = params.customer || '';
  
  // endDateは終日を含むように調整
  if (endDate) {
    endDate.setHours(23, 59, 59, 999);
  }
  if (startDate) {
    startDate.setHours(0, 0, 0, 0);
  }
  
  // 除外する商品分類・商品名
  const excludeCategories = ['送料', '手数料', '追加料金'];
  const excludeProducts = ['送料', 'クール便追加', '代引手数料', '梱包料'];
  
  // 受注IDでグループ化
  const orderMap = new Map();
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const orderId = row[orderIdCol];
    
    if (!orderId) continue;
    
    // 発送日でフィルタ
    const shippingDateRaw = row[shippingDateCol];
    if (shippingDateRaw) {
      const shipDate = new Date(shippingDateRaw);
      shipDate.setHours(0, 0, 0, 0);
      if (startDate && shipDate < startDate) continue;
      if (endDate && shipDate > endDate) continue;
    }
    
    // 発送先フィルタ
    const shippingToName = row[shippingToNameCol] || '';
    if (searchDestination && !shippingToName.includes(searchDestination)) continue;
    
    // 顧客フィルタ
    const customerName = row[customerNameCol] || '';
    if (searchCustomer && !customerName.includes(searchCustomer)) continue;
    
    // グループ化（初回のみ基本情報を設定）
    if (!orderMap.has(orderId)) {
      orderMap.set(orderId, {
        orderId: orderId,
        orderDate: row[orderDateCol],
        shippingDate: row[shippingDateCol],
        shippingDateRaw: shippingDateRaw,  // ソート用
        deliveryDate: row[deliveryDateCol],
        customerName: customerName,
        destinationName: shippingToName,   // HTML側の名前に合わせる
        productName: row[sendProductCol] || '',  // 品名（HTML側の名前に合わせる）
        quantity: 0,  // 商品数（後で計算）
        totalAmount: 0
      });
    }
    
    // 商品情報を集計（除外対象以外）
    const bunrui = row[bunruiCol] || '';
    const product = row[productCol] || '';
    const qty = Number(row[quantityCol]) || 0;
    const price = Number(row[priceCol]) || 0;
    
    if (!excludeCategories.includes(bunrui) && !excludeProducts.includes(product)) {
      const order = orderMap.get(orderId);
      order.quantity += 1;  // 商品行数をカウント
      order.totalAmount += price * qty;
    }
  }
  
  // Mapを配列に変換
  let orders = Array.from(orderMap.values());
  
  // ソート：発送日昇順 → 顧客名順
  orders.sort((a, b) => {
    const dateA = a.shippingDateRaw ? new Date(a.shippingDateRaw) : new Date(0);
    const dateB = b.shippingDateRaw ? new Date(b.shippingDateRaw) : new Date(0);
    if (dateA.getTime() !== dateB.getTime()) {
      return dateA - dateB;
    }
    return (a.customerName || '').localeCompare(b.customerName || '');
  });
  
  // 日付をフォーマット & shippingDateRawを削除
  orders = orders.map(order => {
    return {
      orderId: order.orderId,
      orderDate: formatDateForDisplay(order.orderDate),
      shippingDate: formatDateForDisplay(order.shippingDate),
      deliveryDate: formatDateForDisplay(order.deliveryDate),
      customerName: order.customerName,
      destinationName: order.destinationName,
      productName: order.productName,
      quantity: order.quantity,
      totalAmount: order.totalAmount
    };
  });
  
  return JSON.stringify({
    orders: orders,
    totalCount: orders.length
  });
}

/**
 * 日付を表示用フォーマットに変換（MM/dd形式ではなくISO形式で返す）
 * HTML側のformatDisplayDateで変換するため
 */
function formatDateForDisplay(dateValue) {
  if (!dateValue) return '';
  
  try {
    let date;
    if (dateValue instanceof Date) {
      date = dateValue;
    } else if (typeof dateValue === 'string') {
      date = new Date(dateValue.replace(/\//g, '-'));
    } else {
      return '';
    }
    
    if (isNaN(date.getTime())) return '';
    
    // yyyy-MM-dd形式で返す（HTML側でMM/dd形式に変換）
    return Utilities.formatDate(date, 'JST', 'yyyy-MM-dd');
  } catch (e) {
    return '';
  }
}

/**
 * 発送先検索（モーダル用）
 */
function searchShippingTo(keyword) {
  if (!keyword || keyword.length < 2) return [];
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('発送先情報');
  if (!sheet) return [];
  
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  const companyCol = headers.indexOf('会社名');
  const personCol = headers.indexOf('氏名');
  const addressCol = headers.indexOf('住所１');
  const telCol = headers.indexOf('TEL');
  
  const results = [];
  const keywordLower = keyword.toLowerCase();
  
  for (let i = 1; i < data.length && results.length < 20; i++) {
    const row = data[i];
    const company = row[companyCol] || '';
    const person = row[personCol] || '';
    const address = row[addressCol] || '';
    const tel = row[telCol] || '';
    
    const name = [company, person].filter(Boolean).join('　');
    const searchTarget = (name + address + tel).toLowerCase();
    
    if (searchTarget.includes(keywordLower)) {
      results.push({
        name: name,
        address: address,
        tel: tel
      });
    }
  }
  
  return results;
}

/**
 * 顧客検索（モーダル用）
 */
function searchCustomer(keyword) {
  if (!keyword || keyword.length < 2) return [];
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('顧客情報');
  if (!sheet) return [];
  
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  const companyCol = headers.indexOf('会社名');
  const personCol = headers.indexOf('氏名');
  const displayCol = headers.indexOf('表示名');
  const addressCol = headers.indexOf('住所１');
  const telCol = headers.indexOf('TEL');
  
  const results = [];
  const keywordLower = keyword.toLowerCase();
  
  for (let i = 1; i < data.length && results.length < 20; i++) {
    const row = data[i];
    const company = row[companyCol] || '';
    const person = row[personCol] || '';
    const display = row[displayCol] || '';
    const address = row[addressCol] || '';
    const tel = row[telCol] || '';
    
    const name = display || [company, person].filter(Boolean).join('　');
    const searchTarget = (name + address + tel).toLowerCase();
    
    if (searchTarget.includes(keywordLower)) {
      results.push({
        name: name,
        address: address,
        tel: tel
      });
    }
  }
  
  return results;
}
