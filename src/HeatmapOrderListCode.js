// ============================================
// 発注一覧（ヒートマップ）用 GAS関数
// 略称はシート「商品」の「略称」列から取得
// ============================================

/**
 * 商品分類一覧を取得
 * シート「商品分類」の列構成: 分類ID, 商品分類
 */
function getProductCategories() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('商品分類');
  const values = sheet.getDataRange().getValues();
  const headers = values.shift(); // ヘッダー除去
  
  // 列インデックス取得
  const colIndex = {};
  headers.forEach((header, index) => {
    colIndex[header] = index;
  });
  
  // 「商品分類」列を返す（列名がない場合は2列目）
  const categoryColIndex = colIndex['商品分類'] !== undefined ? colIndex['商品分類'] : 1;
  
  return values.map(row => row[categoryColIndex]).filter(v => v);
}

/**
 * 発送先検索
 * @param {string} keyword - 検索キーワード
 * @returns {Array} 検索結果
 */
function searchShippingTo(keyword) {
  const records = getAllRecords('発送先情報');
  const results = [];
  const lowerKeyword = keyword.toLowerCase();
  
  for (const record of records) {
    const name = record['会社名'] || '';
    const personName = record['氏名'] || '';
    const address = (record['住所１'] || '') + (record['住所２'] || '');
    const tel = record['TEL'] || '';
    const zipcode = record['郵便番号'] || '';
    
    const searchText = `${name}${personName}${address}${tel}${zipcode}`.toLowerCase();
    
    if (searchText.includes(lowerKeyword)) {
      results.push({
        name: name + (personName ? '　' + personName : ''),
        address: address,
        tel: tel
      });
    }
    
    if (results.length >= 20) break;
  }
  
  return results;
}

/**
 * 顧客検索
 * @param {string} keyword - 検索キーワード
 * @returns {Array} 検索結果
 */
function searchCustomer(keyword) {
  const records = getAllRecords('顧客情報');
  const results = [];
  const lowerKeyword = keyword.toLowerCase();
  
  for (const record of records) {
    const displayName = record['表示名'] || '';
    const companyName = record['会社名'] || '';
    const personName = record['氏名'] || '';
    const address = (record['住所１'] || '') + (record['住所２'] || '');
    const tel = record['TEL'] || '';
    const fax = record['FAX'] || '';
    const zipcode = record['郵便番号'] || '';
    
    const searchText = `${displayName}${companyName}${personName}${address}${tel}${fax}${zipcode}`.toLowerCase();
    
    if (searchText.includes(lowerKeyword)) {
      results.push({
        name: displayName || companyName + (personName ? '　' + personName : ''),
        address: address,
        tel: tel
      });
    }
    
    if (results.length >= 20) break;
  }
  
  return results;
}

/**
 * ヒートマップ用データ取得
 * @param {Object} params - 検索パラメータ
 * @returns {string} JSON形式のヒートマップデータ
 */
function getHeatmapData(params) {
  const { dateFrom, dateTo, categories, destination, customer } = params;
  
  // 日付範囲を生成
  const dates = generateDateRange(dateFrom, dateTo);
  
  // 商品マスタ取得（略称列を使用）
  const productMaster = getProductMasterWithAbbreviation(categories);
  
  // 受注データ取得
  const orders = getOrdersForHeatmap(dateFrom, dateTo, categories, destination, customer);
  
  // ヒートマップデータを構築（データがある商品のみ）
  const heatmapResult = buildHeatmapData(dates, productMaster, orders, true);
  
  return JSON.stringify(heatmapResult);
}

/**
 * 日付範囲を生成
 */
function generateDateRange(fromStr, toStr) {
  const dates = [];
  const from = new Date(fromStr);
  const to = new Date(toStr);
  
  const current = new Date(from);
  while (current <= to) {
    dates.push(Utilities.formatDate(current, 'JST', 'yyyy-MM-dd'));
    current.setDate(current.getDate() + 1);
  }
  
  return dates;
}

/**
 * 商品マスタ取得（略称列対応）
 * シート「商品」から「商品名」「略称」「商品分類」を取得
 */
function getProductMasterWithAbbreviation(categoryFilter) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('商品');
  const values = sheet.getDataRange().getValues();
  const headers = values.shift();
  
  // 列インデックス取得
  const colIndex = {};
  headers.forEach((header, index) => {
    colIndex[header] = index;
  });
  
  const result = [];
  const addedProducts = new Set(); // 重複防止用
  
  for (const row of values) {
    const productName = row[colIndex['商品名']] || '';
    const abbreviation = row[colIndex['略称']] || ''; // 略称列
    const category = row[colIndex['商品分類']] || '';
    
    // 空行スキップ
    if (!productName) continue;
    
    // 重複スキップ
    if (addedProducts.has(productName)) continue;
    addedProducts.add(productName);
    
    // カテゴリーフィルタ
    if (categoryFilter && categoryFilter.length > 0) {
      if (!categoryFilter.includes(category)) continue;
    }
    
    // 送料・手数料など除外
    if (category === '送料・代引手数料') continue;
    
    result.push({
      fullName: productName,
      shortName: abbreviation || productName, // 略称がなければ商品名をそのまま使用
      category: category
    });
  }
  
  return result;
}

/**
 * ヒートマップ用の受注データ取得
 */
function getOrdersForHeatmap(dateFrom, dateTo, categories, destination, customer) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('受注');
  const values = sheet.getDataRange().getValues();
  const labels = values.shift();
  
  // 列インデックス取得
  const colIndex = {};
  labels.forEach((label, index) => {
    colIndex[label] = index;
  });
  
  const startDate = new Date(dateFrom);
  const endDate = new Date(dateTo);
  endDate.setHours(23, 59, 59);
  
  const orders = [];
  
  for (const row of values) {
    const shippingDateRaw = row[colIndex['発送日']];
    if (!shippingDateRaw) continue;
    
    const shippingDate = new Date(shippingDateRaw);
    
    // 日付範囲チェック
    if (shippingDate < startDate || shippingDate > endDate) continue;
    
    const productName = row[colIndex['商品名']] || '';
    const productCategory = row[colIndex['商品分類']] || '';
    const destinationName = row[colIndex['発送先名']] || '';
    const customerName = row[colIndex['顧客名']] || '';
    const quantity = Number(row[colIndex['受注数']]) || 0;
    
    // 送料など除外
    if (productName === '送料' || productName === '代引手数料') continue;
    if (productCategory === '送料・代引手数料') continue;
    
    // カテゴリーフィルタ
    if (categories && categories.length > 0) {
      if (!categories.includes(productCategory)) continue;
    }
    
    // 発送先フィルタ
    if (destination) {
      if (!destinationName.includes(destination)) continue;
    }
    
    // 顧客フィルタ
    if (customer) {
      if (!customerName.includes(customer)) continue;
    }
    
    orders.push({
      date: Utilities.formatDate(shippingDate, 'JST', 'yyyy-MM-dd'),
      productName: productName,
      category: productCategory,
      destination: destinationName,
      customer: customerName,
      quantity: quantity
    });
  }
  
  return orders;
}

/**
 * ヒートマップデータを構築
 * @param {Array} dates - 日付配列
 * @param {Array} products - 商品マスタ配列
 * @param {Array} orders - 受注データ配列
 * @param {boolean} onlyWithData - trueの場合、データがある商品のみ返す
 */
function buildHeatmapData(dates, products, orders, onlyWithData) {
  // 日付→インデックスのマップ
  const dateIndexMap = {};
  dates.forEach((date, index) => {
    dateIndexMap[date] = index;
  });
  
  // 商品名→インデックスのマップ
  const productIndexMap = {};
  products.forEach((product, index) => {
    productIndexMap[product.fullName] = index;
  });
  
  // データ配列初期化
  const data = products.map(() => dates.map(() => 0));
  const details = products.map(() => dates.map(() => []));
  
  // 受注データを集計
  for (const order of orders) {
    const dateIndex = dateIndexMap[order.date];
    const productIndex = productIndexMap[order.productName];
    
    if (dateIndex === undefined || productIndex === undefined) continue;
    
    data[productIndex][dateIndex] += order.quantity;
    details[productIndex][dateIndex].push({
      customer: order.customer,
      destination: order.destination,
      quantity: order.quantity
    });
  }
  
  // 商品別合計算出
  const productTotals = products.map((product, pIndex) => {
    return dates.reduce((sum, date, dIndex) => sum + data[pIndex][dIndex], 0);
  });
  
  // データがある商品のみフィルタリング
  let filteredProducts = products;
  let filteredData = data;
  let filteredDetails = details;
  let filteredProductTotals = productTotals;
  
  if (onlyWithData) {
    const indices = [];
    productTotals.forEach((total, index) => {
      if (total > 0) {
        indices.push(index);
      }
    });
    
    filteredProducts = indices.map(i => products[i]);
    filteredData = indices.map(i => data[i]);
    filteredDetails = indices.map(i => details[i]);
    filteredProductTotals = indices.map(i => productTotals[i]);
  }
  
  // 日計算出（フィルタ後のデータで計算）
  const dailyTotals = dates.map((date, dIndex) => {
    return filteredProducts.reduce((sum, product, pIndex) => sum + filteredData[pIndex][dIndex], 0);
  });
  
  return {
    dates: dates,
    products: filteredProducts,  // { fullName, shortName, category }
    data: filteredData,
    details: filteredDetails,
    dailyTotals: dailyTotals,
    productTotals: filteredProductTotals
  };
}
