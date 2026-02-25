/**
 * 受注一覧データを取得
 * @param {Object} params - 検索パラメータ {dateFrom, dateTo, destination, customer, shippingStatus, includeOverdue}
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
  const shippedCol = getColIndex('出荷済');
  const trackingCol = getColIndex('追跡番号');
  const statusCol = getColIndex('ステータス');
  const deliveryMethodCol = getColIndex('納品方法');
  const memoCol = getColIndex('メモ');

  // パラメータ
  const startDate = params.dateFrom ? new Date(params.dateFrom) : null;
  const endDate = params.dateTo ? new Date(params.dateTo) : null;
  const searchDestination = params.destination || '';
  const searchCustomer = params.customer || '';
  const shippingStatus = params.shippingStatus || 'all';
  const includeOverdue = params.includeOverdue || false;  // 予定日超過を含むかどうか

  // endDateは終日を含むように調整
  if (endDate) {
    endDate.setHours(23, 59, 59, 999);
  }
  if (startDate) {
    startDate.setHours(0, 0, 0, 0);
  }

  // 除外する商品分類・商品名
  const excludeCategories = ['送料', '手数料', '追加料金', '送料・代引手数料'];
  const excludeProducts = ['送料', 'クール便追加', '代引手数料', '梱包料'];

  // 予定日超過判定用の今日の日付
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  // 受注IDでグループ化
  const orderMap = new Map();

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const orderId = row[orderIdCol];

    if (!orderId) continue;

    // 発送日でフィルタ（予定日超過を含む場合は過去のデータも含める）
    const shippingDateRaw = row[shippingDateCol];
    const isShippedRow = row[shippedCol] === '○' || row[shippedCol] === true;
    const isCancelledRow = statusCol >= 0 && (row[statusCol] === 'キャンセル' || row[statusCol] === 'cancelled');

    if (shippingDateRaw) {
      const shipDate = new Date(shippingDateRaw);
      shipDate.setHours(0, 0, 0, 0);

      // 予定日超過かどうかを判定
      const isOverdueRow = !isShippedRow && !isCancelledRow && shipDate < today;

      // includeOverdueがtrueで予定日超過の場合は期間条件を無視
      if (!(includeOverdue && isOverdueRow)) {
        if (startDate && shipDate < startDate) continue;
        if (endDate && shipDate > endDate) continue;
      }
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
        shipped: row[shippedCol] || '',  // 出荷済フラグ
        trackingNumber: trackingCol >= 0 ? (row[trackingCol] || '') : '', // 追跡番号
        status: statusCol >= 0 ? (row[statusCol] || '') : '',  // ステータス
        deliveryMethod: deliveryMethodCol >= 0 ? (row[deliveryMethodCol] || '') : '', // 納品方法
        memo: memoCol >= 0 ? (row[memoCol] || '') : '', // メモ
        items: [],  // 商品ごとの配列
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
      // 商品情報を配列に追加
      order.items.push({
        product: product,
        quantity: qty
      });
      order.totalAmount += price * qty;
    }
  }

  // Mapを配列に変換
  let orders = Array.from(orderMap.values());

  // 出荷状態でフィルタリング
  if (shippingStatus !== 'all' || includeOverdue) {
    const today = new Date();
    today.setHours(0, 0, 0, 0);

    orders = orders.filter(order => {
      const isShipped = order.shipped === '○' || order.shipped === true;
      const isCancelled = order.status === 'キャンセル' || order.status === 'cancelled';
      const shippingDate = order.shippingDateRaw ? new Date(order.shippingDateRaw) : null;
      const isOverdue = !isShipped && !isCancelled && shippingDate && shippingDate < today;

      // 基本条件による判定
      let matchesBasicCondition = false;

      if (shippingStatus === 'all') {
        matchesBasicCondition = true;
      } else if (shippingStatus === 'shipped') {
        matchesBasicCondition = isShipped;
      } else if (shippingStatus === 'notShipped') {
        matchesBasicCondition = !isShipped && !isCancelled;
      } else if (shippingStatus === 'cancelled') {
        matchesBasicCondition = isCancelled;
      } else if (shippingStatus === 'harvestWaiting') {
        matchesBasicCondition = order.status === '収穫待ち';
      } else if (shippingStatus === 'shippedToDelivering') {
        matchesBasicCondition = order.status === '出荷済' || order.status === '配達中';
      }

      // includeOverdueがtrueの場合、基本条件 OR 予定日超過でマッチ
      if (includeOverdue) {
        return matchesBasicCondition || isOverdue;
      }

      // includeOverdueがfalseの場合、基本条件のみ
      return matchesBasicCondition;
    });
  }

  // ソート：予定日超過優先 → 発送日昇順 → 顧客名順
  orders.sort((a, b) => {
    const today = new Date();
    today.setHours(0, 0, 0, 0);

    // 予定日超過判定
    const isShippedA = a.shipped === '○' || a.shipped === true;
    const isCancelledA = a.status === 'キャンセル' || a.status === 'cancelled';
    const shippingDateA = a.shippingDateRaw ? new Date(a.shippingDateRaw) : null;
    const isOverdueA = !isShippedA && !isCancelledA && shippingDateA && shippingDateA < today;

    const isShippedB = b.shipped === '○' || b.shipped === true;
    const isCancelledB = b.status === 'キャンセル' || b.status === 'cancelled';
    const shippingDateB = b.shippingDateRaw ? new Date(b.shippingDateRaw) : null;
    const isOverdueB = !isShippedB && !isCancelledB && shippingDateB && shippingDateB < today;

    // 1. 予定日超過を優先（超過が先頭）
    if (isOverdueA && !isOverdueB) return -1;
    if (!isOverdueA && isOverdueB) return 1;

    // 2. 発送日昇順
    const dateA = a.shippingDateRaw ? new Date(a.shippingDateRaw) : new Date(0);
    const dateB = b.shippingDateRaw ? new Date(b.shippingDateRaw) : new Date(0);
    if (dateA.getTime() !== dateB.getTime()) {
      return dateA - dateB;
    }

    // 3. 顧客名順
    return (a.customerName || '').localeCompare(b.customerName || '');
  });

  // 日付をフォーマット & shippingDateRawを削除 & 商品情報を整形
  orders = orders.map(order => {
    return {
      orderId: order.orderId,
      orderDate: formatDateForDisplay(order.orderDate),
      shippingDate: formatDateForDisplay(order.shippingDate),
      deliveryDate: formatDateForDisplay(order.deliveryDate),
      customerName: order.customerName,
      destinationName: order.destinationName,
      shipped: order.shipped,  // 出荷済フラグ
      trackingNumber: order.trackingNumber, // 追跡番号
      status: order.status,    // ステータス
      items: order.items,  // 商品配列をそのまま渡す
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

/**
 * 受注の出荷済ステータスを更新
 * @param {string} orderId - 受注ID
 * @param {string} shippedValue - 出荷済の値（'○' or ''）。'○' の場合は納品方法に応じてステータス列も更新する（配達/店舗受取 → 配達完了、それ以外 → 出荷済）
 * @returns {Object} - 更新結果
 */
function updateOrderShippedStatus(orderId, shippedValue) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('受注');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  // 列インデックスを取得
  const orderIdCol = headers.indexOf('受注ID');
  const shippedCol = headers.indexOf('出荷済');
  const statusCol = headers.indexOf('ステータス');
  const deliveryMethodCol = headers.indexOf('納品方法');

  if (orderIdCol === -1 || shippedCol === -1) {
    throw new Error('必要な列が見つかりません');
  }

  // 該当する受注IDの全行を更新
  let updatedCount = 0;
  for (let i = 1; i < data.length; i++) {
    if (data[i][orderIdCol] === orderId) {
      sheet.getRange(i + 1, shippedCol + 1).setValue(shippedValue);

      // 出荷済にする場合、納品方法に応じてステータスを設定
      if (shippedValue === '○' && statusCol !== -1) {
        const deliveryMethod = deliveryMethodCol !== -1 ? (data[i][deliveryMethodCol] || '') : '';
        const newStatus = (deliveryMethod === '配達' || deliveryMethod === '店舗受取')
          ? '配達完了'
          : '出荷済';
        sheet.getRange(i + 1, statusCol + 1).setValue(newStatus);
      }

      updatedCount++;
    }
  }

  if (updatedCount === 0) {
    throw new Error('該当する受注IDが見つかりません: ' + orderId);
  }

  return { success: true, updatedRows: updatedCount };
}

/**
 * 複数受注の出荷済ステータスを一括更新
 * @param {Array<string>} orderIds - 受注IDの配列
 * @param {string} shippedValue - 出荷済の値（'○' or ''）。'○' の場合は納品方法に応じてステータス列も更新する（配達/店舗受取 → 配達完了、それ以外 → 出荷済）
 * @returns {Object} - 更新結果
 */
function bulkUpdateOrderShippedStatus(orderIds, shippedValue) {
  if (!orderIds || orderIds.length === 0) {
    throw new Error('受注IDが指定されていません');
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('受注');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const orderIdCol = headers.indexOf('受注ID');
  const shippedCol = headers.indexOf('出荷済');
  const statusCol = headers.indexOf('ステータス');
  const deliveryMethodCol = headers.indexOf('納品方法');

  if (orderIdCol === -1 || shippedCol === -1) {
    throw new Error('必要な列が見つかりません');
  }

  const orderIdSet = new Set(orderIds);
  let updatedCount = 0;

  for (let i = 1; i < data.length; i++) {
    if (orderIdSet.has(data[i][orderIdCol])) {
      sheet.getRange(i + 1, shippedCol + 1).setValue(shippedValue);

      // 出荷済にする場合、納品方法に応じてステータスを設定
      if (shippedValue === '○' && statusCol !== -1) {
        const deliveryMethod = deliveryMethodCol !== -1 ? (data[i][deliveryMethodCol] || '') : '';
        const newStatus = (deliveryMethod === '配達' || deliveryMethod === '店舗受取')
          ? '配達完了'
          : '出荷済';
        sheet.getRange(i + 1, statusCol + 1).setValue(newStatus);
      }

      updatedCount++;
    }
  }

  return { success: true, updatedRows: updatedCount, orderCount: orderIds.length };
}

/**
 * 複数受注を一括キャンセル（ステータスを「キャンセル」に更新）
 * @param {Array<string>} orderIds - 受注IDの配列
 * @returns {Object} - 更新結果
 */
function bulkCancelOrders(orderIds) {
  if (!orderIds || orderIds.length === 0) {
    throw new Error('受注IDが指定されていません');
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('受注');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const orderIdCol = headers.indexOf('受注ID');
  let statusCol = headers.indexOf('ステータス');

  if (orderIdCol === -1) {
    throw new Error('受注ID列が見つかりません');
  }

  // ステータス列がなければ追加
  if (statusCol === -1) {
    statusCol = headers.length;
    sheet.getRange(1, statusCol + 1).setValue('ステータス');
  }

  const orderIdSet = new Set(orderIds);
  let updatedCount = 0;

  for (let i = 1; i < data.length; i++) {
    if (orderIdSet.has(data[i][orderIdCol])) {
      sheet.getRange(i + 1, statusCol + 1).setValue('キャンセル');
      updatedCount++;
    }
  }

  return { success: true, updatedRows: updatedCount, orderCount: orderIds.length };
}

/**
 * 複数受注のキャンセルを解除（ステータスを空に更新）
 * @param {Array<string>} orderIds - 受注IDの配列
 * @returns {Object} - 更新結果
 */
function bulkUncancelOrders(orderIds) {
  if (!orderIds || orderIds.length === 0) {
    throw new Error('受注IDが指定されていません');
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('受注');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const orderIdCol = headers.indexOf('受注ID');
  const statusCol = headers.indexOf('ステータス');

  if (orderIdCol === -1) {
    throw new Error('受注ID列が見つかりません');
  }

  if (statusCol === -1) {
    // ステータス列がなければ何もしない
    return { success: true, updatedRows: 0, orderCount: orderIds.length, message: 'ステータス列がありません' };
  }

  const orderIdSet = new Set(orderIds);
  let updatedCount = 0;

  for (let i = 1; i < data.length; i++) {
    if (orderIdSet.has(data[i][orderIdCol])) {
      sheet.getRange(i + 1, statusCol + 1).setValue('');  // ステータスを空にする
      updatedCount++;
    }
  }

  return { success: true, updatedRows: updatedCount, orderCount: orderIds.length };
}

/**
 * 複数受注のステータスを一括更新（収穫待ち等に使用）
 * @param {Array<string>} orderIds - 受注IDの配列
 * @param {string} newStatus - 新しいステータス（'発送前'/'収穫待ち'/'出荷済み'等）
 * @returns {Object} - 更新結果
 */
function bulkUpdateOrderStatus(orderIds, newStatus) {
  if (!orderIds || orderIds.length === 0) {
    throw new Error('受注IDが指定されていません');
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('受注');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const orderIdCol = headers.indexOf('受注ID');
  let statusCol = headers.indexOf('ステータス');
  const shippedCol = headers.indexOf('出荷済');

  if (orderIdCol === -1) {
    throw new Error('受注ID列が見つかりません');
  }

  // ステータス列がなければ追加
  if (statusCol === -1) {
    statusCol = headers.length;
    sheet.getRange(1, statusCol + 1).setValue('ステータス');
  }

  const orderIdSet = new Set(orderIds);
  let updatedCount = 0;

  for (let i = 1; i < data.length; i++) {
    if (orderIdSet.has(data[i][orderIdCol])) {
      sheet.getRange(i + 1, statusCol + 1).setValue(newStatus);
      // 出荷済みの場合は出荷済フラグも更新
      if (newStatus === '出荷済み' && shippedCol !== -1) {
        sheet.getRange(i + 1, shippedCol + 1).setValue('○');
      } else if (newStatus !== '出荷済み' && shippedCol !== -1) {
        sheet.getRange(i + 1, shippedCol + 1).setValue('');
      }
      updatedCount++;
    }
  }

  return { success: true, updatedRows: updatedCount, orderCount: orderIds.length };
}

/**
 * 複数受注を一括削除（ヤマトCSV/佐川CSVも連動削除）
 * @param {Array<string>} orderIds - 受注IDの配列
 * @returns {Object} - 削除結果
 */
function bulkDeleteOrders(orderIds) {
  if (!orderIds || orderIds.length === 0) {
    throw new Error('受注IDが指定されていません');
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const orderSheet = ss.getSheetByName('受注');
  const yamatoSheet = ss.getSheetByName('ヤマトCSV');
  const sagawaSheet = ss.getSheetByName('佐川CSV');

  if (!orderSheet) {
    throw new Error('受注シートが見つかりません');
  }

  const orderData = orderSheet.getDataRange().getValues();
  const orderHeaders = orderData[0];
  const orderIdCol = orderHeaders.indexOf('受注ID');
  const deliveryMethodCol = orderHeaders.indexOf('納品方法');

  if (orderIdCol === -1) {
    throw new Error('受注ID列が見つかりません');
  }

  const orderIdSet = new Set(orderIds);

  // 各受注の納品方法を取得して、対応するCSVも削除
  const yamatoOrderIds = new Set();
  const sagawaOrderIds = new Set();

  for (let i = 1; i < orderData.length; i++) {
    const orderId = orderData[i][orderIdCol];
    if (orderIdSet.has(orderId)) {
      const deliveryMethod = orderData[i][deliveryMethodCol] || '';
      if (deliveryMethod.includes('ヤマト')) {
        yamatoOrderIds.add(orderId);
      } else if (deliveryMethod.includes('佐川')) {
        sagawaOrderIds.add(orderId);
      }
    }
  }

  // 受注シートから削除（逆順で削除）
  const orderRowsToDelete = [];
  for (let i = 1; i < orderData.length; i++) {
    if (orderIdSet.has(orderData[i][orderIdCol])) {
      orderRowsToDelete.push(i + 1);  // 1-indexed
    }
  }
  for (let i = orderRowsToDelete.length - 1; i >= 0; i--) {
    orderSheet.deleteRow(orderRowsToDelete[i]);
  }

  // ヤマトCSVから削除
  if (yamatoSheet && yamatoOrderIds.size > 0) {
    deleteCSVByOrderIds(yamatoSheet, yamatoOrderIds);
  }

  // 佐川CSVから削除
  if (sagawaSheet && sagawaOrderIds.size > 0) {
    deleteCSVByOrderIds(sagawaSheet, sagawaOrderIds);
  }

  return {
    success: true,
    deletedOrderRows: orderRowsToDelete.length,
    orderCount: orderIds.length
  };
}

/**
 * CSVシートから指定された受注IDの行を削除
 * @param {Sheet} sheet - CSVシート
 * @param {Set<string>} orderIds - 削除対象の受注IDセット
 */
function deleteCSVByOrderIds(sheet, orderIds) {
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  // お客様管理番号列を探す（受注IDが格納されている）
  const customerNoCol = headers.indexOf('お客様管理番号');
  if (customerNoCol === -1) return;

  const rowsToDelete = [];
  for (let i = 1; i < data.length; i++) {
    if (orderIds.has(data[i][customerNoCol])) {
      rowsToDelete.push(i + 1);  // 1-indexed
    }
  }

  // 逆順で削除
  for (let i = rowsToDelete.length - 1; i >= 0; i--) {
    sheet.deleteRow(rowsToDelete[i]);
  }
}
