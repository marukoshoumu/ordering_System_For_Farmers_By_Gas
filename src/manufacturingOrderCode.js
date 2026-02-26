/**
 * 製造指示書データを取得
 * @param {string} paramsJson - JSON文字列 { dateFrom, dateTo, category }
 * @returns {string} JSON文字列
 */
function getManufacturingOrderData(paramsJson) {
  const params = JSON.parse(paramsJson);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('受注');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const orderIdCol = headers.indexOf('受注ID');
  const shippingDateCol = headers.indexOf('発送日');
  const customerNameCol = headers.indexOf('顧客名');
  const productCol = headers.indexOf('商品名');
  const categoryCol = headers.indexOf('商品分類');
  const quantityCol = headers.indexOf('受注数');
  const memoCol = headers.indexOf('メモ');
  const statusCol = headers.indexOf('ステータス');

  const dateFrom = params.dateFrom ? new Date(params.dateFrom) : null;
  const dateTo = params.dateTo ? new Date(params.dateTo) : null;
  const filterCategory = params.category || 'all';

  if (dateFrom) dateFrom.setHours(0, 0, 0, 0);
  if (dateTo) dateTo.setHours(23, 59, 59, 999);

  const excludeCategories = ['送料', '手数料', '追加料金', '送料・代引手数料'];
  const excludeProducts = ['送料', 'クール便追加', '代引手数料', '梱包料'];

  const orderMap = new Map();

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const orderId = row[orderIdCol];
    if (!orderId) continue;

    const status = row[statusCol] || '';
    if (status === 'キャンセル' || status === 'cancelled') continue;

    const shippingDateRaw = row[shippingDateCol];
    if (shippingDateRaw && dateFrom && dateTo) {
      const shipDate = new Date(shippingDateRaw);
      shipDate.setHours(0, 0, 0, 0);
      if (shipDate < dateFrom || shipDate > dateTo) continue;
    }

    const category = row[categoryCol] || '';
    if (excludeCategories.includes(category)) continue;
    const product = row[productCol] || '';
    if (excludeProducts.includes(product)) continue;

    if (filterCategory !== 'all' && category !== filterCategory) continue;

    const customerName = row[customerNameCol] || '';
    const quantity = Number(row[quantityCol]) || 0;
    const memo = memoCol !== -1 ? (row[memoCol] || '') : '';

    if (!orderMap.has(orderId)) {
      orderMap.set(orderId, {
        orderId: orderId,
        customerName: customerName,
        items: []
      });
    }

    orderMap.get(orderId).items.push({
      product: product,
      category: category,
      quantity: quantity,
      memo: memo
    });
  }

  // カテゴリ別にグループ化
  const categoryMap = new Map();
  for (const order of orderMap.values()) {
    for (const item of order.items) {
      if (!categoryMap.has(item.category)) {
        categoryMap.set(item.category, []);
      }
      categoryMap.get(item.category).push({
        customerName: order.customerName,
        product: item.product,
        quantity: item.quantity,
        memo: item.memo
      });
    }
  }

  const result = [];
  for (const [category, rows] of categoryMap.entries()) {
    result.push({ category: category, rows: rows });
  }

  return JSON.stringify({ groups: result });
}
