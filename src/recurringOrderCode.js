/**
 * 定期受注（定期便）管理モジュール
 *
 * 定期便として登録された受注を管理し、指定間隔で自動的に
 * 新しい受注を作成する機能を提供します。
 *
 * 主な機能:
 * - 定期受注の登録
 * - 定期受注一覧の取得
 * - 定期受注の更新（間隔変更、解除、削除）
 * - トリガーによる自動受注作成
 */

/**
 * 月を加算する（月末調整あり）
 * 例: 1/30 + 1ヶ月 → 2/28（2月末）
 * @param {Date} date - 基準日
 * @param {number} months - 加算する月数
 * @returns {Date} 加算後の日付
 */
function addMonthsWithAdjustment(date, months) {
  const result = new Date(date);
  const originalDay = result.getDate();
  result.setMonth(result.getMonth() + months);

  // 日がオーバーフローして翌月になった場合、前月の末日に調整
  if (result.getDate() !== originalDay) {
    result.setDate(0); // 前月の末日
  }
  return result;
}

/**
 * 基準日を「カレンダー日のみ」に正規化する（曜日計算のタイムゾーンずれ防止）
 * "YYYY-MM-DD" が UTC で解釈され getDay() が意図とずれる場合に、表示上の日付で再構築する。
 * @param {Date} date - 基準日
 * @returns {Date} 同一カレンダー日のローカル 0:00 の Date
 */
function normalizeToDateOnly(date) {
  const y = date.getFullYear();
  const m = date.getMonth();
  const day = date.getDate();
  return new Date(y, m, day);
}

/**
 * 柔軟な間隔形式で次回発送日を計算する
 * @param {Date} baseDate - 基準日
 * @param {Object} intervalObj - 間隔オブジェクト {type, value, weekday?}
 * @returns {Date} 次回発送日
 */
function calcNextShippingDateFlexible(baseDate, intervalObj) {
  const d = normalizeToDateOnly(new Date(baseDate));
  if (!intervalObj || !intervalObj.type) {
    // 旧形式（数値のみ）の場合
    const months = Number(intervalObj) || 1;
    return addMonthsWithAdjustment(d, months);
  }

  switch (intervalObj.type) {
    case 'weekly':
      // 次の指定曜日（value = 1-7, 1=月曜）
      const targetDay = Number(intervalObj.value) || 1;
      const currentDay = d.getDay() || 7; // 日曜=0 → 7
      let diff = (targetDay - currentDay + 7) % 7;
      if (diff === 0) diff = 7; // 同じ曜日なら来週
      d.setDate(d.getDate() + diff);
      return d;

    case 'nweek':
    case 'biweekly':
    case 'triweekly':
      // n週ごと（value = 週数, weekday = 曜日番号1-7）
      // 指定曜日をベースにし、その日付から n 週間後を次回発送日とする。
      // 基準日が指定曜日でない場合は、基準日以降の「直後のその曜日」をアンカーにしてから n 週間後を返す。
      const n = Number(intervalObj.value) || 2;
      const weekday = Number(intervalObj.weekday) || 1;
      const currentDay2 = d.getDay() || 7; // 日曜=0→7 でフォームの 1-7 と揃える
      let daysToAnchor = (weekday - currentDay2 + 7) % 7;
      if (daysToAnchor === 0) {
        // 基準日がすでに指定曜日 → その日をアンカーとして n 週間後
        d.setDate(d.getDate() + 7 * n);
      } else {
        // 基準日以降の直後の指定曜日をアンカーにして、その日から n 週間後
        d.setDate(d.getDate() + daysToAnchor + 7 * n);
      }
      return d;

    case 'monthly':
      // 指定日（value = 1-31, 'first', 'last'）
      if (intervalObj.value === 'first') {
        d.setMonth(d.getMonth() + 1, 1);
      } else if (intervalObj.value === 'last') {
        d.setMonth(d.getMonth() + 2, 0); // 翌月の0日 = 今月末
      } else {
        const targetDayOfMonth = Number(intervalObj.value) || 1;
        d.setMonth(d.getMonth() + 1);
        // 月末調整
        const lastDay = new Date(d.getFullYear(), d.getMonth() + 1, 0).getDate();
        d.setDate(Math.min(targetDayOfMonth, lastDay));
      }
      return d;

    case 'nmonth':
    case '2month':
    case '3month':
      // nヶ月ごと
      const months2 = Number(intervalObj.value) || 1;
      return addMonthsWithAdjustment(d, months2);

    default:
      // 旧形式互換: 数値として扱う
      const months3 = Number(intervalObj.value) || Number(intervalObj) || 1;
      return addMonthsWithAdjustment(d, months3);
  }
}

/**
 * 定期受注シートのヘッダー定義
 * @returns {Array<string>} ヘッダー配列
 */
function getRecurringOrderHeaders() {
  return [
    '定期受注ID',       // 一意のID
    '登録日',           // 定期便として登録した日
    '更新間隔',         // JSON: {type, value, ...} 例: {type:'monthly',value:15} 旧: 1,2,3
    '次回発送日',       // 次に受注を作成する際の発送日
    '次回納品日',       // 次に受注を作成する際の納品日
    'ステータス',       // 有効/停止/解除
    '最終実行日',       // 最後に自動受注を作成した日
    '顧客名',
    '顧客郵便番号',
    '顧客住所',
    '顧客電話番号',
    '発送先名',
    '発送先郵便番号',
    '発送先住所',
    '発送先電話番号',
    '発送元名',
    '発送元郵便番号',
    '発送元住所',
    '発送元電話番号',
    '受付方法',
    '受付者',
    '納品方法',
    '配達時間帯',
    '納品書',
    '請求書',
    '領収書',
    'パンフ',
    'レシピ',
    'その他添付',
    '品名',
    '送り状種別',
    'クール区分',
    '荷扱い１',
    '荷扱い２',
    '荷扱い３',
    '代引総額',
    '代引内税',
    '発行枚数',
    '社内メモ',
    '送り状備考欄',
    '納品書備考欄',
    'メモ',
    '納品書テキスト',
    '商品データJSON'    // 商品情報をJSON形式で保存
  ];
}

/**
 * 定期受注シートを取得（なければ作成）
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} 定期受注シート
 */
function getRecurringOrderSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('定期受注');

  if (!sheet) {
    sheet = ss.insertSheet('定期受注');
    const headers = getRecurringOrderHeaders();
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
    Logger.log('定期受注シートを作成しました');
  }

  return sheet;
}

/**
 * 定期受注を登録する
 * @param {Object} e - フォームパラメータ
 * @returns {string} 定期受注ID
 */
function createRecurringOrder(e) {
  const sheet = getRecurringOrderSheet();
  const headers = getRecurringOrderHeaders();
  const recurringId = generateId();
  const dateNow = Utilities.formatDate(new Date(), 'JST', 'yyyy/MM/dd');

  // 商品データをJSON化
  const products = [];
  for (let i = 1; i <= 10; i++) {
    const bunrui = e.parameter['bunrui' + i];
    const product = e.parameter['product' + i];
    const price = e.parameter['price' + i];
    const quantity = e.parameter['quantity' + i];

    if (quantity && Number(quantity) > 0) {
      products.push({
        bunrui: bunrui || '',
        product: product || '',
        price: Number(price) || 0,
        quantity: Number(quantity)
      });
    }
  }

  // 柔軟な間隔データを受け取る
  let intervalObj;
  try {
    intervalObj = JSON.parse(e.parameter.recurringInterval);
  } catch {
    // 旧形式（数値のみ）
    intervalObj = { type: 'nmonth', value: Number(e.parameter.recurringInterval) || 1 };
  }
  const baseShippingDate = new Date(e.parameter.shippingDate1);
  const nextShippingDate = calcNextShippingDateFlexible(baseShippingDate, intervalObj);
  // 次回納品日を計算（発送日との差分を維持）
  const baseDeliveryDate = new Date(e.parameter.deliveryDate1);
  const daysDiff = Math.round((baseDeliveryDate - baseShippingDate) / (1000 * 60 * 60 * 24));
  const nextDeliveryDate = new Date(nextShippingDate);
  nextDeliveryDate.setDate(nextDeliveryDate.getDate() + daysDiff);

  const record = {
    '定期受注ID': recurringId,
    '登録日': dateNow,
    '更新間隔': JSON.stringify(intervalObj),
    '次回発送日': Utilities.formatDate(nextShippingDate, 'JST', 'yyyy/MM/dd'),
    '次回納品日': Utilities.formatDate(nextDeliveryDate, 'JST', 'yyyy/MM/dd'),
    'ステータス': '有効',
    '最終実行日': '',
    '顧客名': e.parameter.customerName || '',
    '顧客郵便番号': e.parameter.customerZipcode || '',
    '顧客住所': e.parameter.customerAddress || '',
    '顧客電話番号': e.parameter.customerTel || '',
    '発送先名': e.parameter.shippingToName || '',
    '発送先郵便番号': e.parameter.shippingToZipcode || '',
    '発送先住所': e.parameter.shippingToAddress || '',
    '発送先電話番号': e.parameter.shippingToTel || '',
    '発送元名': e.parameter.shippingFromName || '',
    '発送元郵便番号': e.parameter.shippingFromZipcode || '',
    '発送元住所': e.parameter.shippingFromAddress || '',
    '発送元電話番号': e.parameter.shippingFromTel || '',
    '受付方法': e.parameter.receiptWay || '',
    '受付者': e.parameter.recipient || '',
    '納品方法': e.parameter.deliveryMethod || '',
    '配達時間帯': e.parameter.deliveryTime ? e.parameter.deliveryTime.split(':')[1] : '',
    '納品書': e.parameters.checklist && e.parameters.checklist.includes('納品書') ? '○' : '',
    '請求書': e.parameters.checklist && e.parameters.checklist.includes('請求書') ? '○' : '',
    '領収書': e.parameters.checklist && e.parameters.checklist.includes('領収書') ? '○' : '',
    'パンフ': e.parameters.checklist && e.parameters.checklist.includes('パンフ') ? '○' : '',
    'レシピ': e.parameters.checklist && e.parameters.checklist.includes('レシピ') ? '○' : '',
    'その他添付': e.parameter.otherAttach || '',
    '品名': e.parameter.sendProduct || '',
    '送り状種別': e.parameter.invoiceType ? e.parameter.invoiceType.split(':')[1] : '',
    'クール区分': e.parameter.coolCls ? e.parameter.coolCls.split(':')[1] : '',
    '荷扱い１': e.parameter.cargo1 ? e.parameter.cargo1.split(':')[1] : '',
    '荷扱い２': e.parameter.cargo2 ? e.parameter.cargo2.split(':')[1] : '',
    '荷扱い３': e.parameter.cargo3 ? e.parameter.cargo3.split(':')[1] : '',
    '代引総額': e.parameter.cashOnDelivery || '',
    '代引内税': e.parameter.cashOnDeliTax || '',
    '発行枚数': e.parameter.copiePrint || '',
    '社内メモ': e.parameter.internalMemo || '',
    '送り状備考欄': e.parameter.csvmemo || '',
    '納品書備考欄': e.parameter.deliveryMemo || '',
    'メモ': e.parameter.memo || '',
    '納品書テキスト': e.parameter.deliveryNoteText1 || '',
    '商品データJSON': JSON.stringify(products)
  };

  // ヘッダー順に配列化
  const rowData = headers.map(function(header) {
    return record[header] || '';
  });

  // テキスト形式で書き込み（郵便番号・電話番号の先頭0を保持）
  const lastRow = sheet.getLastRow();
  sheet.getRange(lastRow + 1, 1, 1, rowData.length).setNumberFormat('@').setValues([rowData]);
  Logger.log('定期受注を登録しました: ' + recurringId);

  return recurringId;
}

/**
 * 定期受注一覧を取得する
 * @param {Object} filters - フィルター条件（オプション）
 * @returns {Array<Object>} 定期受注データの配列
 */
function getRecurringOrders(filters) {
  const sheet = getRecurringOrderSheet();
  const data = sheet.getDataRange().getValues();

  if (data.length <= 1) {
    return [];
  }

  const headers = data[0];
  const orders = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const order = {};

    for (let j = 0; j < headers.length; j++) {
      if (headers[j] === '更新間隔') {
        try {
          order[headers[j]] = JSON.parse(row[j]);
        } catch {
          // 旧データは数値
          order[headers[j]] = { type: 'nmonth', value: Number(row[j]) || 1 };
        }
      } else {
        order[headers[j]] = row[j];
      }
    }

    // フィルター適用
    if (filters) {
      if (filters.status && order['ステータス'] !== filters.status) {
        continue;
      }
      if (filters.customerName && order['顧客名'].indexOf(filters.customerName) === -1) {
        continue;
      }
    }

    // 商品データをパース
    try {
      order['商品データ'] = JSON.parse(order['商品データJSON'] || '[]');
    } catch (e) {
      order['商品データ'] = [];
    }

    orders.push(order);
  }

  return orders;
}

/**
 * 定期受注を取得する（ID指定）
 * @param {string} recurringId - 定期受注ID
 * @returns {Object|null} 定期受注データ
 */
function getRecurringOrderById(recurringId) {
  const sheet = getRecurringOrderSheet();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idIndex = headers.indexOf('定期受注ID');

  for (let i = 1; i < data.length; i++) {
    if (data[i][idIndex] === recurringId) {
      const order = {};
      for (let j = 0; j < headers.length; j++) {
        if (headers[j] === '更新間隔') {
          try {
            order[headers[j]] = JSON.parse(data[i][j]);
          } catch {
            order[headers[j]] = { type: 'nmonth', value: Number(data[i][j]) || 1 };
          }
        } else {
          order[headers[j]] = data[i][j];
        }
      }
      try {
        order['商品データ'] = JSON.parse(order['商品データJSON'] || '[]');
      } catch (e) {
        order['商品データ'] = [];
      }
      order['rowIndex'] = i + 1; // 1-indexed
      return order;
    }
  }

  return null;
}

/**
 * 定期受注を更新する
 * @param {string} recurringId - 定期受注ID
 * @param {Object} updates - 更新する項目
 * @returns {boolean} 成功/失敗
 */
function updateRecurringOrder(recurringId, updates) {
  const sheet = getRecurringOrderSheet();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idIndex = headers.indexOf('定期受注ID');
  // テキスト形式が必要なカラム（郵便番号・電話番号）
  const textFormatColumns = ['顧客郵便番号', '顧客電話番号', '発送先郵便番号', '発送先電話番号', '発送元郵便番号', '発送元電話番号'];

  for (let i = 1; i < data.length; i++) {
    if (data[i][idIndex] === recurringId) {
      // 更新対象のカラムを特定して更新
      for (const key in updates) {
        const colIndex = headers.indexOf(key);
        if (colIndex !== -1) {
          const cell = sheet.getRange(i + 1, colIndex + 1);
          // 郵便番号・電話番号はテキスト形式で書き込み
          if (textFormatColumns.indexOf(key) !== -1) {
            cell.setNumberFormat('@').setValue(updates[key]);
          } else {
            cell.setValue(updates[key]);
          }
        }
      }
      Logger.log('定期受注を更新しました: ' + recurringId);
      return true;
    }
  }

  Logger.log('定期受注が見つかりません: ' + recurringId);
  return false;
}

/**
 * 定期受注を削除する
 * @param {string} recurringId - 定期受注ID
 * @returns {boolean} 成功/失敗
 */
function deleteRecurringOrder(recurringId) {
  const sheet = getRecurringOrderSheet();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idIndex = headers.indexOf('定期受注ID');

  for (let i = 1; i < data.length; i++) {
    if (data[i][idIndex] === recurringId) {
      sheet.deleteRow(i + 1);
      Logger.log('定期受注を削除しました: ' + recurringId);
      return true;
    }
  }

  Logger.log('定期受注が見つかりません: ' + recurringId);
  return false;
}

/**
 * 定期受注の間隔を変更する
 * @param {string} recurringId - 定期受注ID
 * @param {number} newInterval - 新しい間隔（1, 2, 3）
 * @param {string} nextShippingDate - 次回発送日（yyyy-MM-dd形式）
 * @param {string} nextDeliveryDate - 次回納品日（yyyy-MM-dd形式）
 * @returns {boolean} 成功/失敗
 */
function changeRecurringInterval(recurringId, newInterval, nextShippingDate, nextDeliveryDate) {
  const order = getRecurringOrderById(recurringId);
  if (!order) {
    return false;
  }

  // 日付をyyyy/MM/dd形式に変換
  const formattedShippingDate = nextShippingDate.replace(/-/g, '/');
  const formattedDeliveryDate = nextDeliveryDate.replace(/-/g, '/');

  let intervalObj;
  try {
    intervalObj = (typeof newInterval === 'string') ? JSON.parse(newInterval) : newInterval;
  } catch {
    intervalObj = { type: 'nmonth', value: Number(newInterval) || 1 };
  }
  return updateRecurringOrder(recurringId, {
    '更新間隔': JSON.stringify(intervalObj),
    '次回発送日': formattedShippingDate,
    '次回納品日': formattedDeliveryDate
  });
}

/**
 * 定期受注を一時停止する
 * @param {string} recurringId - 定期受注ID
 * @returns {boolean} 成功/失敗
 */
function pauseRecurringOrder(recurringId) {
  return updateRecurringOrder(recurringId, {
    'ステータス': '停止'
  });
}

/**
 * 定期受注を再開する
 * 現在の日付から次回発送日を再計算
 * @param {string} recurringId - 定期受注ID
 * @returns {boolean} 成功/失敗
 */
function resumeRecurringOrder(recurringId) {
  const order = getRecurringOrderById(recurringId);
  if (!order) {
    return false;
  }

  // 発送日と納品日の差分を計算
  const currentNextShippingDate = new Date(order['次回発送日']);
  const currentNextDeliveryDate = new Date(order['次回納品日']);
  const daysDiff = Math.round((currentNextDeliveryDate - currentNextShippingDate) / (1000 * 60 * 60 * 24));

  // 現在の日付から次回発送日を計算（今日 + 間隔月数、月末調整あり）
  const today = new Date();
  const intervalObj = order['更新間隔'];
  const newNextShippingDate = calcNextShippingDateFlexible(today, intervalObj);

  // 新しい次回納品日を計算（発送日との差分を維持）
  const newNextDeliveryDate = new Date(newNextShippingDate);
  newNextDeliveryDate.setDate(newNextDeliveryDate.getDate() + daysDiff);

  return updateRecurringOrder(recurringId, {
    'ステータス': '有効',
    '次回発送日': Utilities.formatDate(newNextShippingDate, 'JST', 'yyyy/MM/dd'),
    '次回納品日': Utilities.formatDate(newNextDeliveryDate, 'JST', 'yyyy/MM/dd')
  });
}

/**
 * 定期受注を解除する（論理削除）
 * @param {string} recurringId - 定期受注ID
 * @returns {boolean} 成功/失敗
 */
function cancelRecurringOrder(recurringId) {
  return updateRecurringOrder(recurringId, {
    'ステータス': '解除'
  });
}

/**
 * 【トリガー用】定期受注を処理して受注を自動作成する
 * 毎日実行されることを想定
 */
function processRecurringOrders() {
  const today = new Date();
  const todayStr = Utilities.formatDate(today, 'JST', 'yyyy/MM/dd');
  const orders = getRecurringOrders({ status: '有効' });

  Logger.log('定期受注処理開始: ' + todayStr + ', 対象件数: ' + orders.length);

  let processedCount = 0;

  for (let i = 0; i < orders.length; i++) {
    const order = orders[i];
    // 日付を0:00に正規化
    const nextShippingDate = new Date(order['次回発送日']);
    nextShippingDate.setHours(0, 0, 0, 0);
    const todayDate = new Date(today);
    todayDate.setHours(0, 0, 0, 0);

    // 7日前に受注を作成（トリガー時刻ズレ対策で6〜7日差を許容）
    const diffMs = nextShippingDate - todayDate;
    const diffDays = diffMs / (1000 * 60 * 60 * 24);
    if (diffDays >= 6 && diffDays <= 7) {
      try {
        createOrderFromRecurring(order);
        processedCount++;

        // 次回日程を更新
        const intervalObj = order['更新間隔'];
        const newNextShippingDate = calcNextShippingDateFlexible(nextShippingDate, intervalObj);
        const daysDiff = Math.round((new Date(order['次回納品日']) - nextShippingDate) / (1000 * 60 * 60 * 24));
        const newNextDeliveryDate = new Date(newNextShippingDate);
        newNextDeliveryDate.setDate(newNextDeliveryDate.getDate() + daysDiff);
        updateRecurringOrder(order['定期受注ID'], {
          '次回発送日': Utilities.formatDate(newNextShippingDate, 'JST', 'yyyy/MM/dd'),
          '次回納品日': Utilities.formatDate(newNextDeliveryDate, 'JST', 'yyyy/MM/dd'),
          '最終実行日': todayStr
        });

        Logger.log('定期受注から受注を作成: ' + order['定期受注ID'] + ' -> ' + order['顧客名']);

      } catch (error) {
        Logger.log('定期受注処理エラー: ' + order['定期受注ID'] + ' - ' + error.message);
      }
    }
  }

  Logger.log('定期受注処理完了: ' + processedCount + '件作成');
}

/**
 * 定期受注データから受注を作成する
 * @param {Object} recurringOrder - 定期受注データ
 */
function createOrderFromRecurring(recurringOrder) {
  const orderHeaders = getOrderSheetHeaders();
  const deliveryId = generateId();
  const dateNow = Utilities.formatDate(new Date(), 'JST', 'yyyy/MM/dd');

  const products = recurringOrder['商品データ'] || [];
  if (products.length === 0) {
    Logger.log('商品データがありません: ' + recurringOrder['定期受注ID']);
    return;
  }

  const records = [];
  const createRecords = [];

  for (let i = 0; i < products.length; i++) {
    const p = products[i];
    const record = {};

    record['受注ID'] = deliveryId;
    record['受注日'] = dateNow;
    record['顧客名'] = recurringOrder['顧客名'];
    record['顧客郵便番号'] = recurringOrder['顧客郵便番号'];
    record['顧客住所'] = recurringOrder['顧客住所'];
    record['顧客電話番号'] = recurringOrder['顧客電話番号'];
    record['発送先名'] = recurringOrder['発送先名'];
    record['発送先郵便番号'] = recurringOrder['発送先郵便番号'];
    record['発送先住所'] = recurringOrder['発送先住所'];
    record['発送先電話番号'] = recurringOrder['発送先電話番号'];
    record['発送元名'] = recurringOrder['発送元名'];
    record['発送元郵便番号'] = recurringOrder['発送元郵便番号'];
    record['発送元住所'] = recurringOrder['発送元住所'];
    record['発送元電話番号'] = recurringOrder['発送元電話番号'];
    record['発送日'] = recurringOrder['次回発送日'];
    record['納品日'] = recurringOrder['次回納品日'];
    record['受付方法'] = recurringOrder['受付方法'];
    record['受付者'] = recurringOrder['受付者'];
    record['納品方法'] = recurringOrder['納品方法'];
    record['配達時間帯'] = recurringOrder['配達時間帯'];
    record['納品書'] = recurringOrder['納品書'];
    record['請求書'] = recurringOrder['請求書'];
    record['領収書'] = recurringOrder['領収書'];
    record['パンフ'] = recurringOrder['パンフ'];
    record['レシピ'] = recurringOrder['レシピ'];
    record['その他添付'] = recurringOrder['その他添付'];
    record['品名'] = recurringOrder['品名'];
    record['送り状種別'] = recurringOrder['送り状種別'];
    record['クール区分'] = recurringOrder['クール区分'];
    record['荷扱い１'] = recurringOrder['荷扱い１'];
    record['荷扱い２'] = recurringOrder['荷扱い２'];
    record['荷扱い３'] = recurringOrder['荷扱い３'];
    record['代引総額'] = recurringOrder['代引総額'];
    record['代引内税'] = recurringOrder['代引内税'];
    record['発行枚数'] = recurringOrder['発行枚数'];
    var baseMemo = recurringOrder['社内メモ'] || '';
    record['社内メモ'] = baseMemo + '（定期便自動作成）';
    record['送り状備考欄'] = recurringOrder['送り状備考欄'];
    record['納品書備考欄'] = recurringOrder['納品書備考欄'];
    record['メモ'] = recurringOrder['メモ'];
    record['出荷済'] = '';
    record['ステータス'] = '';
    record['追跡番号'] = '';
    record['紐付け受注ID'] = '';
    record['納品書テキスト'] = recurringOrder['納品書テキスト'];
    record['商品分類'] = p.bunrui;
    record['商品名'] = p.product;
    record['受注数'] = p.quantity;
    record['販売価格'] = p.price;
    record['小計'] = p.quantity * p.price;

    const addRecord = recordToArray(record, orderHeaders);
    records.push(addRecord);
    createRecords.push(record);
  }

  // 受注シートに追加
  addRecords('受注', records);

  // CSV作成（ヤマト/ヤマト伝票、佐川/佐川伝票に対応）
  const deliveryMethod = recurringOrder['納品方法'];
  if (deliveryMethod === 'ヤマト' || deliveryMethod === 'ヤマト伝票') {
    createYamatoCsvFromRecurring(recurringOrder, deliveryId);
  } else if (deliveryMethod === '佐川' || deliveryMethod === '佐川伝票') {
    createSagawaCsvFromRecurring(recurringOrder, deliveryId);
  }

  // 納品書作成
  if (recurringOrder['納品書'] === '○') {
    createFile(createRecords);
  }

  // 領収書作成
  if (recurringOrder['領収書'] === '○') {
    createReceiptFile(createRecords);
  }

  Logger.log('定期便から受注を作成しました: ' + deliveryId);
}

/**
 * 定期便からヤマトCSVを作成
 * orderCode.jsのaddRecordYamato()と同じフォーマットで出力
 *
 * @param {Object} recurringOrder - 定期受注データ
 * @param {string} deliveryId - 受注ID
 */
function createYamatoCsvFromRecurring(recurringOrder, deliveryId) {
  // マスタデータを取得してコード変換
  const invoiceTypes = getAllRecords('送り状種別') || [];
  const coolClss = getAllRecords('クール区分') || [];
  const deliveryTimes = getAllRecords('配送時間帯') || [];
  const cargos = getAllRecords('荷扱い') || [];

  // 表示名からコードを取得するヘルパー関数
  function getCodeFromMaster(masterData, displayName, keyField) {
    keyField = keyField || '種別値';
    for (var i = 0; i < masterData.length; i++) {
      var item = masterData[i];
      var value = item[keyField] || '';
      // "0:発払い" 形式の場合
      if (value.indexOf(':') !== -1) {
        var parts = value.split(':');
        if (parts[1] === displayName) {
          return parts[0];
        }
      }
      // 名前が一致する場合
      if (item['名前'] === displayName || item['種別名'] === displayName) {
        if (value.indexOf(':') !== -1) {
          return value.split(':')[0];
        }
        return value;
      }
    }
    return '';
  }

  // 日付フォーマット
  var shippingDate = recurringOrder['次回発送日'];
  var deliveryDate = recurringOrder['次回納品日'];
  try {
    shippingDate = Utilities.formatDate(new Date(shippingDate), 'JST', 'yyyy/MM/dd');
    deliveryDate = Utilities.formatDate(new Date(deliveryDate), 'JST', 'yyyy/MM/dd');
  } catch (e) {
    Logger.log('日付フォーマットエラー: ' + e.message);
  }

  var yamatoRecord = {};
  yamatoRecord['発送日'] = shippingDate;
  yamatoRecord['お客様管理番号'] = deliveryId || '';
  yamatoRecord['送り状種別'] = getCodeFromMaster(invoiceTypes, recurringOrder['送り状種別']);
  yamatoRecord['クール区分'] = getCodeFromMaster(coolClss, recurringOrder['クール区分']);
  yamatoRecord['伝票番号'] = '';
  yamatoRecord['出荷予定日'] = shippingDate;
  yamatoRecord['お届け予定（指定）日'] = deliveryDate;
  yamatoRecord['配達時間帯'] = getCodeFromMaster(deliveryTimes, recurringOrder['配達時間帯'], '時間指定');
  yamatoRecord['お届け先コード'] = '';
  yamatoRecord['お届け先電話番号'] = recurringOrder['発送先電話番号'] || '';
  yamatoRecord['お届け先電話番号枝番'] = '';
  yamatoRecord['お届け先郵便番号'] = recurringOrder['発送先郵便番号'] || '';
  // ヤマトB2: お届け先住所・お届け先住所（アパートマンション名）は各16文字まで。超える場合は16文字目以降を次カラムへ
  (function () {
    const ADDRESS_MAX_LEN = 16;
    var toAddr = (recurringOrder['発送先住所'] || '').toString();
    yamatoRecord['お届け先住所'] = toAddr.length > ADDRESS_MAX_LEN ? toAddr.substring(0, ADDRESS_MAX_LEN) : toAddr;
    yamatoRecord['お届け先住所（アパートマンション名）'] = toAddr.length > ADDRESS_MAX_LEN ? toAddr.substring(ADDRESS_MAX_LEN) : '';
  })();
  yamatoRecord['お届け先会社・部門名１'] = '';
  yamatoRecord['お届け先会社・部門名２'] = '';
  yamatoRecord['お届け先名'] = recurringOrder['発送先名'] || '';
  yamatoRecord['お届け先名略称カナ'] = '';
  yamatoRecord['敬称'] = '';
  yamatoRecord['ご依頼主コード'] = '';
  yamatoRecord['ご依頼主電話番号'] = recurringOrder['発送元電話番号'] || '';
  yamatoRecord['ご依頼主電話番号枝番'] = '';
  yamatoRecord['ご依頼主郵便番号'] = recurringOrder['発送元郵便番号'] || '';
  // ヤマトB2: ご依頼主住所・ご依頼主住所（アパートマンション名）は各16文字まで。超える場合は16文字目以降を次カラムへ
  (function () {
    const ADDRESS_MAX_LEN = 16;
    var fromAddr = (recurringOrder['発送元住所'] || '').toString();
    yamatoRecord['ご依頼主住所'] = fromAddr.length > ADDRESS_MAX_LEN ? fromAddr.substring(0, ADDRESS_MAX_LEN) : fromAddr;
    yamatoRecord['ご依頼主住所（アパートマンション名）'] = fromAddr.length > ADDRESS_MAX_LEN ? fromAddr.substring(ADDRESS_MAX_LEN) : '';
  })();
  yamatoRecord['ご依頼主名'] = recurringOrder['発送元名'] || '';
  yamatoRecord['ご依頼主略称カナ'] = '';
  yamatoRecord['品名コード１'] = '';
  yamatoRecord['品名１'] = recurringOrder['品名'] || '';
  yamatoRecord['品名コード２'] = '';
  yamatoRecord['品名２'] = '';
  yamatoRecord['荷扱い１'] = getCodeFromMaster(cargos, recurringOrder['荷扱い１']);
  yamatoRecord['荷扱い２'] = getCodeFromMaster(cargos, recurringOrder['荷扱い２']);
  yamatoRecord['記事'] = recurringOrder['送り状備考欄'] || '';
  yamatoRecord['コレクト代金引換額（税込）'] = recurringOrder['代引総額'] || '';
  yamatoRecord['コレクト内消費税額等'] = recurringOrder['代引内税'] || '';
  yamatoRecord['営業所止置き'] = '';
  yamatoRecord['営業所コード'] = '';
  yamatoRecord['発行枚数'] = recurringOrder['発行枚数'] || '1';
  yamatoRecord['個数口枠の印字'] = '';
  yamatoRecord['ご請求先顧客コード'] = '019543385101';
  yamatoRecord['ご請求先分類コード'] = '';
  yamatoRecord['運賃管理番号'] = '01';
  yamatoRecord['クロネコwebコレクトデータ登録'] = '';
  yamatoRecord['クロネコwebコレクト加盟店番号'] = '';
  yamatoRecord['クロネコwebコレクト申込受付番号１'] = '';
  yamatoRecord['クロネコwebコレクト申込受付番号２'] = '';
  yamatoRecord['クロネコwebコレクト申込受付番号３'] = '';
  yamatoRecord['お届け予定ｅメール利用区分'] = '';
  yamatoRecord['お届け予定ｅメールe-mailアドレス'] = '';
  yamatoRecord['入力機種'] = '';
  yamatoRecord['お届け予定eメールメッセージ'] = '';

  // orderCode.jsと同じ順序でCSV行を作成
  var addList = [
    yamatoRecord['発送日'],
    yamatoRecord['お客様管理番号'],
    yamatoRecord['送り状種別'],
    yamatoRecord['クール区分'],
    yamatoRecord['伝票番号'],
    yamatoRecord['出荷予定日'],
    yamatoRecord['お届け予定（指定）日'],
    yamatoRecord['配達時間帯'],
    yamatoRecord['お届け先コード'],
    yamatoRecord['お届け先電話番号'],
    yamatoRecord['お届け先電話番号枝番'],
    yamatoRecord['お届け先郵便番号'],
    yamatoRecord['お届け先住所'],
    yamatoRecord['お届け先住所（アパートマンション名）'],
    yamatoRecord['お届け先会社・部門名１'],
    yamatoRecord['お届け先会社・部門名２'],
    yamatoRecord['お届け先名'],
    yamatoRecord['お届け先名略称カナ'],
    yamatoRecord['敬称'],
    yamatoRecord['ご依頼主コード'],
    yamatoRecord['ご依頼主電話番号'],
    yamatoRecord['ご依頼主電話番号枝番'],
    yamatoRecord['ご依頼主郵便番号'],
    yamatoRecord['ご依頼主住所'],
    yamatoRecord['ご依頼主住所（アパートマンション名）'],
    yamatoRecord['ご依頼主名'],
    yamatoRecord['ご依頼主略称カナ'],
    yamatoRecord['品名コード１'],
    yamatoRecord['品名１'],
    yamatoRecord['品名コード２'],
    yamatoRecord['品名２'],
    yamatoRecord['荷扱い１'],
    yamatoRecord['荷扱い２'],
    yamatoRecord['記事'],
    yamatoRecord['コレクト代金引換額（税込）'],
    yamatoRecord['コレクト内消費税額等'],
    yamatoRecord['営業所止置き'],
    yamatoRecord['営業所コード'],
    yamatoRecord['発行枚数'],
    yamatoRecord['個数口枠の印字'],
    yamatoRecord['ご請求先顧客コード'],
    yamatoRecord['ご請求先分類コード'],
    yamatoRecord['運賃管理番号'],
    yamatoRecord['クロネコwebコレクトデータ登録'],
    yamatoRecord['クロネコwebコレクト加盟店番号'],
    yamatoRecord['クロネコwebコレクト申込受付番号１'],
    yamatoRecord['クロネコwebコレクト申込受付番号２'],
    yamatoRecord['クロネコwebコレクト申込受付番号３'],
    yamatoRecord['お届け予定ｅメール利用区分'],
    yamatoRecord['お届け予定ｅメールe-mailアドレス'],
    yamatoRecord['入力機種'],
    yamatoRecord['お届け予定eメールメッセージ']
  ];

  addRecords('ヤマトCSV', [addList]);
  Logger.log('ヤマトCSV作成（定期便）: ' + deliveryId);
}

/**
 * 定期便から佐川CSVを作成
 * orderCode.jsのaddRecordSagawa()と同じフォーマットで出力
 *
 * @param {Object} recurringOrder - 定期受注データ
 * @param {string} deliveryId - 受注ID
 */
function createSagawaCsvFromRecurring(recurringOrder, deliveryId) {
  // マスタデータを取得してコード変換
  var coolClss = getAllRecords('クール区分') || [];
  var deliveryTimes = getAllRecords('配送時間帯') || [];
  var cargos = getAllRecords('荷扱い') || [];
  var invoiceTypes = getAllRecords('送り状種別') || [];

  // 表示名からコードを取得するヘルパー関数
  function getCodeFromMaster(masterData, displayName, keyField) {
    keyField = keyField || '種別値';
    for (var i = 0; i < masterData.length; i++) {
      var item = masterData[i];
      var value = item[keyField] || '';
      // "0:発払い" 形式の場合
      if (value.indexOf(':') !== -1) {
        var parts = value.split(':');
        if (parts[1] === displayName) {
          return parts[0];
        }
      }
      // 名前が一致する場合
      if (item['名前'] === displayName || item['種別名'] === displayName) {
        if (value.indexOf(':') !== -1) {
          return value.split(':')[0];
        }
        return value;
      }
    }
    return '';
  }

  // 日付フォーマット
  var shippingDate = recurringOrder['次回発送日'];
  var deliveryDate = recurringOrder['次回納品日'];
  try {
    shippingDate = Utilities.formatDate(new Date(shippingDate), 'JST', 'yyyy/MM/dd');
    deliveryDate = Utilities.formatDate(new Date(deliveryDate), 'JST', 'yyyyMMdd');
  } catch (e) {
    Logger.log('日付フォーマットエラー: ' + e.message);
  }

  // 品名を16文字に制限
  var productName = recurringOrder['品名'] || '';
  if (productName.length > 16) {
    productName = productName.substring(0, 16);
  }

  var sagawaRecord = {};
  sagawaRecord['発送日'] = shippingDate;
  sagawaRecord['お届け先コード取得区分'] = '';
  sagawaRecord['お届け先コード'] = '';
  sagawaRecord['お届け先電話番号'] = recurringOrder['発送先電話番号'] || '';
  sagawaRecord['お届け先郵便番号'] = recurringOrder['発送先郵便番号'] || '';
  sagawaRecord['お届け先住所１'] = recurringOrder['発送先住所'] || '';
  sagawaRecord['お届け先住所２'] = '';
  sagawaRecord['お届け先住所３'] = '';
  sagawaRecord['お届け先名称１'] = recurringOrder['発送先名'] || '';
  sagawaRecord['お届け先名称２'] = '';
  sagawaRecord['お客様管理番号'] = deliveryId || '';
  sagawaRecord['お客様コード'] = '';
  sagawaRecord['部署ご担当者コード取得区分'] = '';
  sagawaRecord['部署ご担当者コード'] = '';
  sagawaRecord['部署ご担当者名称'] = '';
  sagawaRecord['荷送人電話番号'] = '';
  sagawaRecord['ご依頼主コード取得区分'] = '';
  sagawaRecord['ご依頼主コード'] = '';
  sagawaRecord['ご依頼主電話番号'] = recurringOrder['発送元電話番号'] || '';
  sagawaRecord['ご依頼主郵便番号'] = recurringOrder['発送元郵便番号'] || '';
  sagawaRecord['ご依頼主住所１'] = recurringOrder['発送元住所'] || '';
  sagawaRecord['ご依頼主住所２'] = '';
  sagawaRecord['ご依頼主名称１'] = recurringOrder['発送元名'] || '';
  sagawaRecord['ご依頼主名称２'] = '';
  sagawaRecord['荷姿'] = '';
  sagawaRecord['品名１'] = productName;
  sagawaRecord['品名２'] = '';
  sagawaRecord['品名３'] = '';
  sagawaRecord['品名４'] = '';
  sagawaRecord['品名５'] = '';
  sagawaRecord['荷札荷姿'] = '';
  sagawaRecord['荷札品名１'] = '';
  sagawaRecord['荷札品名２'] = '';
  sagawaRecord['荷札品名３'] = '';
  sagawaRecord['荷札品名４'] = '';
  sagawaRecord['荷札品名５'] = '';
  sagawaRecord['荷札品名６'] = '';
  sagawaRecord['荷札品名７'] = '';
  sagawaRecord['荷札品名８'] = '';
  sagawaRecord['荷札品名９'] = '';
  sagawaRecord['荷札品名１０'] = '';
  sagawaRecord['荷札品名１１'] = '';
  sagawaRecord['出荷個数'] = recurringOrder['発行枚数'] || '1';
  sagawaRecord['スピード指定'] = '000';
  sagawaRecord['クール便指定'] = getCodeFromMaster(coolClss, recurringOrder['クール区分']);
  sagawaRecord['配達日'] = deliveryDate;
  sagawaRecord['配達指定時間帯'] = getCodeFromMaster(deliveryTimes, recurringOrder['配達時間帯'], '時間指定');
  sagawaRecord['配達指定時間（時分）'] = '';
  sagawaRecord['代引金額'] = recurringOrder['代引総額'] || '';
  sagawaRecord['消費税'] = recurringOrder['代引内税'] || '';
  sagawaRecord['決済種別'] = '';
  sagawaRecord['保険金額'] = '';
  sagawaRecord['指定シール１'] = getCodeFromMaster(cargos, recurringOrder['荷扱い１']);
  sagawaRecord['指定シール２'] = getCodeFromMaster(cargos, recurringOrder['荷扱い２']);
  sagawaRecord['指定シール３'] = getCodeFromMaster(cargos, recurringOrder['荷扱い３']);
  sagawaRecord['営業所受取'] = '';
  sagawaRecord['SRC区分'] = '';
  sagawaRecord['営業所受取営業所コード'] = '';
  sagawaRecord['元着区分'] = getCodeFromMaster(invoiceTypes, recurringOrder['送り状種別']);
  sagawaRecord['メールアドレス'] = '';
  sagawaRecord['ご不在時連絡先'] = recurringOrder['発送先電話番号'] || '';
  sagawaRecord['出荷予定日'] = '';
  sagawaRecord['セット数'] = '';
  sagawaRecord['お問い合せ送り状No.'] = '';
  sagawaRecord['出荷場印字区分'] = '';
  sagawaRecord['集約解除指定'] = '';
  sagawaRecord['編集０１'] = '';
  sagawaRecord['編集０２'] = '';
  sagawaRecord['編集０３'] = '';
  sagawaRecord['編集０４'] = '';
  sagawaRecord['編集０５'] = '';
  sagawaRecord['編集０６'] = '';
  sagawaRecord['編集０７'] = '';
  sagawaRecord['編集０８'] = '';
  sagawaRecord['編集０９'] = '';
  sagawaRecord['編集１０'] = '';

  // orderCode.jsと同じ順序でCSV行を作成
  var addList = [
    sagawaRecord['発送日'],
    sagawaRecord['お届け先コード取得区分'],
    sagawaRecord['お届け先コード'],
    sagawaRecord['お届け先電話番号'],
    sagawaRecord['お届け先郵便番号'],
    sagawaRecord['お届け先住所１'],
    sagawaRecord['お届け先住所２'],
    sagawaRecord['お届け先住所３'],
    sagawaRecord['お届け先名称１'],
    sagawaRecord['お届け先名称２'],
    sagawaRecord['お客様管理番号'],
    sagawaRecord['お客様コード'],
    sagawaRecord['部署ご担当者コード取得区分'],
    sagawaRecord['部署ご担当者コード'],
    sagawaRecord['部署ご担当者名称'],
    sagawaRecord['荷送人電話番号'],
    sagawaRecord['ご依頼主コード取得区分'],
    sagawaRecord['ご依頼主コード'],
    sagawaRecord['ご依頼主電話番号'],
    sagawaRecord['ご依頼主郵便番号'],
    sagawaRecord['ご依頼主住所１'],
    sagawaRecord['ご依頼主住所２'],
    sagawaRecord['ご依頼主名称１'],
    sagawaRecord['ご依頼主名称２'],
    sagawaRecord['荷姿'],
    sagawaRecord['品名１'],
    sagawaRecord['品名２'],
    sagawaRecord['品名３'],
    sagawaRecord['品名４'],
    sagawaRecord['品名５'],
    sagawaRecord['荷札荷姿'],
    sagawaRecord['荷札品名１'],
    sagawaRecord['荷札品名２'],
    sagawaRecord['荷札品名３'],
    sagawaRecord['荷札品名４'],
    sagawaRecord['荷札品名５'],
    sagawaRecord['荷札品名６'],
    sagawaRecord['荷札品名７'],
    sagawaRecord['荷札品名８'],
    sagawaRecord['荷札品名９'],
    sagawaRecord['荷札品名１０'],
    sagawaRecord['荷札品名１１'],
    sagawaRecord['出荷個数'],
    sagawaRecord['スピード指定'],
    sagawaRecord['クール便指定'],
    sagawaRecord['配達日'],
    sagawaRecord['配達指定時間帯'],
    sagawaRecord['配達指定時間（時分）'],
    sagawaRecord['代引金額'],
    sagawaRecord['消費税'],
    sagawaRecord['決済種別'],
    sagawaRecord['保険金額'],
    sagawaRecord['指定シール１'],
    sagawaRecord['指定シール２'],
    sagawaRecord['指定シール３'],
    sagawaRecord['営業所受取'],
    sagawaRecord['SRC区分'],
    sagawaRecord['営業所受取営業所コード'],
    sagawaRecord['元着区分'],
    sagawaRecord['メールアドレス'],
    sagawaRecord['ご不在時連絡先'],
    sagawaRecord['出荷予定日'],
    sagawaRecord['セット数'],
    sagawaRecord['お問い合せ送り状No.'],
    sagawaRecord['出荷場印字区分'],
    sagawaRecord['集約解除指定'],
    sagawaRecord['編集０１'],
    sagawaRecord['編集０２'],
    sagawaRecord['編集０３'],
    sagawaRecord['編集０４'],
    sagawaRecord['編集０５'],
    sagawaRecord['編集０６'],
    sagawaRecord['編集０７'],
    sagawaRecord['編集０８'],
    sagawaRecord['編集０９'],
    sagawaRecord['編集１０']
  ];

  addRecords('佐川CSV', [addList]);
  Logger.log('佐川CSV作成（定期便）: ' + deliveryId);
}


/**
 * 定期受注一覧をHTML用データとして取得
 * @returns {Array<Object>} 定期受注一覧データ
 */
function getRecurringOrderListData() {
  const orders = getRecurringOrders();

  // 日付を文字列に変換するヘルパー関数
  function formatDate(value) {
    if (!value) return '';
    if (value instanceof Date) {
      return Utilities.formatDate(value, 'JST', 'yyyy/MM/dd');
    }
    return String(value);
  }

  return orders.map(function(order) {
    // 商品サマリを作成
    const products = order['商品データ'] || [];
    const productSummary = products.map(function(p) {
      return p.product + ' x' + p.quantity;
    }).join(', ');

    // 合計金額を計算
    const totalAmount = products.reduce(function(sum, p) {
      return sum + (p.price * p.quantity);
    }, 0);

    return {
      id: order['定期受注ID'],
      registrationDate: formatDate(order['登録日']),
      interval: order['更新間隔'],
      nextShippingDate: formatDate(order['次回発送日']),
      nextDeliveryDate: formatDate(order['次回納品日']),
      status: order['ステータス'],
      lastExecutionDate: formatDate(order['最終実行日']),
      customerName: order['顧客名'],
      shippingToName: order['発送先名'],
      deliveryMethod: order['納品方法'],
      productSummary: productSummary,
      totalAmount: totalAmount,
      products: products
    };
  });
}

/**
 * 定期受注詳細を取得（編集用）
 * @param {string} recurringId - 定期受注ID
 * @returns {Object} 定期受注詳細データ
 */
function getRecurringOrderDetail(recurringId) {
  const order = getRecurringOrderById(recurringId);

  if (!order) {
    return null;
  }

  return {
    id: order['定期受注ID'],
    registrationDate: order['登録日'],
    interval: order['更新間隔'],
    nextShippingDate: order['次回発送日'],
    nextDeliveryDate: order['次回納品日'],
    status: order['ステータス'],
    lastExecutionDate: order['最終実行日'],
    customerName: order['顧客名'],
    customerZipcode: order['顧客郵便番号'],
    customerAddress: order['顧客住所'],
    customerTel: order['顧客電話番号'],
    shippingToName: order['発送先名'],
    shippingToZipcode: order['発送先郵便番号'],
    shippingToAddress: order['発送先住所'],
    shippingToTel: order['発送先電話番号'],
    shippingFromName: order['発送元名'],
    shippingFromZipcode: order['発送元郵便番号'],
    shippingFromAddress: order['発送元住所'],
    shippingFromTel: order['発送元電話番号'],
    receiptWay: order['受付方法'],
    recipient: order['受付者'],
    deliveryMethod: order['納品方法'],
    deliveryTime: order['配達時間帯'],
    checklist: {
      delivery: order['納品書'] === '○',
      invoice: order['請求書'] === '○',
      receipt: order['領収書'] === '○',
      pamphlet: order['パンフ'] === '○',
      recipe: order['レシピ'] === '○'
    },
    otherAttach: order['その他添付'],
    sendProduct: order['品名'],
    invoiceType: order['送り状種別'],
    coolCls: order['クール区分'],
    cargo1: order['荷扱い１'],
    cargo2: order['荷扱い２'],
    cargo3: order['荷扱い３'],
    cashOnDelivery: order['代引総額'],
    cashOnDeliTax: order['代引内税'],
    copiePrint: order['発行枚数'],
    internalMemo: order['社内メモ'],
    csvmemo: order['送り状備考欄'],
    deliveryMemo: order['納品書備考欄'],
    memo: order['メモ'],
    deliveryNoteText: order['納品書テキスト'],
    products: order['商品データ']
  };
}

/**
 * 定期受注を完全更新する（編集画面から）
 * @param {Object} e - フォームパラメータ
 * @returns {boolean} 成功/失敗
 */
function updateRecurringOrderFull(e) {
  const recurringId = e.parameter.recurringId;

  if (!recurringId) {
    Logger.log('定期受注IDが指定されていません');
    return false;
  }

  // 商品データをJSON化
  const products = [];
  for (let i = 1; i <= 10; i++) {
    const bunrui = e.parameter['bunrui' + i];
    const product = e.parameter['product' + i];
    const price = e.parameter['price' + i];
    const quantity = e.parameter['quantity' + i];

    if (quantity && Number(quantity) > 0) {
      products.push({
        bunrui: bunrui || '',
        product: product || '',
        price: Number(price) || 0,
        quantity: Number(quantity)
      });
    }
  }

  let intervalObj;
  try {
    intervalObj = JSON.parse(e.parameter.recurringInterval);
  } catch {
    intervalObj = { type: 'nmonth', value: Number(e.parameter.recurringInterval) || 1 };
  }
  const updates = {
    '更新間隔': JSON.stringify(intervalObj),
    '次回発送日': e.parameter.nextShippingDate || '',
    '次回納品日': e.parameter.nextDeliveryDate || '',
    '顧客名': e.parameter.customerName || '',
    '顧客郵便番号': e.parameter.customerZipcode || '',
    '顧客住所': e.parameter.customerAddress || '',
    '顧客電話番号': e.parameter.customerTel || '',
    '発送先名': e.parameter.shippingToName || '',
    '発送先郵便番号': e.parameter.shippingToZipcode || '',
    '発送先住所': e.parameter.shippingToAddress || '',
    '発送先電話番号': e.parameter.shippingToTel || '',
    '発送元名': e.parameter.shippingFromName || '',
    '発送元郵便番号': e.parameter.shippingFromZipcode || '',
    '発送元住所': e.parameter.shippingFromAddress || '',
    '発送元電話番号': e.parameter.shippingFromTel || '',
    '受付方法': e.parameter.receiptWay || '',
    '受付者': e.parameter.recipient || '',
    '納品方法': e.parameter.deliveryMethod || '',
    '配達時間帯': e.parameter.deliveryTime ? e.parameter.deliveryTime.split(':')[1] : '',
    '納品書': e.parameters.checklist && e.parameters.checklist.includes('納品書') ? '○' : '',
    '請求書': e.parameters.checklist && e.parameters.checklist.includes('請求書') ? '○' : '',
    '領収書': e.parameters.checklist && e.parameters.checklist.includes('領収書') ? '○' : '',
    'パンフ': e.parameters.checklist && e.parameters.checklist.includes('パンフ') ? '○' : '',
    'レシピ': e.parameters.checklist && e.parameters.checklist.includes('レシピ') ? '○' : '',
    'その他添付': e.parameter.otherAttach || '',
    '品名': e.parameter.sendProduct || '',
    '送り状種別': e.parameter.invoiceType ? e.parameter.invoiceType.split(':')[1] : '',
    'クール区分': e.parameter.coolCls ? e.parameter.coolCls.split(':')[1] : '',
    '荷扱い１': e.parameter.cargo1 ? e.parameter.cargo1.split(':')[1] : '',
    '荷扱い２': e.parameter.cargo2 ? e.parameter.cargo2.split(':')[1] : '',
    '荷扱い３': e.parameter.cargo3 ? e.parameter.cargo3.split(':')[1] : '',
    '代引総額': e.parameter.cashOnDelivery || '',
    '代引内税': e.parameter.cashOnDeliTax || '',
    '発行枚数': e.parameter.copiePrint || '',
    '社内メモ': e.parameter.internalMemo || '',
    '送り状備考欄': e.parameter.csvmemo || '',
    '納品書備考欄': e.parameter.deliveryMemo || '',
    'メモ': e.parameter.memo || '',
    '納品書テキスト': e.parameter.deliveryNoteText || '',
    '商品データJSON': JSON.stringify(products)
  };

  return updateRecurringOrder(recurringId, updates);
}

/**
 * トリガーを設定する（初回セットアップ用）
 */
function setupRecurringOrderTrigger() {
  // 既存のトリガーを削除
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'processRecurringOrders') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  // 毎日午前6時に実行するトリガーを設定
  ScriptApp.newTrigger('processRecurringOrders')
    .timeBased()
    .atHour(6)
    .everyDays(1)
    .create();

  Logger.log('定期受注処理トリガーを設定しました（毎日午前6時）');
}
