/**
 * ホーム画面の発送予定日選択処理（テスト用）
 * 内部で selectShippingDate() を呼び出します。
 */
function test8() {
  selectShippingDate();
}

/**
 * 発送予定日ごとの受注サマリーデータを取得
 * ホーム画面のダッシュボード表示に使用されます。
 * 
 * @returns {string} - JSON形式の日付ごとの受注データ
 */
function selectShippingDate() {
  const dates = calculateDates();
  const dateStrings = createDateStrings(dates);
  const dataStructure = initializeDataStructure();

  // 商品単位マップの作成
  const masterData = getMasterDataCached();
  const unitMap = {};
  if (masterData && masterData.products) {
    masterData.products.forEach(p => {
      unitMap[p['商品名']] = p['単位'] || '枚';
    });
  }

  // 昨日〜明々後日 + 発送日超過（未出荷・非キャンセルで発送日が今日より前）を取得
  const items = getOrdersByDateRange(dates.yesterday, dates.dayAfter3, true);
  processItems(items, dates, dataStructure, unitMap);

  return JSON.stringify(buildResult(dateStrings, dataStructure, dates));
}
/**
 * 注文のステータスを更新（発送前/収穫待ち/出荷済み）
 * @param {string} orderId - 受注ID
 * @param {string} newStatus - 新しいステータス
 */
function updateOrderStatus(orderId, newStatus) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return updateOrderStatusShared(ss, orderId, newStatus);
}

/**
 * 指定された日付範囲の受注レコードを取得
 * 受注一覧の getOrderListData と同様に、includeOverdue 時は「期間内」OR「予定日超過」のみを含める（期間を一括で拡張しない）。
 *
 * @param {Date} startDate - 開始日
 * @param {Date} endDate - 終了日
 * @param {boolean} [includeOverdue=false] - true の場合、発送日が今日より前で未出荷・非キャンセルの行も含める（当日分として表示するため）
 * @returns {Array<Object>} - ラベルをキーとしたレコードの配列
 * @description 収穫待ちの受注は発送日が範囲外（未来含む）でも常に含める（件数確認の収穫待ち欄に表示するため）
 */
function getOrdersByDateRange(startDate, endDate, includeOverdue) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('受注');
  const values = sheet.getDataRange().getValues();
  const labels = values.shift();

  const dateColIndex = labels.indexOf('発送日');
  const statusColIndex = labels.indexOf('ステータス');
  const shippedColIndex = labels.indexOf('出荷済');

  const startTime = startDate.getTime();
  const endTime = endDate.getTime() + 24 * 60 * 60 * 1000; // 終了日の翌日0時

  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const todayTime = today.getTime();

  const records = [];
  for (const value of values) {
    const status = statusColIndex !== -1 ? value[statusColIndex] : '';
    const isCancelled = status === 'キャンセル' || status === 'cancelled';
    if (isCancelled) continue;

    const rawDate = value[dateColIndex];

    let inRange = false;
    let isOverdueRow = false;
    const isHarvestWaiting = status === '収穫待ち';

    const shippingDate = new Date(rawDate);
    const shippingTime = shippingDate.getTime();
    const isShipped = shippedColIndex !== -1 && (value[shippedColIndex] === '○' || value[shippedColIndex] === true);

    // 期間内チェック
    if (!isNaN(shippingTime)) {
      inRange = shippingTime >= startTime && shippingTime < endTime;
      // 超過判定: 発送日が今日より前 かつ 未出荷
      isOverdueRow = includeOverdue && !isShipped && shippingTime < todayTime;
    }

    // 収穫待ちでも出荷済みの場合は除外
    if (isHarvestWaiting && isShipped) continue;

    // 期間内 または 予定日超過 または 収穫待ち（発送日が先でも収穫待ち欄に表示するため）
    if (!inRange && !isOverdueRow && !isHarvestWaiting) continue;

    const record = {};
    labels.forEach((label, index) => {
      record[label] = value[index];
    });
    // 超過・収穫待ちフラグを付与
    if (isOverdueRow) {
      record._isOverdue = true;
    }
    if (isHarvestWaiting) {
      record._isHarvestWaiting = true;
    }
    records.push(record);
  }

  return records;
}

/**
 * ダッシュボード表示に必要な各日程（昨日、今日、明日、明後日、明々後日）を計算
 * 
 * @returns {Object} - 日付オブジェクトと yyyy/MM/dd 形式の文字列を含むオブジェクト
 */
function calculateDates() {
  const nDate = new Date();
  const dates = {
    yesterday: new Date(nDate.getTime() - 24 * 60 * 60 * 1000),
    today: new Date(nDate),
    tomorrow: new Date(nDate.getTime() + 24 * 60 * 60 * 1000),
    dayAfter2: new Date(nDate.getTime() + 2 * 24 * 60 * 60 * 1000),
    dayAfter3: new Date(nDate.getTime() + 3 * 24 * 60 * 60 * 1000)
  };

  // 文字列形式も事前に計算
  Object.keys(dates).forEach(key => {
    if (!key.endsWith('Str')) {
      dates[key + 'Str'] = Utilities.formatDate(dates[key], 'JST', 'yyyy/MM/dd');
    }
  });

  return dates;
}

/**
 * 表示用の曜日付き日付文字列を作成
 * 
 * @param {Object} dates - calculateDates() で作成された日付オブジェクト
 * @returns {Object} - 「MM/dd (曜日)」形式の文字列
 */
function createDateStrings(dates) {
  const weekdays = ['日', '月', '火', '水', '木', '金', '土'];

  return {
    '昨日': formatDateWithWeekday(dates.yesterday, weekdays),
    '今日': formatDateWithWeekday(dates.today, weekdays),
    '明日': formatDateWithWeekday(dates.tomorrow, weekdays),
    '明後日': formatDateWithWeekday(dates.dayAfter2, weekdays),
    '明々後日': formatDateWithWeekday(dates.dayAfter3, weekdays)
  };
}

/**
 * 日付を MM/dd (曜日) 形式にフォーマット
 * 
 * @param {Date} date - 日付オブジェクト
 * @param {Array<string>} weekdays - 曜日の配列（['日', '月', ...]）
 * @returns {string} - フォーマット済み文字列
 */
function formatDateWithWeekday(date, weekdays) {
  const formattedDate = Utilities.formatDate(date, 'JST', 'MM/dd');
  const weekday = weekdays[date.getDay()];
  return `${formattedDate} (${weekday})`;
}

/**
 * 集計用データ構造を初期化
 * 
 * @returns {Object} - 空の Map とカウントを含む初期構造
 */
function initializeDataStructure() {
  return {
    lists: {
      yesterday: new Map(),
      today: new Map(),
      tomorrow: new Map(),
      dayAfter2: new Map(),
      dayAfter3: new Map()
    },
    invoiceTypes: {
      yesterday: new Map(),
      today: new Map(),
      tomorrow: new Map(),
      dayAfter2: new Map(),
      dayAfter3: new Map()
    },
    counts: {
      yesterday: 0,
      today: 0,
      tomorrow: 0,
      dayAfter2: 0,
      dayAfter3: 0
    },
    // 超過・収穫待ち件数カウント用（UI表示用）
    overdueCounts: {
      today: 0
    },
    harvestWaitingCounts: {
      total: 0
    },
    /** 収穫待ちのみを集約するリスト（日付別リストには含めない） */
    harvestWaitingLists: new Map()
  };
}

/**
 * 取得した受注アイテムを日付ごとに分類して処理
 * 
 * @param {Array<Object>} items - 受注レコードの配列
 * @param {Object} dates - 日程オブジェクト
 * @param {Object} dataStructure - 集計用データ構造
 */
function processItems(items, dates, dataStructure, unitMap) {
  // 逆順でループ（新しいデータから処理）
  for (let i = items.length - 1; i >= 0; i--) {
    const item = items[i];

    if (item['商品名'] === '送料') continue;

    // 収穫待ちは日付別リストに含めず、収穫待ち専用リストにのみ追加（当日発送と混同しないため）
    if (item._isHarvestWaiting) {
      dataStructure.harvestWaitingCounts.total++;
      processHarvestWaitingItem(item, dataStructure, unitMap);
      continue;
    }

    let dateKey = null;

    const itemDate = Utilities.formatDate(new Date(item['発送日']), 'JST', 'yyyy/MM/dd');

    // 予定日超過判定（受注一覧と同様：発送日が今日より前かつ未出荷かつ非キャンセル → 当日分として表示）
    const isCancelled = item['ステータス'] === 'キャンセル' || item['ステータス'] === 'cancelled';
    const isShipped = item['出荷済'] === '○' || item['出荷済'] === true;
    const isOverdue = !isShipped && !isCancelled && itemDate < dates.todayStr;

    if (isOverdue) {
      dateKey = 'today';
      // 超過件数をカウント（別集計）
      dataStructure.overdueCounts.today++;
    } else if (itemDate === dates.yesterdayStr) {
      dateKey = 'yesterday';
    } else if (itemDate === dates.todayStr) {
      dateKey = 'today';
    } else if (itemDate === dates.tomorrowStr) {
      dateKey = 'tomorrow';
    } else if (itemDate === dates.dayAfter2Str) {
      dateKey = 'dayAfter2';
    } else if (itemDate === dates.dayAfter3Str) {
      dateKey = 'dayAfter3';
    }

    if (dateKey) {
      processDateItem(item, dateKey, dataStructure, unitMap);
    }
  }
}

/**
 * 特定の日付における受注アイテムの集計処理
 * 
 * @param {Object} item - 受注レコード単体
 * @param {string} dateKey - 日付キー（yesterday/today/...)
 * @param {Object} dataStructure - 集計用データ構造
 */
function processDateItem(item, dateKey, dataStructure, unitMap) {
  const lists = dataStructure.lists[dateKey];
  const invoiceTypes = dataStructure.invoiceTypes[dateKey];

  // キーの作成（発送先名とクール区分の組み合わせ）
  const listKey = `${item['発送先名']}_${item['クール区分']}_${item['受注ID']}`;

  // リストの処理
  if (lists.has(listKey)) {
    const record = lists.get(listKey);

    // 出荷済ステータスの同期（1つでも未出荷なら未出荷とするか、1つでも出荷済なら出荷済とするか）
    // 通常は一括出荷なので、論理和（OR）で更新
    record.shipped = record.shipped || item['出荷済'] === '○' || item['出荷済'] === true;

    const existingProduct = record['商品'].find(p => p['商品名'] === item['商品名']);

    if (existingProduct) {
      existingProduct['受注数'] += Number(item['受注数']);
    } else {
      record['商品'].push({
        '商品名': item['商品名'],
        '受注数': Number(item['受注数']),
        '単位': unitMap[item['商品名']] || '枚',
        'メモ': item['メモ'] || ""
      });
    }

    // メモの集約（重複排除）
    if (item['メモ'] && String(item['メモ']).trim() !== "") {
      record.uniqueMemos.add(String(item['メモ']).trim());
    }

  } else {
    // 納品方法の集約処理
    let deliveryMethod = item['納品方法'];

    // 伝票受領の集約
    if (deliveryMethod === 'ヤマト伝票受領') {
      deliveryMethod = 'ヤマト';
    } else if (deliveryMethod === '佐川伝票受領') {
      deliveryMethod = '佐川';
    }

    // キーの作成（ヤマト・佐川はクール区分付き、それ以外は納品方法のみ）
    let invoiceKey;
    if (deliveryMethod === 'ヤマト' || deliveryMethod === '佐川') {
      invoiceKey = `${deliveryMethod}_${item['クール区分']}`;
    } else {
      invoiceKey = deliveryMethod;
    }

    // 発行枚数を加算（件数ではなく）
    const shippingCount = Number(item['発行枚数']) || 1;

    if (invoiceTypes.has(invoiceKey)) {
      invoiceTypes.get(invoiceKey)['個数'] += shippingCount;
    } else {
      invoiceTypes.set(invoiceKey, {
        '納品方法': deliveryMethod,
        'クール区分': (deliveryMethod === 'ヤマト' || deliveryMethod === '佐川') ? item['クール区分'] : '',
        '個数': shippingCount
      });
    }

    // 合計カウントも発行枚数ベースに
    dataStructure.counts[dateKey]++;

    // メモ初期化
    const uniqueMemos = new Set();
    if (item['メモ'] && String(item['メモ']).trim() !== "") {
      uniqueMemos.add(String(item['メモ']).trim());
    }

    // 新規レコード作成（発送日は超過を先頭表示するソート用に保持）
    lists.set(listKey, {
      'orderId': item['受注ID'],
      '顧客名': item['顧客名'],
      '発送先名': item['発送先名'],
      'クール区分': item['クール区分'],
      'deliveryMethod': deliveryMethod,
      'trackingNumber': item['追跡番号'] || '',
      'status': item['ステータス'] || '',
      'shipped': item['出荷済'] === '○' || item['出荷済'] === true,
      'shippingDate': item['発送日'],
      'orderDate': item['受注日'],
      '商品': [{
        '商品名': item['商品名'],
        '受注数': Number(item['受注数']),
        '単位': unitMap[item['商品名']] || '枚',
        'メモ': item['メモ'] || ""
      }],
      'uniqueMemos': uniqueMemos // Setとして保持
    });
  }
}

/**
 * 収穫待ちの受注のみを harvestWaitingLists に集約（日付別リスト・件数には含めない）
 * @param {Object} item - 受注レコード単体
 * @param {Object} dataStructure - 集計用データ構造
 * @param {Object} unitMap - 商品単位マップ
 */
function processHarvestWaitingItem(item, dataStructure, unitMap) {
  const lists = dataStructure.harvestWaitingLists;
  const listKey = `${item['発送先名']}_${item['クール区分']}_${item['受注ID']}`;

  if (lists.has(listKey)) {
    const record = lists.get(listKey);
    record.shipped = record.shipped || item['出荷済'] === '○' || item['出荷済'] === true;
    const existingProduct = record['商品'].find(p => p['商品名'] === item['商品名']);
    if (existingProduct) {
      existingProduct['受注数'] += Number(item['受注数']);
    } else {
      record['商品'].push({
        '商品名': item['商品名'],
        '受注数': Number(item['受注数']),
        '単位': unitMap[item['商品名']] || '枚',
        'メモ': item['メモ'] || ""
      });
    }
    if (item['メモ'] && String(item['メモ']).trim() !== "") {
      record.uniqueMemos.add(String(item['メモ']).trim());
    }
  } else {
    let deliveryMethod = item['納品方法'];
    if (deliveryMethod === 'ヤマト伝票受領') deliveryMethod = 'ヤマト';
    else if (deliveryMethod === '佐川伝票受領') deliveryMethod = '佐川';

    const uniqueMemos = new Set();
    if (item['メモ'] && String(item['メモ']).trim() !== "") uniqueMemos.add(String(item['メモ']).trim());

    lists.set(listKey, {
      'orderId': item['受注ID'],
      '顧客名': item['顧客名'],
      '発送先名': item['発送先名'],
      'クール区分': item['クール区分'],
      'deliveryMethod': deliveryMethod,
      'trackingNumber': item['追跡番号'] || '',
      'status': item['ステータス'] || '',
      'shipped': item['出荷済'] === '○' || item['出荷済'] === true,
      'shippingDate': item['発送日'],
      'orderDate': item['受注日'],
      '商品': [{
        '商品名': item['商品名'],
        '受注数': Number(item['受注数']),
        '単位': unitMap[item['商品名']] || '枚',
        'メモ': item['メモ'] || ""
      }],
      'uniqueMemos': uniqueMemos
    });
  }
}

/**
 * 集計済みデータ構造から最終的なレスポンス配列を構築
 *
 * @param {Object} dateStrings - 表示用日付文字列
 * @param {Object} dataStructure - 集計済みデータ構造
 * @param {Object} [dates] - 日程オブジェクト（今日リストで超過を先頭ソートする場合に使用）
 * @returns {Array} - クライアント側に返すデータの配列
 */
function buildResult(dateStrings, dataStructure, dates) {
  const cnt = {
    '昨日': dataStructure.counts.yesterday,
    '今日': dataStructure.counts.today,
    '明日': dataStructure.counts.tomorrow,
    '明後日': dataStructure.counts.dayAfter2,
    '明々後日': dataStructure.counts.dayAfter3,
    '超過': dataStructure.overdueCounts.today,
    '収穫待ち': dataStructure.harvestWaitingCounts.total
  };

  // Mapを配列に変換（ソート機能付き）
  const convertMapToArray = (map) => {
    const array = Array.from(map.values());

    // ソート順を定義
    const sortOrder = {
      'ヤマト_冷蔵': 1,
      'ヤマト_冷凍': 2,
      'ヤマト_常温': 3,
      '佐川_冷蔵': 4,
      '佐川_冷凍': 5,
      '佐川_常温': 6,
      '西濃運輸': 7,
      '配達': 8,
      '店舗受取': 9,
      'レターパック': 10,
      'その他': 99
    };

    return array.sort((a, b) => {
      const keyA = a['クール区分'] ? `${a['納品方法']}_${a['クール区分']}` : a['納品方法'];
      const keyB = b['クール区分'] ? `${b['納品方法']}_${b['クール区分']}` : b['納品方法'];
      return (sortOrder[keyA] || 99) - (sortOrder[keyB] || 99);
    });
  };

  const todayStart = dates ? (function () {
    const t = new Date(dates.today.getTime());
    t.setHours(0, 0, 0, 0);
    return t.getTime();
  })() : null;

  const convertListMap = (map, sortOverdueFirst, sortByOrderDate) => {
    let arr = Array.from(map.values()).map(record => ({
      'orderId': record['orderId'],
      '顧客名': record['顧客名'],
      '発送先名': record['発送先名'],
      'クール区分': record['クール区分'],
      'deliveryMethod': record['deliveryMethod'],
      'trackingNumber': record['trackingNumber'],
      'status': record['status'],
      'shipped': record['shipped'],
      'shippingDate': record['shippingDate'] || null,
      'orderDate': record['orderDate'] || null,
      '商品': record['商品'],
      'aggregatedMemos': Array.from(record.uniqueMemos || [])
    }));
    if (sortOverdueFirst && todayStart != null) {
      arr = arr.sort((a, b) => {
        const aTime = a.shippingDate ? new Date(a.shippingDate).getTime() : 0;
        const bTime = b.shippingDate ? new Date(b.shippingDate).getTime() : 0;
        const aOver = aTime > 0 && aTime < todayStart;
        const bOver = bTime > 0 && bTime < todayStart;
        if (aOver && !bOver) return -1;
        if (!aOver && bOver) return 1;
        if (aOver && bOver) return aTime - bTime;
        return 0;
      });
    } else if (sortByOrderDate) {
      arr = arr.sort((a, b) => {
        const aTime = a.orderDate ? new Date(a.orderDate).getTime() : Number.MAX_SAFE_INTEGER;
        const bTime = b.orderDate ? new Date(b.orderDate).getTime() : Number.MAX_SAFE_INTEGER;
        return aTime - bTime;
      });
    }
    return arr;
  };

  return [
    dateStrings,
    cnt,
    convertMapToArray(dataStructure.invoiceTypes.yesterday),
    convertListMap(dataStructure.lists.yesterday),
    convertMapToArray(dataStructure.invoiceTypes.today),
    convertListMap(dataStructure.lists.today, true),
    convertMapToArray(dataStructure.invoiceTypes.tomorrow),
    convertListMap(dataStructure.lists.tomorrow),
    convertMapToArray(dataStructure.invoiceTypes.dayAfter2),
    convertListMap(dataStructure.lists.dayAfter2),
    convertMapToArray(dataStructure.invoiceTypes.dayAfter3),
    convertListMap(dataStructure.lists.dayAfter3),
    convertListMap(dataStructure.harvestWaitingLists, false, true)
  ];
}
