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

  // 必要な日付範囲のみ取得
  const items = getOrdersByDateRange(dates.yesterday, dates.dayAfter3);
  processItems(items, dates, dataStructure);

  return JSON.stringify(buildResult(dateStrings, dataStructure));
}

/**
 * 指定された日付範囲の受注レコードを取得
 * 
 * @param {Date} startDate - 開始日
 * @param {Date} endDate - 終了日
 * @returns {Array<Object>} - ラベルをキーとしたレコードの配列
 */
function getOrdersByDateRange(startDate, endDate) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('受注');
  const values = sheet.getDataRange().getValues();
  const labels = values.shift();

  // 発送日の列インデックスを取得
  const dateColIndex = labels.indexOf('発送日');

  const startTime = startDate.getTime();
  const endTime = endDate.getTime() + 24 * 60 * 60 * 1000; // 終了日の翌日0時

  const records = [];
  for (const value of values) {
    const shippingDate = new Date(value[dateColIndex]);
    const shippingTime = shippingDate.getTime();

    // 日付範囲内のみ追加
    if (shippingTime >= startTime && shippingTime < endTime) {
      const record = {};
      labels.forEach((label, index) => {
        record[label] = value[index];
      });
      records.push(record);
    }
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
    }
  };
}

/**
 * 取得した受注アイテムを日付ごとに分類して処理
 * 
 * @param {Array<Object>} items - 受注レコードの配列
 * @param {Object} dates - 日程オブジェクト
 * @param {Object} dataStructure - 集計用データ構造
 */
function processItems(items, dates, dataStructure) {
  // 逆順でループ（新しいデータから処理）
  for (let i = items.length - 1; i >= 0; i--) {
    const item = items[i];

    if (item['商品名'] === '送料') continue;

    const itemDate = Utilities.formatDate(new Date(item['発送日']), 'JST', 'yyyy/MM/dd');
    let dateKey = null;

    // どの日付に該当するか判定
    if (itemDate === dates.yesterdayStr) dateKey = 'yesterday';
    else if (itemDate === dates.todayStr) dateKey = 'today';
    else if (itemDate === dates.tomorrowStr) dateKey = 'tomorrow';
    else if (itemDate === dates.dayAfter2Str) dateKey = 'dayAfter2';
    else if (itemDate === dates.dayAfter3Str) dateKey = 'dayAfter3';

    if (dateKey) {
      processDateItem(item, dateKey, dataStructure);
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
function processDateItem(item, dateKey, dataStructure) {
  const lists = dataStructure.lists[dateKey];
  const invoiceTypes = dataStructure.invoiceTypes[dateKey];

  // キーの作成（発送先名とクール区分の組み合わせ）
  const listKey = `${item['発送先名']}_${item['クール区分']}_${item['受注ID']}`;

  // リストの処理
  if (lists.has(listKey)) {
    const record = lists.get(listKey);
    const existingProduct = record['商品'].find(p => p['商品名'] === item['商品名']);

    if (existingProduct) {
      existingProduct['受注数'] += Number(item['受注数']);
    } else {
      record['商品'].push({
        '商品名': item['商品名'],
        '受注数': Number(item['受注数']),
        'メモ': item['メモ'] || ""
      });
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

    // 新規レコード作成
    lists.set(listKey, {
      '顧客名': item['顧客名'],
      '発送先名': item['発送先名'],
      'クール区分': item['クール区分'],
      '商品': [{
        '商品名': item['商品名'],
        '受注数': Number(item['受注数']),
        'メモ': item['メモ'] || ""
      }]
    });
  }
}

/**
 * 集計済みデータ構造から最終的なレスポンス配列を構築
 * 
 * @param {Object} dateStrings - 表示用日付文字列
 * @param {Object} dataStructure - 集計済みデータ構造
 * @returns {Array} - クライアント側に返すデータの配列
 */
function buildResult(dateStrings, dataStructure) {
  const cnt = {
    '昨日': dataStructure.counts.yesterday,
    '今日': dataStructure.counts.today,
    '明日': dataStructure.counts.tomorrow,
    '明後日': dataStructure.counts.dayAfter2,
    '明々後日': dataStructure.counts.dayAfter3
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

  const convertListMap = (map) => {
    return Array.from(map.values()).map(record => ({
      '顧客名': record['顧客名'],
      '発送先名': record['発送先名'],
      'クール区分': record['クール区分'],
      '商品': record['商品']
    }));
  };

  return [
    dateStrings,
    cnt,
    convertMapToArray(dataStructure.invoiceTypes.yesterday),
    convertListMap(dataStructure.lists.yesterday),
    convertMapToArray(dataStructure.invoiceTypes.today),
    convertListMap(dataStructure.lists.today),
    convertMapToArray(dataStructure.invoiceTypes.tomorrow),
    convertListMap(dataStructure.lists.tomorrow),
    convertMapToArray(dataStructure.invoiceTypes.dayAfter2),
    convertListMap(dataStructure.lists.dayAfter2),
    convertMapToArray(dataStructure.invoiceTypes.dayAfter3),
    convertListMap(dataStructure.lists.dayAfter3)
  ];
}
