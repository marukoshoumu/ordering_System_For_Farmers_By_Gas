/**
 * 共通ライブラリ - sharedLib.js
 *
 * メインシステム(src/)とAllUserシステム(AllUser/src/)で共有する関数群
 *
 * 【使用方法】
 * - メインシステム: GASが自動結合するため、そのまま使用可能
 * - AllUserシステム: デプロイ前に sync-shared.sh でコピー
 *
 * 【将来的なGASライブラリ化に向けて】
 * このファイルの関数は外部依存を最小限にし、
 * getSpreadsheet()のような環境依存関数は呼び出し元で注入する設計
 */

// ============================================
// 日付ユーティリティ
// ============================================

/**
 * ダッシュボード表示に必要な各日程（昨日、今日、明日、明後日、明々後日）を計算
 *
 * @returns {Object} - 日付オブジェクトと yyyy/MM/dd 形式の文字列を含むオブジェクト
 */
function calculateDatesShared() {
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
 * @param {Object} dates - calculateDatesShared() で作成された日付オブジェクト
 * @returns {Object} - 「MM/dd (曜日)」形式の文字列
 */
function createDateStringsShared(dates) {
  const weekdays = ['日', '月', '火', '水', '木', '金', '土'];

  return {
    '昨日': formatDateWithWeekdayShared(dates.yesterday, weekdays),
    '今日': formatDateWithWeekdayShared(dates.today, weekdays),
    '明日': formatDateWithWeekdayShared(dates.tomorrow, weekdays),
    '明後日': formatDateWithWeekdayShared(dates.dayAfter2, weekdays),
    '明々後日': formatDateWithWeekdayShared(dates.dayAfter3, weekdays)
  };
}

/**
 * 日付を MM/dd (曜日) 形式にフォーマット
 *
 * @param {Date} date - 日付オブジェクト
 * @param {Array<string>} weekdays - 曜日の配列（['日', '月', ...]）
 * @returns {string} - フォーマット済み文字列
 */
function formatDateWithWeekdayShared(date, weekdays) {
  const formattedDate = Utilities.formatDate(date, 'JST', 'MM/dd');
  const weekday = weekdays[date.getDay()];
  return `${formattedDate} (${weekday})`;
}

/**
 * 間隔オブジェクトを人間が読める形式に変換する（定期便・受注で共有）
 * @param {Object|number|string} intervalObj - 間隔オブジェクト、数値、または文字列
 * @returns {string} 表示用文字列
 */
function formatIntervalForDisplay(intervalObj) {
  if (!intervalObj) return '1ヶ月ごと';

  // 旧形式（数値のみ、または JSON でない文字列）
  if (typeof intervalObj === 'number') {
    return intervalObj + 'ヶ月ごと';
  }
  if (typeof intervalObj === 'string' && !intervalObj.startsWith('{')) {
    return intervalObj + 'ヶ月ごと';
  }
  if (typeof intervalObj === 'string') {
    try {
      intervalObj = JSON.parse(intervalObj);
    } catch (e) {
      return intervalObj + 'ヶ月ごと';
    }
  }

  var type = intervalObj.type;
  var value = intervalObj.value;
  var weekday = intervalObj.weekday;
  var dayNames = ['', '月', '火', '水', '木', '金', '土', '日'];

  switch (type) {
    case 'weekly':
      return '毎週' + (dayNames[value] || '') + '曜日';
    case 'nweek':
    case 'biweekly':
    case 'triweekly':
      var n = Number(value) || 2;
      return n + '週ごと' + (dayNames[weekday] || '') + '曜日';
    case 'monthly':
      if (value === 'first') return '毎月初日';
      if (value === 'last') return '毎月末日';
      return '毎月' + value + '日';
    case 'nmonth':
    case '2month':
    case '3month':
      return (Number(value) || 1) + 'ヶ月ごと';
    default:
      return (Number(value) || 1) + 'ヶ月ごと';
  }
}

// ============================================
// データ構造初期化
// ============================================

/**
 * 集計用データ構造を初期化
 *
 * @returns {Object} - 空の Map とカウントを含む初期構造
 */
function initializeDataStructureShared() {
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

// ============================================
// 受注データ処理
// ============================================

/**
 * 指定された日付範囲の受注レコードを取得
 *
 * @param {Spreadsheet} ss - スプレッドシートオブジェクト
 * @param {Date} startDate - 開始日
 * @param {Date} endDate - 終了日
 * @returns {Array<Object>} - ラベルをキーとしたレコードの配列
 */
function getOrdersByDateRangeShared(ss, startDate, endDate) {
  const sheet = ss.getSheetByName('受注');
  const values = sheet.getDataRange().getValues();
  const labels = values.shift();

  // 列インデックスを取得
  const dateColIndex = labels.indexOf('発送日');
  const statusColIndex = labels.indexOf('ステータス');

  const startTime = startDate.getTime();
  const endTime = endDate.getTime() + 24 * 60 * 60 * 1000; // 終了日の翌日0時

  const records = [];
  for (const value of values) {
    const shippingDate = new Date(value[dateColIndex]);
    const shippingTime = shippingDate.getTime();

    // 日付範囲内のみ追加
    if (shippingTime >= startTime && shippingTime < endTime) {
      // フィルタリング: キャンセルを除外 (出荷済は含める)
      const isCancelled = statusColIndex !== -1 && (value[statusColIndex] === 'キャンセル' || value[statusColIndex] === 'cancelled');

      if (isCancelled) continue;

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
 * 取得した受注アイテムを日付ごとに分類して処理
 *
 * @param {Array<Object>} items - 受注レコードの配列
 * @param {Object} dates - 日程オブジェクト
 * @param {Object} dataStructure - 集計用データ構造
 * @param {Object} unitMap - 商品単位マップ
 */
function processItemsShared(items, dates, dataStructure, unitMap) {
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
      processDateItemShared(item, dateKey, dataStructure, unitMap);
    }
  }
}

/**
 * 特定の日付における受注アイテムの集計処理
 *
 * @param {Object} item - 受注レコード単体
 * @param {string} dateKey - 日付キー（yesterday/today/...)
 * @param {Object} dataStructure - 集計用データ構造
 * @param {Object} unitMap - 商品単位マップ
 */
function processDateItemShared(item, dateKey, dataStructure, unitMap) {
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
      'orderId': item['受注ID'],
      '顧客名': item['顧客名'],
      '発送先名': item['発送先名'],
      'クール区分': item['クール区分'],
      'deliveryMethod': deliveryMethod,
      'trackingNumber': item['追跡番号'] || '',
      'status': item['ステータス'] || '',
      'shipped': item['出荷済'] === '○' || item['出荷済'] === true,
      '商品': [{
        '商品名': item['商品名'],
        '受注数': Number(item['受注数']),
        '単位': unitMap[item['商品名']] || '枚',
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
function buildResultShared(dateStrings, dataStructure) {
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
      'orderId': record['orderId'],
      '顧客名': record['顧客名'],
      '発送先名': record['発送先名'],
      'クール区分': record['クール区分'],
      'deliveryMethod': record['deliveryMethod'],
      'trackingNumber': record['trackingNumber'],
      'status': record['status'],
      'shipped': record['shipped'],
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

// ============================================
// 出荷ステータス更新
// ============================================

/**
 * 受注の出荷済ステータスを更新
 *
 * @param {Spreadsheet} ss - スプレッドシートオブジェクト
 * @param {string} orderId - 受注ID
 * @param {string} shippedValue - 出荷済の値（'○' or ''）
 * @returns {Object} - 更新結果
 */
function updateOrderShippedStatusShared(ss, orderId, shippedValue) {
  const sheet = ss.getSheetByName('受注');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  // 列インデックスを取得
  const orderIdCol = headers.indexOf('受注ID');
  const shippedCol = headers.indexOf('出荷済');

  if (orderIdCol === -1 || shippedCol === -1) {
    throw new Error('必要な列が見つかりません');
  }

  // 該当する受注IDの全行を更新
  let updatedCount = 0;
  const numRows = data.length - 1; // ヘッダー行を除く
  const startRow = 2; // データ行の開始行（1-indexed）
  
  // 出荷済み列の値を2D配列として取得
  const updatedColumnArray = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][orderIdCol] === orderId) {
      updatedColumnArray.push([shippedValue]);
      updatedCount++;
    } else {
      updatedColumnArray.push([data[i][shippedCol]]);
    }
  }
  
  // 一度にすべての変更を書き戻す
  if (updatedCount > 0) {
    sheet.getRange(startRow, shippedCol + 1, numRows, 1).setValues(updatedColumnArray);
  }

  if (updatedCount === 0) {
    throw new Error('該当する受注IDが見つかりません: ' + orderId);
  }

  return { success: true, updatedRows: updatedCount };
}

/**
 * 追跡番号を更新し、ステータスを「出荷済」にする
 *
 * @param {Spreadsheet} ss - スプレッドシートオブジェクト
 * @param {string} orderId - 受注ID
 * @param {string} trackingNumber - 追跡番号（送り状番号）
 * @returns {Object} 実行結果
 */
function updateOrderTrackingStatusShared(ss, orderId, trackingNumber) {
  try {
    const sheet = ss.getSheetByName('受注');
    const data = sheet.getDataRange().getValues();
    const headers = data[0];

    const orderIdCol = headers.indexOf('受注ID');
    const trackingCol = headers.indexOf('追跡番号');
    const shippedCol = headers.indexOf('出荷済');

    if (orderIdCol === -1 || trackingCol === -1 || shippedCol === -1) {
      return { success: false, message: '必要な列（受注ID, 追跡番号, 出荷済）が見つかりません。' };
    }

    let updatedCount = 0;
    for (let i = 1; i < data.length; i++) {
      if (data[i][orderIdCol] === orderId) {
        // 追跡番号をセット
        sheet.getRange(i + 1, trackingCol + 1).setValue(trackingNumber);
        // 出荷済を「○」にセット
        sheet.getRange(i + 1, shippedCol + 1).setValue('○');
        updatedCount++;
      }
    }

    if (updatedCount === 0) {
      return { success: false, message: '該当する受注IDが見つかりませんでした。' };
    }

    return {
      success: true,
      message: `受注ID: ${orderId} に追跡番号 ${trackingNumber} を紐付け、出荷済に更新しました。`
    };
  } catch (error) {
    return { success: false, message: '更新エラー: ' + error.toString() };
  }
}

// ============================================
// OCR / 追跡番号認識
// ============================================

/**
 * Vision APIキーを取得
 *
 * @returns {string|null} - APIキー
 */
function getVisionApiKeyShared() {
  const props = PropertiesService.getScriptProperties();
  return props.getProperty('VISION_API_KEY') || props.getProperty('GEMINI_API_KEY');
}

/**
 * Cloud Vision API を使用して画像から追跡番号を抽出する
 *
 * @param {string} base64Data - 画像のBase64文字列
 * @returns {Object} 抽出結果 {success: boolean, code: string, message: string}
 */
function recognizeTrackingNumberShared(base64Data) {
  const apiKey = getVisionApiKeyShared();
  if (!apiKey) {
    return { success: false, message: 'Google Cloud Vision API キーが設定されていません。' };
  }

  const url = `https://vision.googleapis.com/v1/images:annotate?key=${apiKey}`;
  const requestBody = {
    requests: [
      {
        image: { content: base64Data },
        features: [{ type: 'TEXT_DETECTION' }]
      }
    ]
  };

  try {
    const response = UrlFetchApp.fetch(url, {
      method: 'POST',
      contentType: 'application/json',
      payload: JSON.stringify(requestBody),
      muteHttpExceptions: true
    });

    const result = JSON.parse(response.getContentText());

    if (!result.responses || !Array.isArray(result.responses) || result.responses.length === 0) {
      const apiError = result.error ? (result.error.message || JSON.stringify(result.error)) : '';
      if (apiError) Logger.log('Vision API error: ' + apiError);
      return { success: false, message: '文字が検出されませんでした。' };
    }

    const firstResponse = result.responses[0];
    if (!firstResponse) {
      return { success: false, message: '文字が検出されませんでした。' };
    }

    if (firstResponse.error) {
      const errMsg = firstResponse.error.message || JSON.stringify(firstResponse.error);
      Logger.log('Vision API response error: ' + errMsg);
      return { success: false, message: '文字が検出されませんでした。' };
    }

    const annotations = firstResponse.textAnnotations;

    if (!annotations || annotations.length === 0) {
      return { success: false, message: '文字が検出されませんでした。' };
    }

    // 全体の文字列を取得
    const fullText = annotations[0].description;

    // 配送伝票番号特有のパターンを探す
    // 優先順位: 1. ハイフンあり(12桁) 2. ハイフンあり(10桁) 3. 12桁 4. 10桁
    const patterns = [
      /\d{4}-\d{4}-\d{4}/, // ヤマト/佐川等 (12桁)
      /\d{3}-\d{3}-\d{4}/, // 西濃/郵便等 (10桁/11桁)
      /\d{4}-\d{4}-\d{2}/, // 西濃2
      /\d{12}/,            // 12桁
      /\d{11}/,            // 11桁
      /\d{10}/             // 10桁
    ];

    for (const pattern of patterns) {
      const matches = fullText.match(new RegExp(pattern, 'g'));
      if (matches) {
        for (const m of matches) {
          const cleaned = m.replace(/-/g, '');
          // 日本の電話番号（0から始まる10/11桁）は、他に候補があればスキップ
          if ((cleaned.length === 10 || cleaned.length === 11) && cleaned.startsWith('0')) {
            continue;
          }
          // 最初に見つかった有力な候補を返す
          return { success: true, code: cleaned };
        }
      }
    }

    return { success: false, message: '認識されたテキストに伝票番号が見つかりませんでした。', rawText: fullText };

  } catch (e) {
    return { success: false, message: 'API呼び出しエラー: ' + e.toString() };
  }
}

// ============================================
// 商品マスタ
// ============================================

/**
 * 商品マスタから単位情報を取得
 *
 * @param {Spreadsheet} ss - スプレッドシートオブジェクト
 * @returns {Object} - 商品名をキー、単位を値とするオブジェクト
 */
function getProductUnitMapShared(ss) {
  const unitMap = {};
  try {
    const sheet = ss.getSheetByName('商品');
    if (!sheet) return unitMap;

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const nameCol = headers.indexOf('商品名');
    const unitCol = headers.indexOf('単位');

    if (nameCol === -1) return unitMap;

    for (let i = 1; i < data.length; i++) {
      const name = data[i][nameCol];
      const unit = unitCol !== -1 ? data[i][unitCol] : '枚';
      if (name) {
        unitMap[name] = unit || '枚';
      }
    }
  } catch (e) {
    Logger.log('商品単位取得エラー: ' + e.toString());
  }
  return unitMap;
}
