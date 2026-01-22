/**
 * freee納品書CSV生成機能
 *
 * freee会計ソフトへの納品書インポート用CSVファイルを生成します。
 *
 * 参照: https://support.freee.co.jp/hc/ja/articles/18895523266841
 *
 * 処理フロー:
 * 1. 指定期間・顧客の受注データを取得
 * 2. freee形式のCSV行を生成（本文行 + 明細行）
 * 3. UTF-8 BOM付きでCSVファイルを出力
 * 4. Googleドライブに保存
 *
 * @note freee側での事前準備が必要:
 *   - 取引先（顧客）の登録
 *   - 品目（商品）の登録（任意だが推奨）
 *   - 勘定科目の登録（例: 売上高）
 */

/**
 * freee納品書CSVを作成するメイン関数
 *
 * @param {string} datas - JSON形式の入力パラメータ
 * @param {string} datas.customerName - 対象顧客名（空の場合は全顧客）
 * @param {string} datas.targetDateFrom - 集計期間開始日（ISO形式）
 * @param {string} datas.targetDateTo - 集計期間終了日（ISO形式）
 * @returns {string} 生成されたCSVファイル名（JSON形式）
 *
 * @example
 * const input = JSON.stringify({
 *   customerName: "株式会社サンプル　田中太郎",
 *   targetDateFrom: "2024-01-01",
 *   targetDateTo: "2024-01-31"
 * });
 * const result = createFreeeDeliveryNote(input);
 */
function createFreeeDeliveryNote(datas) {
  const data = JSON.parse(datas);
  const customerName = data['customerName'];
  const targetFrom = Utilities.formatDate(new Date(data['targetDateFrom']), 'JST', 'yyyy/MM/dd');
  const targetTo = Utilities.formatDate(new Date(data['targetDateTo']), 'JST', 'yyyy/MM/dd');

  // 受注データを取得
  const items = getAllRecords('受注');

  // 期間と顧客でフィルタリング
  const targetLists = items.filter(function(target) {
    const deliveryDate = Utilities.formatDate(new Date(target['納品日']), 'JST', 'yyyy/MM/dd');
    const isInPeriod = deliveryDate >= targetFrom && deliveryDate <= targetTo;
    const isTargetCustomer = customerName === '' || target['顧客名'] === customerName;
    return isInPeriod && isTargetCustomer;
  });

  if (targetLists.length === 0) {
    return null;
  }

  // 受注を紐付けグループごとに整理
  const groupedOrders = groupOrdersByLinkedId(targetLists);

  // CSV生成
  const csvRows = generateFreeeCSV(groupedOrders);

  // CSVファイルをGoogleドライブに保存
  const fileName = saveFreeeCSVToDrive(csvRows, targetFrom, targetTo);

  return JSON.stringify([fileName]);
}

/**
 * 受注を紐付けグループごとに整理
 *
 * 複数日程の受注を「紐付け受注ID」でグループ化し、納品日順にソートします。
 * 単独受注（紐付けIDなし）は個別のグループとして扱います。
 *
 * @param {Array<Object>} orders - 受注データの配列
 * @returns {Array<Object>} グループ化された受注データ
 *   各要素は { parentId, customerName, orders: [...] } の形式
 */
function groupOrdersByLinkedId(orders) {
  const groups = {};

  orders.forEach(function(order) {
    const linkedId = order['紐付け受注ID'] || '';
    const orderId = order['受注ID'];
    const customerName = order['顧客名'];

    // グループキーの決定
    let groupKey;
    if (linkedId) {
      // 紐付けIDがある場合（子受注）→ 親IDをキーに使用
      groupKey = linkedId;
    } else {
      // 紐付けIDがない場合（親受注または単独受注）→ 自分のIDをキーに使用
      groupKey = orderId;
    }

    // グループ初期化
    if (!groups[groupKey]) {
      groups[groupKey] = {
        parentId: groupKey,
        customerName: customerName,
        orders: []
      };
    }

    groups[groupKey].orders.push(order);
  });

  // 各グループ内で納品日順にソート
  Object.keys(groups).forEach(function(groupKey) {
    groups[groupKey].orders.sort(function(a, b) {
      const dateA = new Date(a['納品日']);
      const dateB = new Date(b['納品日']);
      return dateA - dateB;
    });
  });

  // グループを配列に変換
  return Object.keys(groups).map(function(key) {
    return groups[key];
  });
}

/**
 * freee形式のCSV行を生成
 *
 * @param {Array<Object>} groupedOrders - 紐付けグループごとに整理された受注データ
 * @returns {Array<Array<string>>} CSV行の2次元配列
 */
function generateFreeeCSV(groupedOrders) {
  const rows = [];

  // ヘッダー行
  rows.push(getFreeeCSVHeader());

  // 顧客マスタと商品マスタを取得
  const customerItems = getAllRecords('顧客情報');
  const productItems = getAllRecords('商品');

  // 各グループの納品書を生成
  groupedOrders.forEach(function(group) {
    const customerName = group.customerName;
    const orders = group.orders;

    // 顧客情報を取得
    const customerInfo = getCustomerInfo(customerName, customerItems);

    // グループ内の最後の納品日を取得（発行日として使用）
    const lastDeliveryDate = Utilities.formatDate(
      new Date(orders[orders.length - 1]['納品日']),
      'JST',
      'yyyy-MM-dd'
    );

    // 親受注ID（グループのキー）
    const parentOrderId = group.parentId;

    // 最初の受注から社内メモと備考を取得
    const shippingMemo = orders[0]['送り状備考欄'] || '';
    const internalMemo = orders[0]['社内メモ'] || '';

    // 本文行を追加（発行日は最後の納品日、番号は親受注ID）
    rows.push(createFreeeMainRow(customerName, lastDeliveryDate, parentOrderId, shippingMemo, internalMemo, customerInfo));

    // 納品日ごとにグループ化
    const ordersByDate = {};
    orders.forEach(function(order) {
      const deliveryDate = Utilities.formatDate(new Date(order['納品日']), 'JST', 'yyyy-MM-dd');
      if (!ordersByDate[deliveryDate]) {
        ordersByDate[deliveryDate] = [];
      }
      ordersByDate[deliveryDate].push(order);
    });

    // 納品日順に処理
    const sortedDates = Object.keys(ordersByDate).sort();
    sortedDates.forEach(function(deliveryDate) {
      const dateOrders = ordersByDate[deliveryDate];

      // 納品書テキストがあればテキスト行を追加
      const deliveryNoteText = dateOrders[0]['納品書テキスト'] || '';
      if (deliveryNoteText) {
        rows.push(createFreeeTextRow(deliveryNoteText));
      }

      // この日付の明細行を追加
      dateOrders.forEach(function(order) {
        const productInfo = getProductInfo(order['商品名'], productItems);
        rows.push(createFreeeDetailRow(order, productInfo, deliveryDate));
      });
    });
  });

  return rows;
}

/**
 * freee CSVヘッダー行を生成
 *
 * @returns {Array<string>} ヘッダー列の配列
 */
function getFreeeCSVHeader() {
  return [
    '行形式',
    '発行日',
    '番号',
    '枝番',
    '件名',
    '発行元担当者氏名',
    '社内メモ',
    '備考',
    '消費税の表示方法',
    '消費税端数の計算方法',
    '取引先名称',
    '取引先宛名',
    '取引先敬称',
    '取引先郵便番号',
    '取引先都道府県',
    '取引先市区町村・番地',
    '取引先建物名・部屋番号',
    '取引先部署',
    '取引先担当者氏名',
    '行の種類',
    '摘要',
    '単価',
    '数量',
    '単位',
    '税率',
    '源泉徴収',
    '発生日',
    '勘定科目',
    '税区分',
    '部門',
    '品目',
    'メモ',
    'セグメント1',
    'セグメント2',
    'セグメント3'
  ];
}

/**
 * freee 本文行を生成
 *
 * @param {string} customerName - 顧客名
 * @param {string} deliveryDate - 納品日（yyyy-MM-dd形式）
 * @param {string} orderId - 受注ID
 * @param {string} shippingMemo - 送り状備考欄（備考列用）
 * @param {string} internalMemo - 社内メモ（社内メモ列用）
 * @param {Object} customerInfo - 顧客情報
 * @returns {Array<string>} 本文行の列データ
 */
function createFreeeMainRow(customerName, deliveryDate, orderId, shippingMemo, internalMemo, customerInfo) {
  const row = new Array(35).fill(''); // 35列全て初期化

  row[0] = '本文';                    // 行形式
  row[1] = deliveryDate;              // 発行日
  row[2] = orderId;                   // 番号（受注ID）
  row[5] = '大森';                    // 発行元担当者氏名（固定値）
  row[6] = internalMemo;              // 社内メモ
  row[7] = shippingMemo;              // 備考（送り状備考欄）
  row[8] = customerInfo['消費税の表示方法'] || ''; // 消費税の表示方法（内税 or 外税）
  row[10] = customerName;             // 取引先名称
  row[12] = customerInfo['敬称'] || ''; // 取引先敬称（様 or 御中）
  row[13] = customerInfo['郵便番号'] || ''; // 取引先郵便番号
  row[14] = customerInfo['都道府県'] || ''; // 取引先都道府県
  row[15] = customerInfo['市区町村・番地'] || customerInfo['住所１'] || ''; // 取引先市区町村・番地
  row[16] = customerInfo['建物名・部屋番号'] || customerInfo['住所２'] || ''; // 取引先建物名・部屋番号
  row[18] = customerInfo['氏名'] || customerInfo['担当者氏名'] || ''; // 取引先担当者氏名

  return row;
}

/**
 * freee テキスト行を生成
 *
 * 納品書に見出しやグループ分け用のテキストを挿入するための行です。
 * 「行の種類」を「テキスト」にして、「摘要」列のみにテキストを設定します。
 *
 * @param {string} text - 表示するテキスト（例: 「月曜コース分の商品」）
 * @returns {Array<string>} テキスト行の列データ
 */
function createFreeeTextRow(text) {
  const row = new Array(35).fill(''); // 35列全て初期化

  row[0] = '明細';                    // 行形式（テキスト行も「明細」）
  row[19] = 'テキスト';               // 行の種類（重要: 「通常」ではなく「テキスト」）
  row[20] = text;                     // 摘要（表示するテキスト）
  // その他の列は空欄（単価、数量、税率などは不要）

  return row;
}

/**
 * freee 明細行を生成
 *
 * @param {Object} order - 受注データ
 * @param {Object} productInfo - 商品マスタ情報
 * @param {string} deliveryDate - 納品日（yyyy-MM-dd形式）
 * @returns {Array<string>} 明細行の列データ
 */
function createFreeeDetailRow(order, productInfo, deliveryDate) {
  const row = new Array(35).fill(''); // 35列全て初期化

  row[0] = '明細';                    // 行形式
  row[19] = '通常';                   // 行の種類
  row[20] = order['商品名'];          // 摘要
  row[21] = order['販売価格'];        // 単価
  row[22] = order['受注数'];          // 数量
  row[23] = productInfo['単位'] || ''; // 単位（kg, 個, 箱など）
  row[24] = convertTaxRateForFreee(productInfo['税率']); // 税率
  row[26] = deliveryDate;             // 発生日
  row[27] = productInfo['勘定科目'] || '売上高'; // 勘定科目

  return row;
}

/**
 * 税率を数値からfreee形式に変換
 *
 * 受注システムでは「8」「10」という数値で管理されているが、
 * freeeでは「8% (軽減税率)」「10%」という文字列形式が必要。
 *
 * @param {number|string} taxRate - 商品マスタの税率（8 or 10）
 * @returns {string} freee形式の税率（"8% (軽減税率)" or "10%"）
 *
 * @note スペースは半角1個、カッコは半角
 */
function convertTaxRateForFreee(taxRate) {
  const rate = Number(taxRate);

  if (rate > 8) {
    return '10%';
  } else {
    return '8% (軽減税率)'; // 半角スペース1個 + 半角カッコ
  }
}

/**
 * 顧客情報を取得
 *
 * @param {string} customerName - 顧客名（"会社名　氏名" または "氏名" 形式）
 * @param {Array<Object>} customerItems - 顧客マスタ
 * @returns {Object} 顧客情報（見つからない場合は空オブジェクト）
 */
function getCustomerInfo(customerName, customerItems) {
  const nameParts = customerName.split('　');

  for (let i = 0; i < customerItems.length; i++) {
    const customer = customerItems[i];

    if (nameParts.length > 1) {
      // "会社名　氏名" 形式
      if (customer['会社名'] === nameParts[0] && customer['氏名'] === nameParts[1]) {
        return customer;
      }
    } else {
      // "氏名" または "会社名" 形式
      if (customer['氏名'] === nameParts[0] || customer['会社名'] === nameParts[0]) {
        return customer;
      }
    }
  }

  return {}; // 見つからない場合
}

/**
 * 商品情報を取得
 *
 * @param {string} productName - 商品名
 * @param {Array<Object>} productItems - 商品マスタ
 * @returns {Object} 商品情報（見つからない場合はデフォルト値）
 */
function getProductInfo(productName, productItems) {
  for (let i = 0; i < productItems.length; i++) {
    if (productItems[i]['商品名'] === productName) {
      return productItems[i];
    }
  }

  // 商品が見つからない場合のデフォルト値
  return {
    '税率': 8,
    '勘定科目': '売上高'
  };
}

/**
 * CSVをGoogleドライブに保存
 *
 * @param {Array<Array<string>>} csvRows - CSV行データ
 * @param {string} targetFrom - 期間開始日
 * @param {string} targetTo - 期間終了日
 * @returns {string} 保存したファイル名
 */
function saveFreeeCSVToDrive(csvRows, targetFrom, targetTo) {
  // UTF-8 BOM付きCSVを生成
  const csvContent = generateCSVWithBOM(csvRows);

  // ファイル名を生成（freee納品書CSV_YYYYMMDD-YYYYMMDD）
  const fromDate = targetFrom.replace(/\//g, '');
  const toDate = targetTo.replace(/\//g, '');
  const fileName = 'freee納品書CSV_' + fromDate + '-' + toDate;

  // Googleドライブに保存
  const folderId = getFreeeCSVFolderId();
  const folder = DriveApp.getFolderById(folderId);

  const blob = Utilities.newBlob(csvContent, 'text/csv', fileName + '.csv');
  folder.createFile(blob);

  return fileName;
}

/**
 * UTF-8 BOM付きCSV文字列を生成
 *
 * freeeのインポート仕様では「UTF-8 with BOM」が必須。
 * BOM（Byte Order Mark）は \uFEFF で表現される。
 *
 * @param {Array<Array<string>>} rows - CSV行データ
 * @returns {string} BOM付きCSV文字列
 */
function generateCSVWithBOM(rows) {
  const BOM = '\uFEFF';

  // CSVエスケープ処理
  const csvLines = rows.map(function(row) {
    return row.map(function(cell) {
      const cellStr = String(cell);
      // カンマ、改行、ダブルクォートを含む場合はエスケープ
      if (cellStr.includes(',') || cellStr.includes('\n') || cellStr.includes('"')) {
        return '"' + cellStr.replace(/"/g, '""') + '"';
      }
      return cellStr;
    }).join(',');
  });

  return BOM + csvLines.join('\n');
}

/**
 * freee CSV出力フォルダIDを取得
 *
 * スクリプトプロパティから取得。未設定の場合は請求書と同じフォルダを使用。
 *
 * @returns {string} フォルダID
 */
function getFreeeCSVFolderId() {
  const props = PropertiesService.getScriptProperties();
  let folderId = props.getProperty('FREEE_CSV_FOLDER_ID');

  // 未設定の場合は請求書フォルダを使用
  if (!folderId) {
    folderId = getBillPdfFolderId();
  }

  return folderId;
}
