// ============================================
// 修正版: getOrderByOrderId関数
// 
// 修正内容:
// - 納品方法（ヤマト/佐川等）に応じた配達時間帯・送り状種別・クール区分・荷扱いの逆引き
// - マスタの「納品方法」列でフィルタリングして正しい種別値を取得
// ============================================

/**
 * 受注IDから受注データを取得（受注修正画面用）
 *
 * 指定された受注IDの全データを取得し、受注修正画面のフォームに展開可能な形式で返却します。
 * 納品方法（ヤマト/佐川）に応じて、配達時間帯・送り状種別・クール区分・荷扱いを
 * 表示名から種別値に逆変換します（受注シートには表示名が保存されているため）。
 *
 * 重要な処理:
 * - 納品方法でマスタデータをフィルタリングして逆引きマップを作成
 * - 受注シートの表示名（「午前中」「発払い」等）→ 種別値（「0812:午前中」「0:発払い」）に変換
 * - 商品情報は複数行対応（同一受注IDの全行を取得）
 *
 * 処理フロー:
 * 1. アクティブスプレッドシート取得
 * 2. '受注'シートから全データ取得
 * 3. 受注IDが一致する行を全て検索
 * 4. 1行目から納品方法を取得
 * 5. getMasterDataCached() でマスタデータ取得
 * 6. 納品方法でフィルタリングして逆引きマップを作成:
 *    - 送り状種別マップ（種別 → 種別値）
 *    - クール区分マップ（種別 → 種別値）
 *    - 荷扱いマップ（種別 → 種別値）
 *    - 配達時間帯マップ（時間指定 → 時間指定値）
 * 7. 基本情報をオブジェクトに格納（発送先、顧客、発送元、日程、受付情報）
 * 8. 表示名を種別値に変換（配達時間帯、送り状種別、クール区分、荷扱い）
 * 9. 商品情報配列を作成（同一受注IDの全行から抽出）
 * 10. 受注データオブジェクトを返却
 *
 * @param {string} orderId - 受注ID（例: "ORD001"）
 * @returns {Object|null} 受注データオブジェクト、該当なしの場合はnull
 *   {
 *     orderId: string, orderDate: Date,
 *     shippingToName: string, shippingToZipcode: string, shippingToAddress: string, shippingToTel: string,
 *     customerName: string, customerZipcode: string, customerAddress: string, customerTel: string,
 *     shippingFromName: string, shippingFromZipcode: string, shippingFromAddress: string, shippingFromTel: string,
 *     shippingDate: string, deliveryDate: string,
 *     receiptWay: string, recipient: string, deliveryMethod: string,
 *     deliveryTime: string,  // 種別値形式（例: "0812:午前中"）
 *     checklist: { deliverySlip: boolean, bill: boolean, receipt: boolean, pamphlet: boolean, recipe: boolean },
 *     otherAttach: string, sendProduct: string,
 *     invoiceType: string,   // 種別値形式（例: "0:発払い"）
 *     coolCls: string,       // 種別値形式（例: "0:なし"）
 *     cargo1: string,        // 種別値形式（例: "0:なし"）
 *     cargo2: string,        // 種別値形式（例: "0:なし"）
 *     cashOnDelivery: number, cashOnDeliTax: number, copiePrint: number,
 *     internalMemo: string, csvmemo: string, deliveryMemo: string, memo: string,
 *     items: [
 *       { bunrui: string, product: string, quantity: number, price: number },
 *       ...
 *     ]
 *   }
 *
 * 逆変換の仕組み:
 * - 受注シート保存時: 「0812:午前中」形式（種別値形式）
 * - 受注シート表示時: 「午前中」のみ表示（表示名）
 * - 修正画面展開時: 「0812:午前中」形式に逆変換（本関数で実施）
 *
 * @see formatDateForInput() - 日付をinput[type="date"]用に変換
 * @see getMasterDataCached() - マスタデータ取得（キャッシュ付き）
 * @see deleteOrderByOrderId() - 受注削除
 * @see getOrderDataForInherit() - 受注引継ぎ用データ取得（customerCode.js）
 *
 * 呼び出し元: orderList.html の編集ボタン、doPost() の editOrderId パラメータ処理
 */
function getOrderByOrderId(orderId) {
  if (!orderId) return null;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('受注');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  // ヘッダーからインデックスを取得
  const getColIndex = (name) => headers.indexOf(name);

  // ============================================
  // 紐付け受注対応: 関連する全受注を取得
  // ============================================

  // 指定された受注を取得
  const targetRows = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][getColIndex('受注ID')] === orderId) {
      targetRows.push(data[i]);
    }
  }

  if (targetRows.length === 0) return null;

  const targetRow = targetRows[0];

  // 紐付けIDを確認して親IDを決定
  const linkedId = targetRow[getColIndex('紐付け受注ID')] || '';
  let parentId;

  if (linkedId) {
    // 紐付けがある（子受注） → 親IDを使用
    parentId = linkedId;
  } else {
    // 紐付けがない（親受注または単独） → 自分のIDを使用
    parentId = orderId;
  }

  // 親ID + 紐付きIDが親IDのものをすべて取得してグループ化
  const allGroupOrders = {};  // 受注ID → 行配列のマップ

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const rowOrderId = row[getColIndex('受注ID')];
    const rowLinkedId = row[getColIndex('紐付け受注ID')] || '';

    // 親自身、または親に紐付いている子
    if (rowOrderId === parentId || rowLinkedId === parentId) {
      if (!allGroupOrders[rowOrderId]) {
        allGroupOrders[rowOrderId] = [];
      }
      allGroupOrders[rowOrderId].push(row);
    }
  }

  // 受注IDごとの配列に変換し、発送日でソート
  const orderGroups = Object.keys(allGroupOrders).map(id => {
    return {
      orderId: id,
      rows: allGroupOrders[id]
    };
  }).sort((a, b) => {
    const dateA = new Date(a.rows[0][getColIndex('発送日')]);
    const dateB = new Date(b.rows[0][getColIndex('発送日')]);
    return dateA - dateB;
  });

  // 親グループを優先的に選択（親IDと一致するグループを探す）
  const parentGroup = orderGroups.find(g => g.orderId === parentId);
  const selectedGroup = parentGroup || orderGroups[0];

  // 選択されたグループの最初の行を基本情報として使用
  const matchingRows = selectedGroup.rows;
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
    // 親IDを正規のIDとして使用（子受注から開いた場合も親グループ全体を編集/削除対象とする）
    orderId: parentId,
    originalOrderId: orderId,  // 元々リクエストされたID（デバッグ用）
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
    trackingNumber: firstRow[getColIndex('追跡番号')] || '',

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
    internalMemo: firstRow[getColIndex('社内メモ')] || '',
    csvmemo: firstRow[getColIndex('送り状備考欄')] || '',
    deliveryMemo: firstRow[getColIndex('納品書備考欄')] || '',
    memo: firstRow[getColIndex('メモ')] || '',

    // 商品情報（複数行）
    items: []
  };

  // 商品情報を取得（後方互換性のため、最初の日程の商品をitemsに格納）
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

  // ============================================
  // 複数日程対応: dates配列を追加
  // ============================================
  result.dates = orderGroups.map(group => {
    const groupFirstRow = group.rows[0];
    const dateItems = [];

    // この日程の商品情報を取得
    group.rows.forEach(row => {
      const bunrui = row[getColIndex('商品分類')];
      const product = row[getColIndex('商品名')];
      const quantity = row[getColIndex('受注数')];
      const price = row[getColIndex('販売価格')];

      if (product) {
        dateItems.push({
          bunrui: bunrui || '',
          product: product || '',
          quantity: quantity || 0,
          price: price || 0
        });
      }
    });

    return {
      orderId: group.orderId,
      shippingDate: formatDateForInput(groupFirstRow[getColIndex('発送日')]),
      deliveryDate: formatDateForInput(groupFirstRow[getColIndex('納品日')]),
      deliveryNoteText: groupFirstRow[getColIndex('納品書テキスト')] || '',
      items: dateItems
    };
  });

  return result;
}

/**
 * 日付をinput[type="date"]用のフォーマットに変換（yyyy-MM-dd形式）
 *
 * Date オブジェクトまたは日付文字列を HTML の input[type="date"] で使用可能な
 * yyyy-MM-dd 形式に変換します。受注修正画面の日付フィールドに値を設定する際に使用されます。
 *
 * 処理フロー:
 * 1. dateValue が空の場合: 空文字列返却
 * 2. dateValue が Date オブジェクトの場合: そのまま使用
 * 3. dateValue が文字列の場合:
 *    - yyyy/MM/dd 形式を yyyy-MM-dd に置換
 *    - new Date() で Date オブジェクトに変換
 * 4. 無効な日付の場合（isNaN(date.getTime())）: 空文字列返却
 * 5. Utilities.formatDate() で 'yyyy-MM-dd' 形式に変換
 * 6. エラー発生時: Logger.log() でログ出力、空文字列返却
 *
 * @param {Date|string|*} dateValue - 日付（Date オブジェクトまたは文字列）
 * @returns {string} yyyy-MM-dd 形式の日付文字列、無効な場合は空文字列
 *
 * 変換例:
 * - new Date('2024-05-15') → '2024-05-15'
 * - '2024/05/15' → '2024-05-15'
 * - '2024-05-15' → '2024-05-15'
 * - null → ''
 * - '無効な日付' → ''
 *
 * @see getOrderByOrderId() - 受注データ取得時に使用
 * @see Utilities.formatDate() - Google Apps Script の日付フォーマット関数
 *
 * 呼び出し元: getOrderByOrderId() の shippingDate, deliveryDate 設定
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
 * 受注IDで受注データを削除（受注シートから該当レコード全削除）
 *
 * 指定された受注IDの全レコードを「受注」シートから削除します。
 * 同一受注IDで複数行存在する場合（複数商品の場合）、全ての行を削除します。
 * 後ろから削除することで行番号のズレを防ぎます。
 *
 * 処理フロー:
 * 1. orderId が空の場合: 0 を返却
 * 2. アクティブスプレッドシート取得
 * 3. '受注'シートのデータ全取得
 * 4. ヘッダ行から '受注ID' 列のインデックス取得
 * 5. '受注ID' 列が見つからない場合: エラーログ、0 を返却
 * 6. データを後ろから走査して該当行を特定
 * 7. 削除対象行を配列に格納（シートの行番号: 1-indexed）
 * 8. 後ろから順に deleteRow() で行削除
 * 9. Logger.log() で削除件数をログ出力
 * 10. 削除した行数を返却
 *
 * @param {string} orderId - 受注ID（例: "ORD001"）
 * @returns {number} 削除した行数（0=該当なし）
 *
 * 使用例:
 * const deletedCount = deleteOrderByOrderId("ORD001");
 * // 返却: 3 （3商品の受注だった場合）
 *
 * 注意:
 * - この関数は受注シートのみを削除します
 * - ヤマトCSV・佐川CSVは deleteYamatoCSVByOrderId(), deleteSagawaCSVByOrderId() で別途削除
 * - 削除は不可逆的なため、事前に確認が必要
 *
 * @see deleteYamatoCSVByOrderId() - ヤマトCSVデータ削除
 * @see deleteSagawaCSVByOrderId() - 佐川CSVデータ削除
 * @see getOrderByOrderId() - 受注データ取得
 *
 * 呼び出し元: orderList.html の削除ボタン、受注修正時の再登録前削除
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
 * 受注IDでヤマトCSVデータを削除（ヤマトCSVシートから該当レコード全削除）
 *
 * 指定された受注IDの全レコードを「ヤマトCSV」シートから削除します。
 * 「お客様管理番号」列に受注IDが保存されているため、これをキーに検索・削除します。
 * 受注削除時または受注修正時に呼び出されます。
 *
 * 処理フロー:
 * 1. orderId が空の場合: 0 を返却
 * 2. アクティブスプレッドシート取得
 * 3. 'ヤマトCSV'シート取得
 * 4. シートが存在しない場合: エラーログ、0 を返却
 * 5. シートのデータ全取得
 * 6. ヘッダ行から 'お客様管理番号' 列のインデックス取得
 * 7. 'お客様管理番号' 列が見つからない場合: エラーログ、0 を返却
 * 8. データを後ろから走査して該当行を特定
 * 9. 削除対象行を配列に格納（シートの行番号: 1-indexed）
 * 10. 後ろから順に deleteRow() で行削除
 * 11. Logger.log() で削除件数をログ出力
 * 12. 削除した行数を返却
 *
 * @param {string} orderId - 受注ID（例: "ORD001"）
 * @returns {number} 削除した行数（0=該当なし、シートなし）
 *
 * 使用例:
 * const deletedCount = deleteYamatoCSVByOrderId("ORD001");
 * // 返却: 3 （3商品分のヤマトCSV行が削除された）
 *
 * ヤマトCSVシート構造:
 * - 「お客様管理番号」列に受注IDが保存されている
 * - 1受注IDで複数行存在する場合がある（複数商品の場合）
 * - 送り状作成時に createCsv() で生成される
 *
 * @see deleteOrderByOrderId() - 受注データ削除
 * @see deleteSagawaCSVByOrderId() - 佐川CSVデータ削除
 * @see createCsv() - ヤマトCSV生成（createCSVFile.js）
 *
 * 呼び出し元: 受注削除処理、受注修正処理
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
 * 受注IDで佐川CSVデータを削除（佐川CSVシートから該当レコード全削除）
 *
 * 指定された受注IDの全レコードを「佐川CSV」シートから削除します。
 * 「お客様管理番号」列に受注IDが保存されているため、これをキーに検索・削除します。
 * 受注削除時または受注修正時に呼び出されます。
 *
 * 処理フロー:
 * 1. orderId が空の場合: 0 を返却
 * 2. アクティブスプレッドシート取得
 * 3. '佐川CSV'シート取得
 * 4. シートが存在しない場合: エラーログ、0 を返却
 * 5. シートのデータ全取得
 * 6. ヘッダ行から 'お客様管理番号' 列のインデックス取得
 * 7. 'お客様管理番号' 列が見つからない場合: エラーログ、0 を返却
 * 8. データを後ろから走査して該当行を特定
 * 9. 削除対象行を配列に格納（シートの行番号: 1-indexed）
 * 10. 後ろから順に deleteRow() で行削除
 * 11. Logger.log() で削除件数をログ出力
 * 12. 削除した行数を返却
 *
 * @param {string} orderId - 受注ID（例: "ORD001"）
 * @returns {number} 削除した行数（0=該当なし、シートなし）
 *
 * 使用例:
 * const deletedCount = deleteSagawaCSVByOrderId("ORD001");
 * // 返却: 2 （2商品分の佐川CSV行が削除された）
 *
 * 佐川CSVシート構造:
 * - 「お客様管理番号」列に受注IDが保存されている
 * - 1受注IDで複数行存在する場合がある（複数商品の場合）
 * - 送り状作成時に createCsv() で生成される
 *
 * @see deleteOrderByOrderId() - 受注データ削除
 * @see deleteYamatoCSVByOrderId() - ヤマトCSVデータ削除
 * @see createCsv() - 佐川CSV生成（createCSVFile.js）
 *
 * 呼び出し元: 受注削除処理、受注修正処理
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
