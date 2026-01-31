/**
 * 商品マスタ管理モジュール
 *
 * 商品マスタの一覧表示、登録、編集、削除機能を提供します。
 * 商品分類マスタの管理も含みます。
 *
 * シート構造:
 * - 商品シート（17列）: 商品分類, 商品名, 価格（P), 商品原価（V）, 粗利（M), 在庫数, 単位, 税率, 備考, 更新日, 送り状品名, 略称, 勘定科目, 税区分, 部門, 品目, メモ
 * - 商品分類シート（3列）: 分類ID, 商品分類, 更新日
 */

/**
 * 商品シートのヘッダー定義
 * @returns {Array<string>} ヘッダー配列
 */
function getProductSheetHeaders() {
  return [
    '商品分類',
    '商品名',
    '価格（P)',
    '商品原価（V）',
    '粗利（M)',
    '在庫数',
    '単位',
    '税率',
    '備考',
    '更新日',
    '送り状品名',
    '略称',
    '勘定科目',
    '税区分',
    '部門',
    '品目',
    'メモ'
  ];
}

/**
 * 商品分類シートのヘッダー定義
 * @returns {Array<string>} ヘッダー配列
 */
function getProductCategoryHeaders() {
  return ['分類ID', '商品分類', '更新日'];
}

/**
 * 商品一覧を取得する
 * @param {Object} filters - フィルター条件（オプション）
 * @param {string} filters.category - 商品分類でフィルター
 * @param {string} filters.keyword - 商品名・略称でキーワード検索
 * @returns {Array<Object>} 商品データの配列
 */
function getProductList(filters) {
  const products = getAllRecords('商品');

  if (!products || products.length === 0) {
    return [];
  }

  let result = products;

  // フィルター適用
  if (filters) {
    if (filters.category && filters.category !== '') {
      result = result.filter(function(p) {
        return p['商品分類'] === filters.category;
      });
    }

    if (filters.keyword && filters.keyword !== '') {
      var keyword = filters.keyword.toLowerCase();
      result = result.filter(function(p) {
        var name = (p['商品名'] || '').toLowerCase();
        var abbr = (p['略称'] || '').toLowerCase();
        return name.indexOf(keyword) !== -1 || abbr.indexOf(keyword) !== -1;
      });
    }
  }

  // 行番号を付与（編集・削除時に使用）
  return result.map(function(p) {
    // 元のインデックスを取得するため、productsから検索
    var originalIndex = products.indexOf(p);
    p.rowIndex = originalIndex + 2; // ヘッダー行を考慮
    return p;
  });
}

/**
 * 商品一覧をHTML用データとして取得
 * @param {Object} filters - フィルター条件（オプション）
 * @returns {string} JSON文字列
 */
function getProductListData(filters) {
  var products = getProductList(filters);

  // 日付を文字列に変換
  var result = products.map(function(p) {
    return {
      rowIndex: p.rowIndex,
      category: p['商品分類'] || '',
      name: p['商品名'] || '',
      price: p['価格（P)'] || 0,
      cost: p['商品原価（V）'] || 0,
      margin: p['粗利（M)'] || 0,
      stock: p['在庫数'] || 0,
      unit: p['単位'] || '',
      taxRate: p['税率'] || '',
      note: p['備考'] || '',
      updatedAt: formatDateValue(p['更新日']),
      invoiceName: p['送り状品名'] || '',
      abbr: p['略称'] || '',
      accountTitle: p['勘定科目'] || '',
      taxCategory: p['税区分'] || '',
      department: p['部門'] || '',
      item: p['品目'] || '',
      memo: p['メモ'] || ''
    };
  });

  return JSON.stringify(result);
}

/**
 * 日付値をフォーマット
 * @param {*} value - 日付値
 * @returns {string} フォーマット済み日付文字列
 */
function formatDateValue(value) {
  if (!value) return '';
  if (value instanceof Date) {
    return Utilities.formatDate(value, 'JST', 'yyyy/MM/dd');
  }
  return String(value);
}

/**
 * 商品を新規登録する
 * @param {Object} data - 商品データ
 * @returns {Object} 結果 { success: boolean, message: string }
 */
function addProduct(data) {
  try {
    // 必須項目バリデーション
    if (!data || typeof data !== 'object') {
      return { success: false, message: '入力データが不正です' };
    }
    var name = (data.name || '').toString().trim();
    if (!name) {
      return { success: false, message: 'name is required' };
    }
    // 他の必須項目があればここで追加
    // 例: price必須の場合
    // var priceVal = (data.price !== undefined && data.price !== null) ? String(data.price).trim() : '';
    // if (!priceVal) {
    //   return { success: false, message: 'price is required' };
    // }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
      ss = SpreadsheetApp.openById(getMasterSpreadsheetId());
    }
    var sheet = ss.getSheetByName('商品');

    if (!sheet) {
      return { success: false, message: '商品シートが見つかりません' };
    }

    var dateNow = Utilities.formatDate(new Date(), 'JST', 'yyyy/MM/dd');

    // 粗利を計算
    var price = Number(data.price) || 0;
    var cost = Number(data.cost) || 0;
    var margin = price - cost;

    var row = [
      data.category || '',
      name,
      price,
      cost,
      margin,
      data.stock || '',
      data.unit || '',
      data.taxRate || '',
      data.note || '',
      dateNow,
      data.invoiceName || '',
      data.abbr || '',
      data.accountTitle || '',
      data.taxCategory || '',
      data.department || '',
      data.item || '',
      data.memo || ''
    ];

    sheet.appendRow(row);

    // キャッシュを更新
    try {
      refreshMasterDataCache();
    } catch (e) {
      Logger.log('キャッシュ更新エラー: ' + e.message);
    }

    return { success: true, message: '商品を登録しました' };

  } catch (error) {
    Logger.log('addProduct error: ' + error.message);
    return { success: false, message: '登録に失敗しました: ' + error.message };
  }
}

/**
 * 商品を更新する
 * @param {number} rowIndex - 行番号
 * @param {Object} data - 商品データ
 * @returns {Object} 結果 { success: boolean, message: string }
 */
function updateProduct(rowIndex, data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
      ss = SpreadsheetApp.openById(getMasterSpreadsheetId());
    }
    var sheet = ss.getSheetByName('商品');

    if (!sheet) {
      return { success: false, message: '商品シートが見つかりません' };
    }

    var dateNow = Utilities.formatDate(new Date(), 'JST', 'yyyy/MM/dd');

    // 粗利を計算
    var price = Number(data.price) || 0;
    var cost = Number(data.cost) || 0;
    var margin = price - cost;

    var row = [
      data.category || '',
      data.name || '',
      price,
      cost,
      margin,
      data.stock || '',
      data.unit || '',
      data.taxRate || '',
      data.note || '',
      dateNow,
      data.invoiceName || '',
      data.abbr || '',
      data.accountTitle || '',
      data.taxCategory || '',
      data.department || '',
      data.item || '',
      data.memo || ''
    ];

    sheet.getRange(rowIndex, 1, 1, row.length).setValues([row]);

    // キャッシュを更新
    try {
      refreshMasterDataCache();
    } catch (e) {
      Logger.log('キャッシュ更新エラー: ' + e.message);
    }

    return { success: true, message: '商品を更新しました' };

  } catch (error) {
    Logger.log('updateProduct error: ' + error.message);
    return { success: false, message: '更新に失敗しました: ' + error.message };
  }
}

/**
 * 商品を削除する
 * @param {number} rowIndex - 行番号
 * @returns {Object} 結果 { success: boolean, message: string }
 */
function deleteProduct(rowIndex) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
      ss = SpreadsheetApp.openById(getMasterSpreadsheetId());
    }
    var sheet = ss.getSheetByName('商品');

    if (!sheet) {
      return { success: false, message: '商品シートが見つかりません' };
    }

    var lastRow = sheet.getLastRow();
    if (!Number.isInteger(rowIndex) || !isFinite(rowIndex) || rowIndex <= 1 || rowIndex > lastRow) {
      return { success: false, message: 'Invalid row index' };
    }

    sheet.deleteRow(rowIndex);

    // キャッシュを更新
    try {
      refreshMasterDataCache();
    } catch (e) {
      Logger.log('キャッシュ更新エラー: ' + e.message);
    }

    return { success: true, message: '商品を削除しました' };

  } catch (error) {
    Logger.log('deleteProduct error: ' + error.message);
    return { success: false, message: '削除に失敗しました: ' + error.message };
  }
}

/**
 * 商品分類一覧を取得する
 * @returns {Array<Object>} 商品分類データの配列
 */
function getProductCategories() {
  var categories = getAllRecords('商品分類');

  if (!categories || categories.length === 0) {
    return [];
  }

  return categories.map(function(c, index) {
    return {
      id: c['分類ID'] || '',
      name: c['商品分類'] || '',
      updatedAt: formatDateValue(c['更新日']),
      rowIndex: index + 2
    };
  });
}

/**
 * 商品分類一覧をHTML用データとして取得
 * @returns {string} JSON文字列
 */
function getProductCategoriesData() {
  return JSON.stringify(getProductCategories());
}

/**
 * 商品分類を新規登録する
 * @param {string} categoryName - 分類名
 * @returns {Object} 結果 { success: boolean, message: string, category: Object }
 */
function addProductCategory(categoryName) {
  try {
    if (!categoryName || categoryName.trim() === '') {
      return { success: false, message: '分類名を入力してください' };
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
      ss = SpreadsheetApp.openById(getMasterSpreadsheetId());
    }
    var sheet = ss.getSheetByName('商品分類');

    if (!sheet) {
      return { success: false, message: '商品分類シートが見つかりません' };
    }

    // 既存の分類をチェック（重複防止）
    var existingCategories = getProductCategories();
    for (var i = 0; i < existingCategories.length; i++) {
      if (existingCategories[i].name === categoryName.trim()) {
        return { success: false, message: 'この分類名は既に存在します' };
      }
    }

    // 最大IDを取得
    var maxId = 0;
    existingCategories.forEach(function(c) {
      var id = parseInt(c.id, 10);
      if (!isNaN(id) && id > maxId) {
        maxId = id;
      }
    });

    var newId = maxId + 1;
    var dateNow = Utilities.formatDate(new Date(), 'JST', 'yyyy/MM/dd');

    var row = [newId, categoryName.trim(), dateNow];
    sheet.appendRow(row);

    // キャッシュを更新
    try {
      refreshMasterDataCache();
    } catch (e) {
      Logger.log('キャッシュ更新エラー: ' + e.message);
    }

    return {
      success: true,
      message: '商品分類を登録しました',
      category: {
        id: newId,
        name: categoryName.trim(),
        updatedAt: dateNow
      }
    };

  } catch (error) {
    Logger.log('addProductCategory error: ' + error.message);
    return { success: false, message: '登録に失敗しました: ' + error.message };
  }
}

/**
 * 商品分類を削除する
 * @param {number} rowIndex - 行番号
 * @returns {Object} 結果 { success: boolean, message: string }
 */
function deleteProductCategory(rowIndex) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
      ss = SpreadsheetApp.openById(getMasterSpreadsheetId());
    }
    var sheet = ss.getSheetByName('商品分類');

    if (!sheet) {
      return { success: false, message: '商品分類シートが見つかりません' };
    }

    var lastRow = sheet.getLastRow();
    if (!Number.isInteger(rowIndex) || !isFinite(rowIndex) || rowIndex <= 1 || rowIndex > lastRow) {
      return { success: false, message: '無効な行番号' };
    }

    // 削除する分類名を取得
    var categoryName = sheet.getRange(rowIndex, 2).getValue();

    // この分類を使用している商品があるかチェック
    var products = getAllRecords('商品');
    var usingProducts = products.filter(function(p) {
      return p['商品分類'] === categoryName;
    });

    if (usingProducts.length > 0) {
      return {
        success: false,
        message: 'この分類は ' + usingProducts.length + ' 件の商品で使用されているため削除できません'
      };
    }

    sheet.deleteRow(rowIndex);

    // キャッシュを更新
    try {
      refreshMasterDataCache();
    } catch (e) {
      Logger.log('キャッシュ更新エラー: ' + e.message);
    }

    return { success: true, message: '商品分類を削除しました' };

  } catch (error) {
    Logger.log('deleteProductCategory error: ' + error.message);
    return { success: false, message: '削除に失敗しました: ' + error.message };
  }
}

/**
 * 商品詳細を取得する（編集用）
 * @param {number} rowIndex - 行番号
 * @returns {Object|null} 商品詳細データ
 */
function getProductDetail(rowIndex) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
      ss = SpreadsheetApp.openById(getMasterSpreadsheetId());
    }
    var sheet = ss.getSheetByName('商品');

    if (!sheet) {
      return null;
    }

    var lastRow = sheet.getLastRow();
    if (!Number.isInteger(rowIndex) || !isFinite(rowIndex) || rowIndex < 2 || rowIndex > lastRow) {
      return null;
    }

    var headers = getProductSheetHeaders();
    var values = sheet.getRange(rowIndex, 1, 1, headers.length).getValues()[0];

    var product = {};
    for (var i = 0; i < headers.length; i++) {
      product[headers[i]] = values[i];
    }
    product.rowIndex = rowIndex;

    return product;

  } catch (error) {
    Logger.log('getProductDetail error: ' + error.message);
    return null;
  } 
}
