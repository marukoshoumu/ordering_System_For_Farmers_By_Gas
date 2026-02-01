/**
 * 顧客マスタ管理モジュール
 *
 * 顧客マスタの一覧表示、登録、編集、削除機能を提供します。
 *
 * シート構造:
 * - 顧客情報シート（18列）: 顧客分類, 表示名, フリガナ, 会社名, 部署, 役職, 氏名, 郵便番号, 住所１, 住所２, TEL, 携帯電話, FAX, メールアドレス, 請求書有無, 入金期日, 消費税の表示方法, 備考
 */

/**
 * 顧客シートのヘッダー定義
 * @returns {Array<string>} ヘッダー配列
 */
function getCustomerSheetHeaders() {
  return [
    '顧客分類',
    '表示名',
    'フリガナ',
    '会社名',
    '部署',
    '役職',
    '氏名',
    '郵便番号',
    '住所１',
    '住所２',
    'TEL',
    '携帯電話',
    'FAX',
    'メールアドレス',
    '請求書有無',
    '入金期日',
    '消費税の表示方法',
    '備考'
  ];
}

/**
 * 顧客分類の定義
 * @returns {Array<Object>} 顧客分類配列
 */
function getCustomerCategories() {
  return [
    { id: '1', name: '一次卸し・大口' },
    { id: '2', name: '小売店' },
    { id: '3', name: '飲食店・ホテル' },
    { id: '4', name: '地元飲食店' },
    { id: '5', name: '通信販売' },
    { id: '6', name: '一般販売' },
    { id: '7', name: 'その他' },
    { id: '8', name: '職域' }
  ];
}

/**
 * 顧客分類IDから名前を取得
 * @param {string} id - 顧客分類ID
 * @returns {string} 顧客分類名
 */
function getCustomerCategoryName(id) {
  var categories = getCustomerCategories();
  for (var i = 0; i < categories.length; i++) {
    if (categories[i].id === String(id)) {
      return categories[i].name;
    }
  }
  return id || '';
}

/**
 * 顧客一覧を取得する
 * @param {Object} filters - フィルター条件（オプション）
 * @param {string} filters.category - 顧客分類でフィルター（"1"〜"8"）
 * @param {string} filters.keyword - 表示名・会社名・氏名・フリガナでキーワード検索
 * @returns {Array<Object>} 顧客データの配列
 */
function getCustomerList(filters) {
  var customers = getAllRecords('顧客情報');

  if (!customers || customers.length === 0) {
    return [];
  }

  var result = customers;

  // フィルター適用
  if (filters) {
    if (filters.category && filters.category !== '') {
      result = result.filter(function(c) {
        return String(c['顧客分類']) === String(filters.category);
      });
    }

    if (filters.keyword && filters.keyword !== '') {
      var keyword = filters.keyword.toLowerCase();
      result = result.filter(function(c) {
        var showName = (c['表示名'] || '').toLowerCase();
        var companyName = (c['会社名'] || '').toLowerCase();
        var personName = (c['氏名'] || '').toLowerCase();
        var furigana = (c['フリガナ'] || '').toLowerCase();
        return showName.indexOf(keyword) !== -1 ||
               companyName.indexOf(keyword) !== -1 ||
               personName.indexOf(keyword) !== -1 ||
               furigana.indexOf(keyword) !== -1;
      });
    }
  }

  // 行番号を付与（編集・削除時に使用）
  return result.map(function(c) {
    var originalIndex = customers.indexOf(c);
    c.rowIndex = originalIndex + 2; // ヘッダー行を考慮
    return c;
  });
}

/**
 * 顧客一覧をHTML用データとして取得
 * @param {Object} filters - フィルター条件（オプション）
 * @returns {string} JSON文字列
 */
function getCustomerListData(filters) {
  var customers = getCustomerList(filters);

  var result = customers.map(function(c) {
    return {
      rowIndex: c.rowIndex,
      class: c['顧客分類'] || '',
      className: getCustomerCategoryName(c['顧客分類']),
      showName: c['表示名'] || '',
      furigana: c['フリガナ'] || '',
      companyName: c['会社名'] || '',
      department: c['部署'] || '',
      post: c['役職'] || '',
      personName: c['氏名'] || '',
      zipcode: c['郵便番号'] || '',
      address1: c['住所１'] || '',
      address2: c['住所２'] || '',
      tel: c['TEL'] || '',
      phone: c['携帯電話'] || '',
      fax: c['FAX'] || '',
      email: c['メールアドレス'] || '',
      bill: c['請求書有無'] || '',
      depositDate: c['入金期日'] || '',
      taxDisplay: c['消費税の表示方法'] || '',
      memo: c['備考'] || ''
    };
  });

  return JSON.stringify(result);
}


/**
 * 顧客を新規登録する（customerListCode用ラッパー）
 * @param {Object} data - 顧客データ
 * @returns {Object} 結果 { success: boolean, message: string }
 */
function addCustomerFromList(data) {
  try {
    // 必須項目バリデーション
    if (!data || typeof data !== 'object') {
      return { success: false, message: '入力データが不正です' };
    }

    var companyName = (data.companyName || '').toString().trim();
    var personName = (data.personName || '').toString().trim();

    if (!companyName && !personName) {
      return { success: false, message: '会社名または氏名を入力してください' };
    }

    var zipcode = (data.zipcode || '').toString().trim();
    if (!zipcode) {
      return { success: false, message: '郵便番号を入力してください' };
    }
    if (!/^\d{7}$/.test(zipcode)) {
      return { success: false, message: '郵便番号は7桁の数字で入力してください' };
    }

    var address1 = (data.address1 || '').toString().trim();
    if (!address1) {
      return { success: false, message: '住所１を入力してください' };
    }

    var tel = (data.tel || '').toString().trim();
    if (!tel) {
      return { success: false, message: '電話番号を入力してください' };
    }
    if (!/^0\d{9,10}$/.test(tel)) {
      return { success: false, message: '電話番号は10〜11桁の数字で入力してください' };
    }

    // オプション項目のバリデーション
    var phone = (data.phone || '').toString().trim();
    if (phone && !/^0\d{9,10}$/.test(phone)) {
      return { success: false, message: '携帯電話は10〜11桁の数字で入力してください' };
    }

    var fax = (data.fax || '').toString().trim();
    if (fax && !/^0\d{9,10}$/.test(fax)) {
      return { success: false, message: 'FAXは10〜11桁の数字で入力してください' };
    }

    var email = (data.email || '').toString().trim();
    if (email && !/^([a-zA-Z0-9])+([a-zA-Z0-9\._-])*@([a-zA-Z0-9_-])+([a-zA-Z0-9\._-]+)+$/.test(email)) {
      return { success: false, message: 'メールアドレスは正しい形式で入力してください' };
    }

    // 既存のaddCustomer関数用のrecord形式に変換
    var record = {
      customerClassDg: data.class || '',
      customerShowNameDg: data.showName || '',
      customerFuriganaDg: data.furigana || '',
      customerNameDg: companyName,
      customerDepDg: data.department || '',
      customerPostDg: data.post || '',
      customerIdentDg: personName,
      customerZipcodeDg: zipcode,
      customerAddress1Dg: address1,
      customerAddress2Dg: data.address2 || '',
      customerTelDg: tel,
      customerPhoneDg: phone,
      customerFaxDg: fax,
      customerMailDg: email,
      customerBillDg: data.bill || '',
      customerDepositDateDg: data.depositDate || '',
      customerTaxDisplayDg: data.taxDisplay || '',
      customerMemoDg: data.memo || ''
    };

    // 既存のaddCustomer関数を呼び出し
    addCustomer(JSON.stringify(record));

    // キャッシュを更新
    try {
      refreshMasterDataCache();
    } catch (e) {
      Logger.log('キャッシュ更新エラー: ' + e.message);
    }

    return { success: true, message: '顧客情報を登録しました' };

  } catch (error) {
    Logger.log('addCustomerFromList error: ' + error.message);
    return { success: false, message: '登録に失敗しました: ' + error.message };
  }
}

/**
 * 顧客を更新する
 * @param {number} rowIndex - 行番号
 * @param {Object} data - 顧客データ
 * @returns {Object} 結果 { success: boolean, message: string }
 */
function updateCustomer(rowIndex, data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
      ss = SpreadsheetApp.openById(getMasterSpreadsheetId());
    }
    var sheet = ss.getSheetByName('顧客情報');
    var sheetBk = ss.getSheetByName('顧客情報BK');

    if (!sheet) {
      return { success: false, message: '顧客情報シートが見つかりません' };
    }

    var validation = normalizeAndValidateRowIndex(rowIndex, sheet);
    if (!validation.success) {
      return { success: false, message: validation.message };
    }
    rowIndex = validation.rowIndex;

    // バリデーション
    var companyName = (data.companyName || '').toString().trim();
    var personName = (data.personName || '').toString().trim();

    if (!companyName && !personName) {
      return { success: false, message: '会社名または氏名を入力してください' };
    }

    var zipcode = (data.zipcode || '').toString().trim();
    if (!zipcode) {
      return { success: false, message: '郵便番号を入力してください' };
    }
    if (!/^\d{7}$/.test(zipcode)) {
      return { success: false, message: '郵便番号は7桁の数字で入力してください' };
    }

    var address1 = (data.address1 || '').toString().trim();
    if (!address1) {
      return { success: false, message: '住所１を入力してください' };
    }

    var tel = (data.tel || '').toString().trim();
    if (!tel) {
      return { success: false, message: '電話番号を入力してください' };
    }
    if (!/^0\d{9,10}$/.test(tel)) {
      return { success: false, message: '電話番号は10〜11桁の数字で入力してください' };
    }

    var phone = (data.phone || '').toString().trim();
    if (phone && !/^0\d{9,10}$/.test(phone)) {
      return { success: false, message: '携帯電話は10〜11桁の数字で入力してください' };
    }

    var fax = (data.fax || '').toString().trim();
    if (fax && !/^0\d{9,10}$/.test(fax)) {
      return { success: false, message: 'FAXは10〜11桁の数字で入力してください' };
    }

    var email = (data.email || '').toString().trim();
    if (email && !/^([a-zA-Z0-9])+([a-zA-Z0-9\._-])*@([a-zA-Z0-9_-])+([a-zA-Z0-9\._-]+)+$/.test(email)) {
      return { success: false, message: 'メールアドレスは正しい形式で入力してください' };
    }

    var row = [
      data.class || '',
      data.showName || '',
      data.furigana || '',
      companyName,
      data.department || '',
      data.post || '',
      personName,
      zipcode,
      address1,
      data.address2 || '',
      tel,
      phone,
      fax,
      email,
      data.bill || '',
      data.depositDate || '',
      data.taxDisplay || '',
      data.memo || ''
    ];

    // メインシートを更新
    sheet.getRange(rowIndex, 1, 1, row.length)
      .setNumberFormat('@')
      .setValues([row])
      .setBorder(true, true, true, true, true, true);

    // BKシートも更新
    if (sheetBk) {
      sheetBk.getRange(rowIndex, 1, 1, row.length)
        .setNumberFormat('@')
        .setValues([row])
        .setBorder(true, true, true, true, true, true);
    }

    // キャッシュを更新
    try {
      refreshMasterDataCache();
    } catch (e) {
      Logger.log('キャッシュ更新エラー: ' + e.message);
    }

    return { success: true, message: '顧客情報を更新しました' };

  } catch (error) {
    Logger.log('updateCustomer error: ' + error.message);
    return { success: false, message: '更新に失敗しました: ' + error.message };
  }
}

/**
 * 顧客を削除する
 * @param {number} rowIndex - 行番号
 * @returns {Object} 結果 { success: boolean, message: string }
 */
function deleteCustomer(rowIndex) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
      ss = SpreadsheetApp.openById(getMasterSpreadsheetId());
    }
    var sheet = ss.getSheetByName('顧客情報');
    var sheetBk = ss.getSheetByName('顧客情報BK');

    if (!sheet) {
      return { success: false, message: '顧客情報シートが見つかりません' };
    }

    var validation = normalizeAndValidateRowIndex(rowIndex, sheet);
    if (!validation.success) {
      return { success: false, message: validation.message };
    }
    rowIndex = validation.rowIndex;

    // メインシートから削除
    sheet.deleteRow(rowIndex);

    // BKシートからも削除
    if (sheetBk) {
      sheetBk.deleteRow(rowIndex);
    }

    // キャッシュを更新
    try {
      refreshMasterDataCache();
    } catch (e) {
      Logger.log('キャッシュ更新エラー: ' + e.message);
    }

    return { success: true, message: '顧客情報を削除しました' };

  } catch (error) {
    Logger.log('deleteCustomer error: ' + error.message);
    return { success: false, message: '削除に失敗しました: ' + error.message };
  }
}

/**
 * 顧客詳細を取得する（編集用）
 * @param {number} rowIndex - 行番号
 * @returns {Object|null} 顧客詳細データ
 */
function getCustomerDetail(rowIndex) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
      ss = SpreadsheetApp.openById(getMasterSpreadsheetId());
    }
    var sheet = ss.getSheetByName('顧客情報');

    if (!sheet) {
      return null;
    }

    var validation = normalizeAndValidateRowIndex(rowIndex, sheet);
    if (!validation.success) {
      return null;
    }
    rowIndex = validation.rowIndex;

    var headers = getCustomerSheetHeaders();
    var values = sheet.getRange(rowIndex, 1, 1, headers.length).getValues()[0];

    var customer = {};
    for (var i = 0; i < headers.length; i++) {
      customer[headers[i]] = values[i];
    }
    customer.rowIndex = rowIndex;

    return customer;

  } catch (error) {
    Logger.log('getCustomerDetail error: ' + error.message);
    return null;
  }
}

/**
 * 顧客分類一覧をHTML用データとして取得
 * @returns {string} JSON文字列
 */
function getCustomerCategoriesData() {
  return JSON.stringify(getCustomerCategories());
}

/**
 * 顧客情報と発送先情報を同時登録（customerList用）
 * @param {Object} data - 顧客データ（customerList.html のフォームから）
 * @returns {Object} 結果 { success: boolean, message: string }
 */
function addCustomerShippingToFromList(data) {
  try {
    // 必須項目バリデーション（addCustomerFromList と同じ）
    if (!data || typeof data !== 'object') {
      return { success: false, message: '入力データが不正です' };
    }

    var companyName = (data.companyName || '').toString().trim();
    var personName = (data.personName || '').toString().trim();

    if (!companyName && !personName) {
      return { success: false, message: '会社名または氏名を入力してください' };
    }

    var zipcode = (data.zipcode || '').toString().trim();
    if (!zipcode) {
      return { success: false, message: '郵便番号を入力してください' };
    }
    if (!/^\d{7}$/.test(zipcode)) {
      return { success: false, message: '郵便番号は7桁の数字で入力してください' };
    }

    var address1 = (data.address1 || '').toString().trim();
    if (!address1) {
      return { success: false, message: '住所１を入力してください' };
    }

    var tel = (data.tel || '').toString().trim();
    if (!tel) {
      return { success: false, message: '電話番号を入力してください' };
    }
    if (!/^0\d{9,10}$/.test(tel)) {
      return { success: false, message: '電話番号は10〜11桁の数字で入力してください' };
    }

    // オプション項目のバリデーション
    var phone = (data.phone || '').toString().trim();
    if (phone && !/^0\d{9,10}$/.test(phone)) {
      return { success: false, message: '携帯電話は10〜11桁の数字で入力してください' };
    }

    var fax = (data.fax || '').toString().trim();
    if (fax && !/^0\d{9,10}$/.test(fax)) {
      return { success: false, message: 'FAXは10〜11桁の数字で入力してください' };
    }

    var email = (data.email || '').toString().trim();
    if (email && !/^([a-zA-Z0-9])+([a-zA-Z0-9\._-])*@([a-zA-Z0-9_-])+([a-zA-Z0-9\._-]+)+$/.test(email)) {
      return { success: false, message: 'メールアドレスは正しい形式で入力してください' };
    }

    // customerCode.js の addCustomerShippingTo() に渡す形式に変換
    var record = {
      customerClassDg: data.class || '',
      customerShowNameDg: data.showName || '',
      customerFuriganaDg: data.furigana || '',
      customerNameDg: companyName,
      customerDepDg: data.department || '',
      customerPostDg: data.post || '',
      customerIdentDg: personName,
      customerZipcodeDg: zipcode,
      customerAddress1Dg: address1,
      customerAddress2Dg: data.address2 || '',
      customerTelDg: tel,
      customerPhoneDg: phone,
      customerFaxDg: fax,
      customerMailDg: email,
      customerBillDg: data.bill || '',
      customerDepositDateDg: data.depositDate || '翌末日',
      customerTaxDisplayDg: data.taxDisplay || '',
      customerMemoDg: data.memo || ''
    };

    // customerCode.js の addCustomerShippingTo() を呼び出し
    addCustomerShippingTo(JSON.stringify(record));

    // キャッシュを更新
    try {
      refreshMasterDataCache();
    } catch (e) {
      Logger.log('キャッシュ更新エラー: ' + e.message);
    }

    return { success: true, message: '顧客情報と発送先情報を登録しました' };

  } catch (error) {
    Logger.log('addCustomerShippingToFromList error: ' + error.message);
    return { success: false, message: '登録に失敗しました: ' + error.message };
  }
}

/**
 * 顧客情報から対応する発送先情報を検索
 * @param {string} companyName - 会社名
 * @param {string} personName - 氏名
 * @param {string} zipcode - 郵便番号
 * @returns {Object} { found: boolean, rowIndex?: number, data?: Object, reason?: string, error?: string }
 */
function findShippingToByCustomer(companyName, personName, zipcode) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
      ss = SpreadsheetApp.openById(getMasterSpreadsheetId());
    }
    var sheet = ss.getSheetByName('発送先情報');

    if (!sheet) {
      return { found: false, reason: 'no_sheet' };
    }

    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      return { found: false, reason: 'no_data' };
    }

    // 会社名・氏名・郵便番号で一致する発送先を検索
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var rowCompanyName = (row[0] || '').toString().trim();
      var rowPersonName = (row[2] || '').toString().trim();
      var rowZipcode = (row[3] || '').toString().trim();

      if (rowCompanyName === companyName &&
          rowPersonName === personName &&
          rowZipcode === zipcode) {
        return {
          found: true,
          rowIndex: i + 1, // ヘッダー行を考慮
          data: {
            companyName: rowCompanyName,
            department: (row[1] || '').toString().trim(),
            personName: rowPersonName,
            zipcode: rowZipcode,
            address1: (row[4] || '').toString().trim(),
            address2: (row[5] || '').toString().trim(),
            tel: (row[6] || '').toString().trim(),
            email: (row[7] || '').toString().trim(),
            memo: (row[8] || '').toString().trim()
          }
        };
      }
    }

    return { found: false, reason: 'not_found' };
  } catch (error) {
    Logger.log('findShippingToByCustomer error: ' + error.message);
    return { found: false, error: error.message };
  }
}

/**
 * 顧客情報と発送先情報を同時更新（customerList用）
 * @param {number} customerRowIndex - 顧客の行番号
 * @param {Object} customerData - 顧客データ
 * @returns {Object} 結果 { success: boolean, message: string }
 */
function updateCustomerShippingTo(customerRowIndex, customerData) {
  try {
    // まず顧客情報を更新
    var customerResult = updateCustomer(customerRowIndex, customerData);
    if (!customerResult.success) {
      return customerResult;
    }

    // 発送先情報を検索
    var companyName = (customerData.companyName || '').toString().trim();
    var personName = (customerData.personName || '').toString().trim();
    var zipcode = (customerData.zipcode || '').toString().trim();

    var shippingTo = findShippingToByCustomer(companyName, personName, zipcode);

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
      ss = SpreadsheetApp.openById(getMasterSpreadsheetId());
    }
    var shippingToSheet = ss.getSheetByName('発送先情報');
    var shippingToBkSheet = ss.getSheetByName('発送先情報BK');

    if (!shippingToSheet) {
      return { success: false, message: '発送先情報シートが見つかりません' };
    }

    // 発送先情報のデータを構築（customerCode.js の addCustomerShippingTo と同じ形式）
    var shippingToRow = [
      companyName,
      customerData.department || '',
      personName,
      zipcode,
      customerData.address1 || '',
      customerData.address2 || '',
      customerData.tel || '',
      customerData.email || '',
      customerData.memo || ''
    ];

    if (shippingTo.found) {
      // 既存の発送先情報を更新
      var shippingToRowIndex = shippingTo.rowIndex;

      shippingToSheet.getRange(shippingToRowIndex, 1, 1, shippingToRow.length)
        .setNumberFormat('@')
        .setValues([shippingToRow])
        .setBorder(true, true, true, true, true, true);

      if (shippingToBkSheet) {
        shippingToBkSheet.getRange(shippingToRowIndex, 1, 1, shippingToRow.length)
          .setNumberFormat('@')
          .setValues([shippingToRow])
          .setBorder(true, true, true, true, true, true);
      }
    } else {
      // 新規作成
      var shippingToLastRow = shippingToSheet.getLastRow();
      var shippingToBkLastRow = shippingToBkSheet ? shippingToBkSheet.getLastRow() : 0;

      shippingToSheet.getRange(shippingToLastRow + 1, 1, 1, shippingToRow.length)
        .setNumberFormat('@')
        .setValues([shippingToRow])
        .setBorder(true, true, true, true, true, true);

      if (shippingToBkSheet) {
        shippingToBkSheet.getRange(shippingToBkLastRow + 1, 1, 1, shippingToRow.length)
          .setNumberFormat('@')
          .setValues([shippingToRow])
          .setBorder(true, true, true, true, true, true);
      }
    }

    // キャッシュを更新
    try {
      refreshMasterDataCache();
    } catch (e) {
      Logger.log('キャッシュ更新エラー: ' + e.message);
    }

    return { success: true, message: '顧客情報と発送先情報を更新しました' };

  } catch (error) {
    Logger.log('updateCustomerShippingTo error: ' + error.message);
    return { success: false, message: '更新に失敗しました: ' + error.message };
  }
}
