/**
 * 発送先マスタ管理モジュール
 *
 * 発送先マスタの一覧表示、登録、編集、削除機能を提供します。
 *
 * シート構造（9列）:
 * - 発送先情報シート: 会社名, 部署, 氏名, 郵便番号, 住所１, 住所２, TEL, メールアドレス, 備考
 * - 発送先情報BKシート: バックアップ（同構造）
 */

/**
 * 発送先シートのヘッダー定義
 * @returns {Array<string>} ヘッダー配列
 */
function getShippingToSheetHeaders() {
  return [
    '会社名',
    '部署',
    '氏名',
    '郵便番号',
    '住所１',
    '住所２',
    'TEL',
    'メールアドレス',
    '備考'
  ];
}

/**
 * 発送先一覧を取得する
 * @param {Object} filters - フィルター条件（オプション）
 * @param {string} filters.keyword - 会社名・氏名でキーワード検索
 * @returns {Array<Object>} 発送先データの配列
 */
function getShippingToList(filters) {
  var shippingTos = getAllRecords('発送先情報');

  if (!shippingTos || shippingTos.length === 0) {
    return [];
  }

  var result = shippingTos;

  // フィルター適用
  if (filters) {
    if (filters.keyword && filters.keyword !== '') {
      var keyword = filters.keyword.toLowerCase();
      result = result.filter(function(s) {
        var companyName = (s['会社名'] || '').toLowerCase();
        var personName = (s['氏名'] || '').toLowerCase();
        return companyName.indexOf(keyword) !== -1 ||
               personName.indexOf(keyword) !== -1;
      });
    }
  }

  // 行番号を付与（編集・削除時に使用）
  return result.map(function(s) {
    var originalIndex = shippingTos.indexOf(s);
    s.rowIndex = originalIndex + 2; // ヘッダー行を考慮
    return s;
  });
}

/**
 * 発送先一覧をHTML用データとして取得
 * @param {Object} filters - フィルター条件（オプション）
 * @returns {string} JSON文字列
 */
function getShippingToListData(filters) {
  var shippingTos = getShippingToList(filters);

  var result = shippingTos.map(function(s) {
    return {
      rowIndex: s.rowIndex,
      companyName: s['会社名'] || '',
      department: s['部署'] || '',
      personName: s['氏名'] || '',
      zipcode: s['郵便番号'] || '',
      address1: s['住所１'] || '',
      address2: s['住所２'] || '',
      tel: s['TEL'] || '',
      email: s['メールアドレス'] || '',
      memo: s['備考'] || ''
    };
  });

  return JSON.stringify(result);
}

/**
 * rowIndexを正規化して検証する
 * @param {*} rowIndex - 行番号（数値または文字列）
 * @param {Sheet} sheet - シートオブジェクト
 * @returns {Object} { success: boolean, rowIndex?: number, message?: string }
 */
function normalizeAndValidateShippingToRowIndex(rowIndex, sheet) {
  var normalized = Number(rowIndex);
  if (!isFinite(normalized)) {
    return { success: false, message: '無効なrowIndex' };
  }

  normalized = Math.floor(normalized);
  var maxRow = sheet.getLastRow();

  if (normalized < 2 || normalized > maxRow) {
    return { success: false, message: '無効なrowIndex' };
  }

  return { success: true, rowIndex: normalized };
}

/**
 * 発送先を新規登録する（shippingToListCode用ラッパー）
 * @param {Object} data - 発送先データ
 * @returns {Object} 結果 { success: boolean, message: string }
 */
function addShippingToFromList(data) {
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
    var email = (data.email || '').toString().trim();
    if (email && !/^([a-zA-Z0-9])+([a-zA-Z0-9\._-])*@([a-zA-Z0-9_-])+([a-zA-Z0-9\._-]+)+$/.test(email)) {
      return { success: false, message: 'メールアドレスは正しい形式で入力してください' };
    }

    // 既存のaddShippingTo関数用のrecord形式に変換
    var record = {
      shippingToNameDg: companyName,
      shippingToDepDg: data.department || '',
      shippingToIdentDg: personName,
      shippingToZipcodeDg: zipcode,
      shippingToAddress1Dg: address1,
      shippingToAddress2Dg: data.address2 || '',
      shippingToTelDg: tel,
      shippingToMailDg: email,
      shippingToMemoDg: data.memo || ''
    };

    // 既存のaddShippingTo関数を呼び出し
    addShippingTo(JSON.stringify(record));

    // キャッシュを更新
    try {
      refreshMasterDataCache();
    } catch (e) {
      Logger.log('キャッシュ更新エラー: ' + e.message);
    }

    return { success: true, message: '発送先情報を登録しました' };

  } catch (error) {
    Logger.log('addShippingToFromList error: ' + error.message);
    return { success: false, message: '登録に失敗しました: ' + error.message };
  }
}

/**
 * 発送先を更新する
 * @param {number} rowIndex - 行番号
 * @param {Object} data - 発送先データ
 * @returns {Object} 結果 { success: boolean, message: string }
 */
function updateShippingTo(rowIndex, data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
      ss = SpreadsheetApp.openById(getMasterSpreadsheetId());
    }
    var sheet = ss.getSheetByName('発送先情報');
    var sheetBk = ss.getSheetByName('発送先情報BK');

    if (!sheet) {
      return { success: false, message: '発送先情報シートが見つかりません' };
    }

    var validation = normalizeAndValidateShippingToRowIndex(rowIndex, sheet);
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

    var email = (data.email || '').toString().trim();
    if (email && !/^([a-zA-Z0-9])+([a-zA-Z0-9\._-])*@([a-zA-Z0-9_-])+([a-zA-Z0-9\._-]+)+$/.test(email)) {
      return { success: false, message: 'メールアドレスは正しい形式で入力してください' };
    }

    var row = [
      companyName,
      data.department || '',
      personName,
      zipcode,
      address1,
      data.address2 || '',
      tel,
      email,
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

    return { success: true, message: '発送先情報を更新しました' };

  } catch (error) {
    Logger.log('updateShippingTo error: ' + error.message);
    return { success: false, message: '更新に失敗しました: ' + error.message };
  }
}

/**
 * 発送先を削除する
 * @param {number} rowIndex - 行番号
 * @returns {Object} 結果 { success: boolean, message: string }
 */
function deleteShippingTo(rowIndex) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
      ss = SpreadsheetApp.openById(getMasterSpreadsheetId());
    }
    var sheet = ss.getSheetByName('発送先情報');
    var sheetBk = ss.getSheetByName('発送先情報BK');

    if (!sheet) {
      return { success: false, message: '発送先情報シートが見つかりません' };
    }

    var validation = normalizeAndValidateShippingToRowIndex(rowIndex, sheet);
    if (!validation.success) {
      return { success: false, message: validation.message };
    }
    rowIndex = validation.rowIndex;

    // 行データをキャッシュ（ロールバック用）
    var cachedMainRow = null;
    var cachedBkRow = null;
    var mainDeletePerformed = false;
    var bkDeletePerformed = false;

    try {
      // メインシートの行データをキャッシュ
      var mainLastCol = sheet.getLastColumn();
      if (mainLastCol > 0) {
        cachedMainRow = sheet.getRange(rowIndex, 1, 1, mainLastCol).getValues();
      }

      // BKシートの行データをキャッシュ（存在する場合）
      if (sheetBk) {
        var bkLastCol = sheetBk.getLastColumn();
        if (bkLastCol > 0) {
          cachedBkRow = sheetBk.getRange(rowIndex, 1, 1, bkLastCol).getValues();
        }
      }

      // アトミック削除: BKシートから先に削除（存在する場合）
      if (sheetBk) {
        try {
          sheetBk.deleteRow(rowIndex);
          bkDeletePerformed = true;
        } catch (bkError) {
          Logger.log('BKシート削除エラー: ' + bkError.message);
          throw bkError; // BK削除失敗時は処理を中断
        }
      }

      // メインシートから削除
      try {
        sheet.deleteRow(rowIndex);
        mainDeletePerformed = true;
      } catch (mainError) {
        Logger.log('メインシート削除エラー: ' + mainError.message);
        
        // メインシート削除失敗時はBKシートを復元
        if (bkDeletePerformed && sheetBk && cachedBkRow) {
          try {
            sheetBk.insertRowBefore(rowIndex);
            var bkRange = sheetBk.getRange(rowIndex, 1, 1, cachedBkRow[0].length);
            bkRange.setValues(cachedBkRow);
            Logger.log('BKシートをロールバックしました');
          } catch (rollbackError) {
            Logger.log('BKシートロールバックエラー: ' + rollbackError.message);
            // ロールバック失敗も含めて元のエラーを再スロー
          }
        }
        throw mainError; // エラーを再スローして呼び出し元で処理可能にする
      }

    } catch (deleteError) {
      Logger.log('deleteShippingTo 削除処理エラー: ' + deleteError.message);
      throw deleteError; // エラーを再スロー
    }

    // キャッシュを更新
    try {
      refreshMasterDataCache();
    } catch (e) {
      Logger.log('キャッシュ更新エラー: ' + e.message);
    }

    return { success: true, message: '発送先情報を削除しました' };

  } catch (error) {
    Logger.log('deleteShippingTo error: ' + error.message);
    return { success: false, message: '削除に失敗しました: ' + error.message };
  }
}

/**
 * 発送先詳細を取得する（編集用）
 * @param {number} rowIndex - 行番号
 * @returns {Object|null} 発送先詳細データ
 */
function getShippingToDetail(rowIndex) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
      ss = SpreadsheetApp.openById(getMasterSpreadsheetId());
    }
    var sheet = ss.getSheetByName('発送先情報');

    if (!sheet) {
      return null;
    }

    var validation = normalizeAndValidateShippingToRowIndex(rowIndex, sheet);
    if (!validation.success) {
      return null;
    }
    rowIndex = validation.rowIndex;

    var headers = getShippingToSheetHeaders();
    var values = sheet.getRange(rowIndex, 1, 1, headers.length).getValues()[0];

    var shippingTo = {};
    for (var i = 0; i < headers.length; i++) {
      shippingTo[headers[i]] = values[i];
    }
    shippingTo.rowIndex = rowIndex;

    return shippingTo;

  } catch (error) {
    Logger.log('getShippingToDetail error: ' + error.message);
    return null;
  }
}
