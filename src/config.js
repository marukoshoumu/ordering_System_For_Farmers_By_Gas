/**
 * 設定ファイル - 環境変数・定数管理
 * 
 * 機密情報は GAS の PropertiesService.getScriptProperties() で管理します。
 * ローカル開発時は .env ファイルを参照してください。
 * 
 * 【設定手順】
 * 1. GAS エディタを開く
 * 2. プロジェクトの設定 > スクリプトプロパティ
 * 3. 以下のキーと値を追加
 * 
 * 【パフォーマンス最適化】
 * - CacheServiceで6時間キャッシュ（複数リクエスト間で共有）
 * - メモリキャッシュ併用で同一リクエスト内は即座に返却
 */

// ========================================
// キャッシュ機構
// ========================================

// キャッシュキー
const CONFIG_CACHE_KEY = 'SCRIPT_PROPERTIES_CACHE';
// キャッシュ有効期間（秒）: 6時間
const CONFIG_CACHE_DURATION = 21600;

// メモリキャッシュ（同一リクエスト内で高速化）
let _configCache = null;

/**
 * スクリプトプロパティを一括取得してキャッシュ
 * - 同一リクエスト内: メモリキャッシュから取得
 * - 別リクエスト: CacheServiceから取得（6時間有効）
 * - キャッシュなし: PropertiesServiceから取得してキャッシュ
 * @returns {Object} 全プロパティのオブジェクト
 */
function getConfigAll() {
  // 1. メモリキャッシュをチェック（同一リクエスト内で最速）
  if (_configCache !== null) {
    return _configCache;
  }

  // 2. CacheServiceをチェック（複数リクエスト間で共有）
  const cache = CacheService.getScriptCache();
  const cachedData = cache.get(CONFIG_CACHE_KEY);

  if (cachedData) {
    try {
      _configCache = JSON.parse(cachedData);
      return _configCache;
    } catch (e) {
      // パースエラーの場合は再取得
    }
  }

  // 3. PropertiesServiceから取得してキャッシュ
  _configCache = PropertiesService.getScriptProperties().getProperties();

  // CacheServiceに保存（6時間有効）
  try {
    cache.put(CONFIG_CACHE_KEY, JSON.stringify(_configCache), CONFIG_CACHE_DURATION);
  } catch (e) {
    // キャッシュ保存エラーは無視（動作には影響なし）
  }

  return _configCache;
}

/**
 * スクリプトプロパティを取得（キャッシュ利用）
 * @param {string} key - プロパティキー
 * @returns {string} プロパティ値
 */
function getConfig(key) {
  const props = getConfigAll();
  return props[key] || '';
}

/**
 * 設定キャッシュをクリア
 * スクリプトプロパティを変更した後に呼び出す
 */
function clearConfigCache() {
  _configCache = null;
  const cache = CacheService.getScriptCache();
  cache.remove(CONFIG_CACHE_KEY);
}

/**
 * LINE Bot 用スプレッドシート ID を取得
 * @returns {string} スプレッドシート ID
 */
function getLineBotSpreadsheetId() {
  return getConfig('LINE_BOT_SPREADSHEET_ID');
}

/**
 * マスタスプレッドシート ID を取得
 * @returns {string} スプレッドシート ID
 */
function getMasterSpreadsheetId() {
  return getConfig('MASTER_SPREADSHEET_ID');
}

/**
 * ヤマトCSV出力用フォルダ ID を取得
 * @returns {string} フォルダ ID
 */
function getYamatoFolderId() {
  return getConfig('YAMATO_FOLDER_ID');
}

/**
 * 佐川CSV出力用フォルダ ID を取得
 * @returns {string} フォルダ ID
 */
function getSagawaFolderId() {
  return getConfig('SAGAWA_FOLDER_ID');
}

/**
 * 納品書PDF出力用フォルダ ID を取得
 * @returns {string} フォルダ ID
 */
function getDeliveredPdfFolderId() {
  return getConfig('DELIVERED_PDF_FOLDER_ID');
}

/**
 * 領収書PDF出力用フォルダ ID を取得
 * @returns {string} フォルダ ID
 */
function getReceiptPdfFolderId() {
  return getConfig('RECEIPT_PDF_FOLDER_ID');
}

/**
 * 請求書PDF出力用フォルダ ID を取得
 * @returns {string} フォルダ ID
 */
function getBillPdfFolderId() {
  return getConfig('BILL_PDF_FOLDER_ID');
}

/**
 * 見積書フォルダ ID を取得
 * @returns {string} フォルダ ID
 */
function getQuotationFolderId() {
  return getConfig('QUOTATION_FOLDER_ID');
}
/**
 * FreeeCSV出力用フォルダ ID を取得
 * @returns {string} フォルダ ID
 */
function getFreeeCSVFolderId() {
  return getConfig('FREEE_CSV_FOLDER_ID');
}
/**
 * 納品書テンプレートドキュメント ID を取得
 * @returns {string} ドキュメント ID
 */
function getDeliveredTemplateId() {
  return getConfig('DELIVERED_TEMPLATE_ID');
}

/**
 * 領収書テンプレートドキュメント ID を取得
 * @returns {string} ドキュメント ID
 */
function getReceiptTemplateId() {
  return getConfig('RECEIPT_TEMPLATE_ID');
}

/**
 * 請求書テンプレートドキュメント ID を取得
 * @returns {string} ドキュメント ID
 */
function getBillTemplateId() {
  return getConfig('BILL_TEMPLATE_ID');
}

/**
 * Gemini API キーを取得
 * @returns {string} API キー
 */
function getGeminiApiKey() {
  return getConfig('GEMINI_API_KEY');
}

/**
 * Cloud Vision API キーを取得（未設定時は Gemini API キーを流用）
 * @returns {string} API キー
 */
function getVisionApiKey() {
  return getConfig('VISION_API_KEY') || getConfig('GEMINI_API_KEY');
}

/**
 * アプリケーションの表示用会社名を取得
 * @returns {string} 会社名
 */
function getCompanyDisplayName() {
  return getConfig('COMPANY_DISPLAY_NAME') || '受注管理システム';
}

// ========================================
// 会社情報（定数）
// ========================================

const CONFIG = {
  // 会社情報
  COMPANY: {
    NAME: getConfig('COMPANY_NAME') || '株式会社○○○○',
    ZIPCODE: getConfig('COMPANY_ZIPCODE') || '000-0000',
    ADDRESS: getConfig('COMPANY_ADDRESS') || '○○県○○市○○町0-0-0',
    TEL: getConfig('COMPANY_TEL') || '0000-00-0000'
  },

  // 外部連携先
  TABECHOKU: {
    COMPANY: getConfig('TABECHOKU_COMPANY') || '',
    ZIPCODE: getConfig('TABECHOKU_ZIPCODE') || '000-0000',
    ADDRESS: getConfig('TABECHOKU_ADDRESS') || ''
  },

  POCKEMARU: {
    COMPANY: getConfig('POCKEMARU_COMPANY') || '',
    ZIPCODE: getConfig('POCKEMARU_ZIPCODE') || '000-0000',
    ADDRESS: getConfig('POCKEMARU_ADDRESS') || ''
  },

  FURUSATO: {
    COMPANY: getConfig('FURUSATO_COMPANY') || '',
    ZIPCODE: getConfig('FURUSATO_ZIPCODE') || '000-0000',
    ADDRESS: getConfig('FURUSATO_ADDRESS') || '',
    TEL: getConfig('FURUSATO_TEL') || ''
  }
};

// ========================================
// ヘルパー関数
// ========================================

/**
 * 自社情報を顧客マスタから取得（発送元情報用）
 * COMPANY_DISPLAY_NAME と同じ名前の顧客レコードを取得
 */
function getCompanyRecordForShipping() {
  const companyName = getCompanyDisplayName();
  return getAllRecordsInternal(companyName, true);
}

// ========================================
// フォルダ URL 取得関数
// ========================================

/**
 * 見積書フォルダの URL を取得
 * @returns {string} フォルダ URL
 */
function getQuotationFolderUrl() {
  const folderId = getQuotationFolderId();
  return folderId ? `https://drive.google.com/drive/folders/${folderId}` : '';
}

/**
 * 請求書フォルダの URL を取得
 * @returns {string} フォルダ URL
 */
function getBillFolderUrl() {
  const folderId = getBillPdfFolderId();
  return folderId ? `https://drive.google.com/drive/folders/${folderId}` : '';
}

/**
 * 納品書フォルダの URL を取得
 * @returns {string} フォルダ URL
 */
function getDeliveredFolderUrl() {
  const folderId = getDeliveredPdfFolderId();
  return folderId ? `https://drive.google.com/drive/folders/${folderId}` : '';
}

/**
 * ヤマトCSVフォルダの URL を取得
 * @returns {string} フォルダ URL
 */
function getYamatoFolderUrl() {
  const folderId = getYamatoFolderId();
  return folderId ? `https://drive.google.com/drive/folders/${folderId}` : '';
}

/**
 * 佐川CSVフォルダの URL を取得
 * @returns {string} フォルダ URL
 */
function getSagawaFolderUrl() {
  const folderId = getSagawaFolderId();
  return folderId ? `https://drive.google.com/drive/folders/${folderId}` : '';
}

/**
 * freee納品書CSVフォルダの URL を取得
 * @returns {string} フォルダ URL
 */
function getFreeeCSVFolderUrl() {
  const folderId = getFreeeCSVFolderId();
  return folderId ? `https://drive.google.com/drive/folders/${folderId}` : '';
}
