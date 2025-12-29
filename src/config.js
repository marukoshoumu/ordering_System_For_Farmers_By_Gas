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
 */

// ========================================
// ScriptProperties から取得する設定
// ========================================

/**
 * スクリプトプロパティを取得
 * @param {string} key - プロパティキー
 * @returns {string} プロパティ値
 */
function getConfig(key) {
  return PropertiesService.getScriptProperties().getProperty(key) || '';
}

// === スプレッドシート ID ===
function getLineBotSpreadsheetId() {
  return getConfig('LINE_BOT_SPREADSHEET_ID');
}

function getMasterSpreadsheetId() {
  return getConfig('MASTER_SPREADSHEET_ID');
}

// === ドライブ フォルダ ID ===
function getYamatoFolderId() {
  return getConfig('YAMATO_FOLDER_ID');
}

function getSagawaFolderId() {
  return getConfig('SAGAWA_FOLDER_ID');
}

function getDeliveredPdfFolderId() {
  return getConfig('DELIVERED_PDF_FOLDER_ID');
}

function getReceiptPdfFolderId() {
  return getConfig('RECEIPT_PDF_FOLDER_ID');
}

function getBillPdfFolderId() {
  return getConfig('BILL_PDF_FOLDER_ID');
}

function getQuotationFolderId() {
  return getConfig('QUOTATION_FOLDER_ID');
}

// === テンプレート ID ===
function getDeliveredTemplateId() {
  return getConfig('DELIVERED_TEMPLATE_ID');
}

function getReceiptTemplateId() {
  return getConfig('RECEIPT_TEMPLATE_ID');
}

function getBillTemplateId() {
  return getConfig('BILL_TEMPLATE_ID');
}

// === API キー ===
function getGeminiApiKey() {
  return getConfig('GEMINI_API_KEY');
}

// === 表示用会社名 ===
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
  return getAllRecords(companyName, true);
}

// ========================================
// フォルダ URL 取得関数
// ========================================

function getQuotationFolderUrl() {
  const folderId = getQuotationFolderId();
  return folderId ? `https://drive.google.com/drive/folders/${folderId}` : '';
}

function getBillFolderUrl() {
  const folderId = getBillPdfFolderId();
  return folderId ? `https://drive.google.com/drive/folders/${folderId}` : '';
}

function getDeliveredFolderUrl() {
  const folderId = getDeliveredPdfFolderId();
  return folderId ? `https://drive.google.com/drive/folders/${folderId}` : '';
}

function getYamatoFolderUrl() {
  const folderId = getYamatoFolderId();
  return folderId ? `https://drive.google.com/drive/folders/${folderId}` : '';
}

function getSagawaFolderUrl() {
  const folderId = getSagawaFolderId();
  return folderId ? `https://drive.google.com/drive/folders/${folderId}` : '';
}
