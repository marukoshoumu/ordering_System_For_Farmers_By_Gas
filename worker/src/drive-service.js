/**
 * Google Drive サービスモジュール
 * 
 * Google Drive API v3 を使用して、ダウンロードした伝票PDFを
 * 指定フォルダへアップロードする。
 * 
 * 認証方式:
 * - サービスアカウント（JSON キーファイル）
 * - GOOGLE_SERVICE_ACCOUNT_KEY_PATH で指定
 * 
 * フォルダID:
 * - DRIVE_YAMATO_FOLDER_ID: ヤマト伝票PDF保存先
 * - DRIVE_SAGAWA_FOLDER_ID: 佐川伝票PDF保存先
 */

const { google } = require('googleapis');
const fs = require('fs');
const path = require('path');

/**
 * Google Drive API クライアント（遅延初期化）
 */
let _driveClient = null;

/**
 * Drive API クライアントを取得・初期化
 * 
 * @returns {google.drive_v3.Drive}
 */
function getDriveClient() {
    if (_driveClient) return _driveClient;

    const keyPath = process.env.GOOGLE_SERVICE_ACCOUNT_KEY_PATH;
    if (!keyPath) {
        throw new Error('GOOGLE_SERVICE_ACCOUNT_KEY_PATH が設定されていません');
    }

    if (!fs.existsSync(keyPath)) {
        throw new Error(`サービスアカウントキーファイルが見つかりません: ${keyPath}`);
    }

    const keyFile = JSON.parse(fs.readFileSync(keyPath, 'utf-8'));

    const auth = new google.auth.GoogleAuth({
        credentials: keyFile,
        scopes: ['https://www.googleapis.com/auth/drive.file'],
    });

    _driveClient = google.drive({ version: 'v3', auth });
    console.log('Google Drive API クライアント初期化完了');

    return _driveClient;
}

/**
 * 配送業者に対応するDrive保存先フォルダIDを取得
 * 
 * @param {string} carrier - 'yamato' | 'sagawa'
 * @returns {string} フォルダID
 */
function getFolderIdForCarrier(carrier) {
    const folderId = carrier === 'yamato'
        ? process.env.DRIVE_YAMATO_FOLDER_ID
        : process.env.DRIVE_SAGAWA_FOLDER_ID;

    if (!folderId) {
        throw new Error(
            `${carrier === 'yamato' ? 'DRIVE_YAMATO_FOLDER_ID' : 'DRIVE_SAGAWA_FOLDER_ID'} が設定されていません`
        );
    }

    return folderId;
}

/**
 * PDFファイルをGoogle Driveにアップロード
 * 
 * @param {string} pdfPath - アップロードするPDFのローカルパス
 * @param {string} carrier - 配送業者 ('yamato' | 'sagawa')
 * @param {string} shippingDate - 発送日（ファイル名に使用）
 * @returns {Promise<{ fileId: string, webViewLink: string }>}
 */
async function uploadToDrive(pdfPath, carrier, shippingDate, jobId) {
    const drive = getDriveClient();
    const folderId = getFolderIdForCarrier(carrier);
    const fileName = path.basename(pdfPath);

    console.log('Drive アップロード開始', { fileName, carrier, folderId });

    try {
        const fileMetadata = {
            name: fileName,
            parents: [folderId],
            description: JSON.stringify({
                status: 'pending',
                jobId: jobId || 'unknown',
                carrier,
                shippingDate,
                createdAt: new Date().toISOString(),
            }),
        };

        const media = {
            mimeType: 'application/pdf',
            body: fs.createReadStream(pdfPath),
        };

        const response = await drive.files.create({
            resource: fileMetadata,
            media: media,
            fields: 'id, name, webViewLink, webContentLink',
        });

        const fileId = response.data.id;
        const webViewLink = response.data.webViewLink || '';

        console.log('Drive アップロード完了', {
            fileId,
            fileName: response.data.name,
            webViewLink,
        });

        return { fileId, webViewLink };
    } catch (error) {
        console.error('Drive アップロード失敗', {
            error: error.message,
            carrier,
            fileName,
        });
        throw new Error(`Google Driveアップロードエラー: ${error.message}`);
    }
}

/**
 * Google Drive 設定の検証
 * 
 * @returns {{ configured: boolean, message: string }}
 */
function validateDriveConfig() {
    const keyPath = process.env.GOOGLE_SERVICE_ACCOUNT_KEY_PATH;
    const yamatoFolder = process.env.DRIVE_YAMATO_FOLDER_ID;
    const sagawaFolder = process.env.DRIVE_SAGAWA_FOLDER_ID;

    const missing = [];
    if (!keyPath) missing.push('GOOGLE_SERVICE_ACCOUNT_KEY_PATH');
    if (!yamatoFolder) missing.push('DRIVE_YAMATO_FOLDER_ID');
    if (!sagawaFolder) missing.push('DRIVE_SAGAWA_FOLDER_ID');

    if (missing.length > 0) {
        return {
            configured: false,
            message: `未設定: ${missing.join(', ')}`,
        };
    }

    if (!fs.existsSync(keyPath)) {
        return {
            configured: false,
            message: `キーファイルが存在しません: ${keyPath}`,
        };
    }

    return { configured: true, message: 'OK' };
}

module.exports = { uploadToDrive, validateDriveConfig };
