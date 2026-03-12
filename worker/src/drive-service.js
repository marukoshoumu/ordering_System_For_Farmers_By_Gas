/**
 * Google Drive サービスモジュール
 *
 * Google Drive API v3 を使用して、ダウンロードした伝票PDFを
 * 指定フォルダへアップロードする。
 *
 * 認証方式（優先順）:
 * 1. OAuth 2.0（GOOGLE_OAUTH_CLIENT_ID + SECRET + REFRESH_TOKEN）
 * 2. サービスアカウント（GOOGLE_SERVICE_ACCOUNT_KEY_PATH）
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
 * Drive API クライアントをリセット（次回 getDriveClient() で再初期化される）
 * テストや認証情報変更後に使用する。
 */
function resetDriveClient() {
    _driveClient = null;
}

/**
 * Drive API クライアントを外部から設定（テスト用モックや事前設定クライアント）
 * @param {google.drive_v3.Drive|null} client - 設定するクライアント（null でリセットと同様）
 */
function setDriveClient(client) {
    _driveClient = client;
}

/**
 * Drive API クライアントを取得・初期化
 *
 * @returns {google.drive_v3.Drive}
 */
function getDriveClient() {
    if (_driveClient) return _driveClient;

    const clientId = process.env.GOOGLE_OAUTH_CLIENT_ID;
    const clientSecret = process.env.GOOGLE_OAUTH_CLIENT_SECRET;
    const refreshToken = process.env.GOOGLE_OAUTH_REFRESH_TOKEN;

    if (clientId && clientSecret && refreshToken) {
        // OAuth 2.0 方式（ユーザーアカウントのストレージを使用）
        const oauth2Client = new google.auth.OAuth2(clientId, clientSecret);
        oauth2Client.setCredentials({ refresh_token: refreshToken });
        _driveClient = google.drive({ version: 'v3', auth: oauth2Client });
        console.log('Google Drive API クライアント初期化完了（OAuth 2.0）');
        return _driveClient;
    }

    // サービスアカウント方式（共有ドライブ専用）
    const keyPath = process.env.GOOGLE_SERVICE_ACCOUNT_KEY_PATH;
    if (!keyPath) {
        throw new Error('Drive認証が未設定です。OAuth（GOOGLE_OAUTH_CLIENT_ID/SECRET/REFRESH_TOKEN）またはサービスアカウント（GOOGLE_SERVICE_ACCOUNT_KEY_PATH）を設定してください');
    }

    if (!fs.existsSync(keyPath)) {
        throw new Error(`サービスアカウントキーファイルが見つかりません: ${keyPath}`);
    }

    let keyFile;
    try {
        keyFile = JSON.parse(fs.readFileSync(keyPath, 'utf-8'));
    } catch (err) {
        const msg = `サービスアカウントキーファイルの読み込みに失敗しました: ${keyPath} — ${err.message}`;
        console.error(msg);
        throw new Error(msg);
    }

    const auth = new google.auth.GoogleAuth({
        credentials: keyFile,
        scopes: ['https://www.googleapis.com/auth/drive.file'],
    });

    _driveClient = google.drive({ version: 'v3', auth });
    console.log('Google Drive API クライアント初期化完了（サービスアカウント）');

    return _driveClient;
}

/**
 * 配送業者に対応するDrive保存先フォルダIDを取得
 * 
 * @param {string} carrier - 'yamato' | 'sagawa'
 * @returns {string} フォルダID
 */
function getFolderIdForCarrier(carrier) {
    const normalized = typeof carrier === 'string' ? carrier.trim().toLowerCase() : '';
    const allowed = ['yamato', 'sagawa'];
    if (!allowed.includes(normalized)) {
        throw new Error(
            `無効な配送業者です: "${carrier}"。許可値: ${allowed.join(', ')}`
        );
    }

    const folderId = normalized === 'yamato'
        ? process.env.DRIVE_YAMATO_FOLDER_ID
        : process.env.DRIVE_SAGAWA_FOLDER_ID;

    if (!folderId) {
        throw new Error(
            `${normalized === 'yamato' ? 'DRIVE_YAMATO_FOLDER_ID' : 'DRIVE_SAGAWA_FOLDER_ID'} が設定されていません`
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
 * @param {string} jobId - ジョブ識別子（アップロード先のメタデータ等に使用）
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
            supportsAllDrives: true,
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
            pdfPath,
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
    const yamatoFolder = process.env.DRIVE_YAMATO_FOLDER_ID;
    const sagawaFolder = process.env.DRIVE_SAGAWA_FOLDER_ID;

    const missing = [];
    if (!yamatoFolder) missing.push('DRIVE_YAMATO_FOLDER_ID');
    if (!sagawaFolder) missing.push('DRIVE_SAGAWA_FOLDER_ID');

    // OAuth方式チェック
    const hasOAuth = process.env.GOOGLE_OAUTH_CLIENT_ID
        && process.env.GOOGLE_OAUTH_CLIENT_SECRET
        && process.env.GOOGLE_OAUTH_REFRESH_TOKEN;

    // サービスアカウント方式チェック
    const keyPath = process.env.GOOGLE_SERVICE_ACCOUNT_KEY_PATH;
    const hasServiceAccount = keyPath && fs.existsSync(keyPath);

    if (!hasOAuth && !hasServiceAccount) {
        missing.push('認証設定（OAuth or サービスアカウント）');
    }

    if (missing.length > 0) {
        return {
            configured: false,
            message: `未設定: ${missing.join(', ')}`,
        };
    }

    const authMethod = hasOAuth ? 'OAuth 2.0' : 'サービスアカウント';
    return { configured: true, message: `OK (${authMethod})` };
}

module.exports = {
    uploadToDrive,
    validateDriveConfig,
    getDriveClient,
    resetDriveClient,
    setDriveClient,
};
