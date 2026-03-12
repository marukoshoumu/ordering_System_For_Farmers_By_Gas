#!/usr/bin/env node
/**
 * Google Drive OAuth 2.0 リフレッシュトークン取得スクリプト
 *
 * 使い方:
 * 1. GCP Console → APIとサービス → 認証情報 → OAuth 2.0 クライアントID作成（デスクトップアプリ）
 * 2. クライアントIDとシークレットを .env に設定:
 *    GOOGLE_OAUTH_CLIENT_ID=xxxxx.apps.googleusercontent.com
 *    GOOGLE_OAUTH_CLIENT_SECRET=xxxxx
 * 3. このスクリプトを実行: node scripts/get-oauth-token.js
 * 4. ブラウザで認証後、表示されるリフレッシュトークンを .env に追加:
 *    GOOGLE_OAUTH_REFRESH_TOKEN=xxxxx
 */

require('dotenv').config();
const { google } = require('googleapis');
const http = require('http');
const url = require('url');

const CLIENT_ID = process.env.GOOGLE_OAUTH_CLIENT_ID;
const CLIENT_SECRET = process.env.GOOGLE_OAUTH_CLIENT_SECRET;
const REDIRECT_PORT = 3001;
const REDIRECT_URI = `http://localhost:${REDIRECT_PORT}/callback`;

if (!CLIENT_ID || !CLIENT_SECRET) {
    console.error('エラー: .env に以下を設定してください:');
    console.error('  GOOGLE_OAUTH_CLIENT_ID=xxxxx');
    console.error('  GOOGLE_OAUTH_CLIENT_SECRET=xxxxx');
    console.error('');
    console.error('GCP Console → APIとサービス → 認証情報 → OAuth 2.0 クライアントID作成（デスクトップアプリ）');
    process.exit(1);
}

const oauth2Client = new google.auth.OAuth2(CLIENT_ID, CLIENT_SECRET, REDIRECT_URI);

const authUrl = oauth2Client.generateAuthUrl({
    access_type: 'offline',
    prompt: 'consent',
    scope: ['https://www.googleapis.com/auth/drive.file'],
});

const server = http.createServer(async (req, res) => {
    const parsed = url.parse(req.url, true);
    if (parsed.pathname !== '/callback') {
        res.writeHead(404);
        res.end('Not Found');
        return;
    }

    const code = parsed.query.code;
    if (!code) {
        res.writeHead(400);
        res.end('認証コードがありません');
        return;
    }

    try {
        const { tokens } = await oauth2Client.getToken(code);
        const refreshToken = tokens.refresh_token;

        res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' });
        res.end(`
            <h1>認証成功!</h1>
            <p>以下のリフレッシュトークンを <code>.env</code> に追加してください:</p>
            <pre style="background:#f0f0f0;padding:10px;word-break:break-all;">GOOGLE_OAUTH_REFRESH_TOKEN=${refreshToken}</pre>
            <p>このページを閉じてください。</p>
        `);

        console.log('');
        console.log('=== 認証成功 ===');
        console.log('');
        console.log('.env に以下を追加してください:');
        console.log(`GOOGLE_OAUTH_REFRESH_TOKEN=${refreshToken}`);
        console.log('');

        server.close();
        process.exit(0);
    } catch (err) {
        res.writeHead(500);
        res.end('トークン取得エラー: ' + err.message);
        console.error('トークン取得エラー:', err.message);
    }
});

server.listen(REDIRECT_PORT, () => {
    console.log('');
    console.log('=== Google Drive OAuth 認証 ===');
    console.log('');
    console.log('以下のURLをブラウザで開いてください:');
    console.log('');
    console.log(authUrl);
    console.log('');
    console.log(`認証後、http://localhost:${REDIRECT_PORT}/callback にリダイレクトされます。`);
});
