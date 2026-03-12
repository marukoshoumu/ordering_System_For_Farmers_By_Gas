# worker スクリプト

## OAuth トークン取得（get-oauth-token.js）

Google Drive API 用の OAuth 2.0 リフレッシュトークンを取得するスクリプト。

### 手順

1. GCP Console → APIとサービス → 認証情報 → OAuth 2.0 クライアントID（デスクトップアプリ）を作成
2. `worker/.env` に `GOOGLE_OAUTH_CLIENT_ID` と `GOOGLE_OAUTH_CLIENT_SECRET` を設定
3. 実行: `node scripts/get-oauth-token.js`
4. ブラウザで認証後、表示されるリフレッシュトークンを `worker/.env` の `GOOGLE_OAUTH_REFRESH_TOKEN` に設定

## 本番フロー（CSVアップロード〜PDF取得）のテスト

API 経由で実際の伝票発行を試せます。

1. `worker/.env` に認証情報を設定（`HEADLESS=false` で目視確認推奨）
2. サーバー起動: `cd worker && npm start`
3. 別ターミナルから `POST /api/print-slip` を呼ぶ（body: `carrier`, `csvContent`, `shippingDate`, `apiKey`）

### セレクタが合わない場合

- `worker/src/selectors.json` の該当サイト（yamato / sagawa）のセレクタを修正する
- ログイン画面: `loginId`, `password`, `loginButton`
- ログイン後の遷移: `entrySteps` に、クリックする順でセレクタを配列で指定
- サイトのHTMLが変わっている場合は、ブラウザの開発者ツールで要素を検証し、正しいセレクタに更新する
