# print-agent

現場 PC で Google Drive 同期フォルダを監視し、新規 PDF を検知したらデフォルトプリンタで印刷し、成功時に `printed/` へ移動する Node.js 常駐プログラム。

## セットアップ

1. `npm install`
2. `config.example.json` を `config.json` にコピーし、`watchDir` を **Google Drive のローカル同期フォルダのパス** に編集する
3. （任意）`printerName` でプリンタを指定。空の場合はシステムのデフォルトプリンタを使用
4. `npm start`

## 設定

### config.json

| キー | 説明 |
|------|------|
| `watchDir` | 監視するフォルダの絶対パス（必須）。Drive 同期フォルダを指定する |
| `printedDirName` | 印刷後の移動先フォルダ名（既定: `printed`） |
| `printerName` | プリンタ名。空ならデフォルトプリンタ |
| `debounceMs` | 同一ファイルの再検知を無視するミリ秒（既定: 60000） |
| `logDir` | ログ出力先ディレクトリ（既定: `./logs`） |

### 環境変数

上記は環境変数でも上書き可能です。

- `WATCH_DIR` / `PRINTED_DIR_NAME` / `PRINTER_NAME` / `DEBOUNCE_MS` / `LOG_DIR`

## ログ

- **保存場所:** 実行ディレクトリ直下の `./logs/`（`logDir` で変更可）
- **ファイル名:** `print-agent-YYYY-MM-DD.log`（日付別）
- 同じ内容が標準出力にも出るため、launchd の `StandardOutPath` でファイルに残せる

## 常駐起動（macOS launchd）

`launchd/com.local.print-agent.plist.example` を参考に、`~/Library/LaunchAgents/` に plist を配置し、`launchctl load` で有効化する。WorkingDirectory を print-agent のパスに、ProgramArguments に `node` のフルパスと `index.js` を指定する。

## トラブルシュート

- **起動直後に終了する:** 監視フォルダ（`watchDir`）が存在しない、または空の設定の可能性。`config.json` の `watchDir` が正しい絶対パスか、Drive 同期が有効でそのフォルダが存在するか確認する。
- **印刷されない:** デフォルトプリンタが Brother になっているか、`printerName` で正しいプリンタ名を指定しているか確認。ターミナルで `lp -d "プリンタ名" ファイル.pdf` が動くか試す。
- **重複印刷:** `debounceMs` を大きくする（例: 120000）。
