#!/usr/bin/env node
/**
 * macOS launchd サービス登録スクリプト
 *
 * print-agent を Mac 起動時に自動起動するよう launchd に登録する。
 * npm run setup-service で実行。
 */
const fs = require('fs');
const path = require('path');
const os = require('os');
const { execFileSync } = require('child_process');

const home = process.env.HOME || os.homedir();
if (!home) {
  console.error('エラー: HOME ディレクトリを取得できません。');
  process.exit(1);
}
const LABEL = 'com.local.print-agent';
const PLIST_PATH = path.join(home, 'Library', 'LaunchAgents', `${LABEL}.plist`);
const AGENT_DIR = path.resolve(__dirname, '..');
const LOG_PATH = path.join(AGENT_DIR, 'logs', 'print-agent.log');
const NODE_PATH = process.execPath;

// config.json の存在確認
const configPath = path.join(AGENT_DIR, 'config.json');
if (!fs.existsSync(configPath)) {
  console.error('エラー: config.json が見つかりません。');
  console.error('config.example.json をコピーして config.json を作成し、watchDir を設定してください。');
  console.error('');
  console.error('  cp config.example.json config.json');
  console.error('  # config.json の watchDir を Google Drive 同期フォルダのパスに変更');
  process.exit(1);
}

// watchDir の確認
let config;
try {
  config = JSON.parse(fs.readFileSync(configPath, 'utf8'));
} catch (e) {
  console.error('エラー: config.json の読み込みに失敗しました:', e.message);
  process.exit(1);
}

if (!config.watchDir || config.watchDir.includes('/path/to/')) {
  console.error('エラー: config.json の watchDir が未設定です。');
  console.error('Google Drive for Desktop の同期フォルダのパスを設定してください。');
  console.error('');
  console.error('例: /Users/username/Library/CloudStorage/GoogleDrive-xxx/マイドライブ/伝票PDF');
  process.exit(1);
}

if (!fs.existsSync(config.watchDir)) {
  console.error(`エラー: watchDir が存在しません: ${config.watchDir}`);
  console.error('Google Drive for Desktop が起動しているか確認してください。');
  process.exit(1);
}

// ログディレクトリ作成
const logDir = path.dirname(LOG_PATH);
if (!fs.existsSync(logDir)) {
  fs.mkdirSync(logDir, { recursive: true });
}

// 既存の plist があればアンロード
if (fs.existsSync(PLIST_PATH)) {
  try {
    execFileSync('launchctl', ['unload', PLIST_PATH], { stdio: 'ignore' });
  } catch { /* ignore */ }
}

const plist = `<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
  <key>Label</key>
  <string>${LABEL}</string>
  <key>WorkingDirectory</key>
  <string>${AGENT_DIR}</string>
  <key>ProgramArguments</key>
  <array>
    <string>${NODE_PATH}</string>
    <string>index.js</string>
  </array>
  <key>RunAtLoad</key>
  <true/>
  <key>KeepAlive</key>
  <true/>
  <key>StandardOutPath</key>
  <string>${LOG_PATH}</string>
  <key>StandardErrorPath</key>
  <string>${LOG_PATH}</string>
  <key>EnvironmentVariables</key>
  <dict>
    <key>PATH</key>
    <string>${path.dirname(NODE_PATH)}:/usr/local/bin:/usr/bin:/bin</string>
  </dict>
</dict>
</plist>
`;

// LaunchAgents ディレクトリ確認
const launchAgentsDir = path.dirname(PLIST_PATH);
if (!fs.existsSync(launchAgentsDir)) {
  fs.mkdirSync(launchAgentsDir, { recursive: true });
}

fs.writeFileSync(PLIST_PATH, plist);
console.log(`plist 作成: ${PLIST_PATH}`);

// 登録
try {
  execFileSync('launchctl', ['load', PLIST_PATH]);
  console.log('launchd 登録完了');
} catch (e) {
  console.error('launchd 登録失敗:', e.message);
  process.exit(1);
}

console.log('');
console.log('=== セットアップ完了 ===');
console.log(`print-agent: ${AGENT_DIR}`);
console.log(`監視フォルダ: ${config.watchDir}`);
console.log(`プリンター:   ${config.printerName || '(デフォルト)'}`);
console.log(`Node:         ${NODE_PATH}`);
console.log(`ログ:         ${LOG_PATH}`);
console.log('');
console.log('Mac を再起動しても print-agent は自動で起動します。');
console.log('解除するには: npm run uninstall-service');
console.log('ログ確認:     npm run logs');
