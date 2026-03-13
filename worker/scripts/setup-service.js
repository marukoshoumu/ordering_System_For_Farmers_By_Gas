#!/usr/bin/env node
/**
 * macOS launchd サービス登録スクリプト
 *
 * Worker を Mac 起動時に自動起動するよう launchd に登録する。
 * npm run setup-service で実行。
 */
const fs = require('fs');
const path = require('path');
const { execFileSync } = require('child_process');

const LABEL = 'com.local.delivery-slip-worker';
const PLIST_PATH = path.join(process.env.HOME, 'Library', 'LaunchAgents', `${LABEL}.plist`);
const WORKER_DIR = path.resolve(__dirname, '..');
const LOG_PATH = path.join(WORKER_DIR, 'logs', 'worker.log');
const NODE_PATH = process.execPath;

// .env の存在確認
const envPath = path.join(WORKER_DIR, '.env');
if (!fs.existsSync(envPath)) {
  console.error('エラー: .env ファイルが見つかりません。先に .env を作成してください。');
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
  <string>${WORKER_DIR}</string>
  <key>ProgramArguments</key>
  <array>
    <string>${NODE_PATH}</string>
    <string>src/server.js</string>
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
console.log(`Worker: ${WORKER_DIR}`);
console.log(`Node:   ${NODE_PATH}`);
console.log(`ログ:   ${LOG_PATH}`);
console.log('');
console.log('Mac を再起動しても Worker は自動で起動します。');
console.log('解除するには: npm run uninstall-service');
console.log('ログ確認:     npm run logs');
