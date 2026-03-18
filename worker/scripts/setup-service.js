#!/usr/bin/env node
/**
 * サービス登録スクリプト（macOS launchd / Windows タスクスケジューラ）
 *
 * Worker を OS 起動時に自動起動するよう登録する。
 * npm run setup-service で実行。
 */
const fs = require('fs');
const path = require('path');
const os = require('os');
const { execFileSync } = require('child_process');

const IS_WIN = os.platform() === 'win32';
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

// OS 別サービス登録
if (IS_WIN) {
  setupWindows();
} else {
  setupMacOS();
}

console.log('');
console.log('=== セットアップ完了 ===');
console.log(`Worker: ${WORKER_DIR}`);
console.log(`Node:   ${NODE_PATH}`);
console.log(`ログ:   ${LOG_PATH}`);
console.log('');
console.log('OS 再起動後も Worker は自動で起動します。');
console.log('解除するには: npm run uninstall-service');
console.log('ログ確認:     npm run logs');

// ────────────────────────────────────────────

function setupMacOS() {
  const home = process.env.HOME || os.homedir();
  if (!home) {
    console.error('エラー: HOME ディレクトリを取得できません。');
    process.exit(1);
  }
  const LABEL = 'com.local.delivery-slip-worker';
  const PLIST_PATH = path.join(home, 'Library', 'LaunchAgents', `${LABEL}.plist`);

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

  const launchAgentsDir = path.dirname(PLIST_PATH);
  if (!fs.existsSync(launchAgentsDir)) {
    fs.mkdirSync(launchAgentsDir, { recursive: true });
  }

  fs.writeFileSync(PLIST_PATH, plist);
  console.log(`plist 作成: ${PLIST_PATH}`);

  try {
    execFileSync('launchctl', ['load', PLIST_PATH]);
    console.log('launchd 登録完了');
  } catch (e) {
    console.error('launchd 登録失敗:', e.message);
    process.exit(1);
  }
}

function setupWindows() {
  const TASK_NAME = 'DeliverySlipWorker';

  // 既存タスクを削除
  try {
    execFileSync('schtasks', ['/Delete', '/TN', TASK_NAME, '/F'], { stdio: 'ignore' });
  } catch { /* ignore */ }

  // ログ出力付きで起動するバッチファイルを作成
  const batPath = path.join(WORKER_DIR, 'start-service.bat');
  const batContent = `@echo off\r\ncd /d "${WORKER_DIR}"\r\n"${NODE_PATH}" "src\\server.js" >> "${LOG_PATH}" 2>&1\r\n`;
  fs.writeFileSync(batPath, batContent);

  // VBSラッパーでコンソールウィンドウを非表示にして起動
  const vbsPath = path.join(WORKER_DIR, 'start-service-silent.vbs');
  const vbsContent = `Set WshShell = CreateObject("WScript.Shell")\r\nWshShell.Run """${batPath}""", 0, False\r\n`;
  fs.writeFileSync(vbsPath, vbsContent);

  try {
    execFileSync('schtasks', [
      '/Create',
      '/TN', TASK_NAME,
      '/TR', `wscript.exe "${vbsPath}"`,
      '/SC', 'ONLOGON',
      '/RL', 'HIGHEST',
      '/F',
    ]);
    console.log('タスクスケジューラ登録完了');
  } catch (e) {
    console.error('タスクスケジューラ登録失敗:', e.message);
    process.exit(1);
  }

  // 即時起動
  try {
    execFileSync('schtasks', ['/Run', '/TN', TASK_NAME]);
    console.log('Worker を起動しました');
  } catch (e) {
    console.error('起動失敗（手動で npm start を実行してください）:', e.message);
  }
}
