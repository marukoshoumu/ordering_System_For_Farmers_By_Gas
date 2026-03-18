#!/usr/bin/env node
/**
 * サービス登録スクリプト（macOS launchd / Windows タスクスケジューラ）
 *
 * print-agent を OS 起動時に自動起動するよう登録する。
 * npm run setup-service で実行。
 */
const fs = require('fs');
const path = require('path');
const os = require('os');
const { execFileSync } = require('child_process');

const IS_WIN = os.platform() === 'win32';
const AGENT_DIR = path.resolve(__dirname, '..');
const LOG_PATH = path.join(AGENT_DIR, 'logs', 'print-agent.log');
const NODE_PATH = process.execPath;

// ── config.json の検証 ──

const configPath = path.join(AGENT_DIR, 'config.json');
if (!fs.existsSync(configPath)) {
  console.error('エラー: config.json が見つかりません。');
  console.error('config.example.json をコピーして config.json を作成し、watchDirs を設定してください。');
  console.error('');
  console.error('  cp config.example.json config.json');
  console.error('  # config.json の watchDirs を Google Drive 同期フォルダのパスに変更');
  process.exit(1);
}

let config;
try {
  config = JSON.parse(fs.readFileSync(configPath, 'utf8'));
} catch (e) {
  console.error('エラー: config.json の読み込みに失敗しました:', e.message);
  process.exit(1);
}

// watchDirs の正規化（文字列/オブジェクト混在対応）
const rawDirs = config.watchDirs || (config.watchDir ? [config.watchDir] : []);
const arr = Array.isArray(rawDirs) ? rawDirs : [rawDirs];
const watchDirs = arr
  .map((item) => (typeof item === 'string' ? { path: item } : item))
  .filter((item) => item && item.path);

if (watchDirs.length === 0 || watchDirs.some(d => d.path.includes('/path/to/'))) {
  console.error('エラー: config.json の watchDirs が未設定です。');
  console.error('Google Drive for Desktop の同期フォルダのパスを設定してください。');
  console.error('');
  console.error('例:');
  console.error('  "watchDirs": [');
  console.error('    { "path": "/Users/.../ヤマト伝票PDF", "printerName": "EPSON" },');
  console.error('    { "path": "/Users/.../佐川伝票PDF", "printerName": "Brother" }');
  console.error('  ]');
  process.exit(1);
}

for (const entry of watchDirs) {
  if (!fs.existsSync(entry.path)) {
    console.error(`エラー: watchDir が存在しません: ${entry.path}`);
    console.error('Google Drive for Desktop が起動しているか確認してください。');
    process.exit(1);
  }
}

// ── ログディレクトリ作成 ──

const logDir = path.dirname(LOG_PATH);
if (!fs.existsSync(logDir)) {
  fs.mkdirSync(logDir, { recursive: true });
}

// ── OS 別サービス登録 ──

if (IS_WIN) {
  setupWindows();
} else {
  setupMacOS();
}

// ── 完了表示 ──

console.log('');
console.log('=== セットアップ完了 ===');
console.log(`print-agent: ${AGENT_DIR}`);
for (const entry of watchDirs) {
  const printer = entry.printerName ? ` → printer: ${entry.printerName}` : '';
  console.log(`監視フォルダ: ${entry.path}${printer}`);
}
console.log(`プリンター:   ${config.printerName || '(デフォルト / watchDirs個別設定)'}`);
console.log(`Node:         ${NODE_PATH}`);
console.log(`ログ:         ${LOG_PATH}`);
console.log('');
console.log('OS 再起動後も print-agent は自動で起動します。');
console.log('解除するには: npm run uninstall-service');
console.log('ログ確認:     npm run logs');

// ────────────────────────────────────────────

function setupMacOS() {
  const home = process.env.HOME || os.homedir();
  if (!home) {
    console.error('エラー: HOME ディレクトリを取得できません。');
    process.exit(1);
  }
  const LABEL = 'com.local.print-agent';
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
  const TASK_NAME = 'PrintAgent';

  // 既存タスクを削除
  try {
    execFileSync('schtasks', ['/Delete', '/TN', TASK_NAME, '/F'], { stdio: 'ignore' });
  } catch { /* ignore */ }

  // ログ出力付きで起動するバッチファイルを作成
  const batPath = path.join(AGENT_DIR, 'start-service.bat');
  const batContent = `@echo off\r\n"${NODE_PATH}" "${path.join(AGENT_DIR, 'index.js')}" >> "${LOG_PATH}" 2>&1\r\n`;
  fs.writeFileSync(batPath, batContent);

  // VBSラッパーでコンソールウィンドウを非表示にして起動
  const vbsPath = path.join(AGENT_DIR, 'start-service-silent.vbs');
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
    console.log('print-agent を起動しました');
  } catch (e) {
    console.error('起動失敗（手動で npm start を実行してください）:', e.message);
  }
}
