#!/usr/bin/env node
/**
 * サービス解除スクリプト（macOS launchd / Windows タスクスケジューラ）
 *
 * npm run uninstall-service で実行。
 */
const fs = require('fs');
const path = require('path');
const os = require('os');
const { execFileSync } = require('child_process');

const IS_WIN = os.platform() === 'win32';

if (IS_WIN) {
  uninstallWindows();
} else {
  uninstallMacOS();
}

function uninstallMacOS() {
  const home = process.env.HOME || os.homedir();
  if (!home) {
    console.error('エラー: HOME ディレクトリを取得できません。');
    process.exit(1);
  }
  const LABEL = 'com.local.delivery-slip-worker';
  const PLIST_PATH = path.join(home, 'Library', 'LaunchAgents', `${LABEL}.plist`);

  if (!fs.existsSync(PLIST_PATH)) {
    console.log('サービスは登録されていません。');
    process.exit(0);
  }

  try {
    execFileSync('launchctl', ['unload', PLIST_PATH]);
    console.log('launchd 解除完了');
  } catch { /* ignore */ }

  try {
    fs.unlinkSync(PLIST_PATH);
    console.log(`plist 削除: ${PLIST_PATH}`);
  } catch (e) {
    if (e.code !== 'ENOENT') {
      console.error(`plist 削除に失敗: ${PLIST_PATH}`, e.message);
      process.exit(1);
    }
  }
  console.log('');
  console.log('Worker の自動起動を解除しました。');
}

function uninstallWindows() {
  const TASK_NAME = 'DeliverySlipWorker';

  try {
    execFileSync('schtasks', ['/Delete', '/TN', TASK_NAME, '/F']);
    console.log('タスクスケジューラから削除完了');
  } catch (e) {
    if (e.message && e.message.includes('not exist')) {
      console.log('サービスは登録されていません。');
      process.exit(0);
    }
    console.error('タスク削除失敗:', e.message);
    process.exit(1);
  }

  // バッチファイルも削除
  const batPath = path.join(path.resolve(__dirname, '..'), 'start-service.bat');
  try {
    fs.unlinkSync(batPath);
    console.log(`バッチファイル削除: ${batPath}`);
  } catch (e) {
    if (e.code !== 'ENOENT') {
      console.error(`バッチファイル削除に失敗: ${batPath}`, e.message);
    }
  }

  console.log('');
  console.log('Worker の自動起動を解除しました。');
}
