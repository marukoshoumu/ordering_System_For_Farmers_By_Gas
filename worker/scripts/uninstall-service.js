#!/usr/bin/env node
/**
 * macOS launchd サービス解除スクリプト
 *
 * npm run uninstall-service で実行。
 */
const fs = require('fs');
const path = require('path');
const { execFileSync } = require('child_process');

const LABEL = 'com.local.delivery-slip-worker';
const PLIST_PATH = path.join(process.env.HOME, 'Library', 'LaunchAgents', `${LABEL}.plist`);

if (!fs.existsSync(PLIST_PATH)) {
  console.log('サービスは登録されていません。');
  process.exit(0);
}

try {
  execFileSync('launchctl', ['unload', PLIST_PATH]);
  console.log('launchd 解除完了');
} catch { /* ignore */ }

fs.unlinkSync(PLIST_PATH);
console.log(`plist 削除: ${PLIST_PATH}`);
console.log('');
console.log('Worker の自動起動を解除しました。');
