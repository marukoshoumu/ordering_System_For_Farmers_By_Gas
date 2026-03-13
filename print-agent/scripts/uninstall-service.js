#!/usr/bin/env node
/**
 * macOS launchd サービス解除スクリプト
 *
 * npm run uninstall-service で実行。
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
console.log('print-agent の自動起動を解除しました。');
