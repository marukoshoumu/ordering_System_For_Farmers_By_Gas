const fs = require('fs');
const path = require('path');

const MAX_RETRIES = 5;
const RETRY_DELAY_MS = 2000;

function sleepSync(ms) {
  Atomics.wait(new Int32Array(new SharedArrayBuffer(4)), 0, 0, ms);
}

function moveToPrinted(filePath, watchDir, printedDirName) {
  const printedDir = path.join(watchDir, printedDirName);
  if (!fs.existsSync(printedDir)) {
    fs.mkdirSync(printedDir, { recursive: true });
  }
  const basename = path.basename(filePath);
  const destPath = path.join(printedDir, basename);

  for (let attempt = 1; attempt <= MAX_RETRIES; attempt++) {
    try {
      if (!fs.existsSync(filePath)) {
        console.error('moveToPrinted: ソースファイルが存在しません', { source: filePath });
        return false;
      }
      try {
        fs.renameSync(filePath, destPath);
      } catch (renameErr) {
        if (renameErr.code === 'EXDEV') {
          // ファイルシステム跨ぎ: copy + unlink にフォールバック
          fs.copyFileSync(filePath, destPath);
          fs.unlinkSync(filePath);
        } else {
          throw renameErr;
        }
      }
      return true;
    } catch (e) {
      if (attempt < MAX_RETRIES) {
        console.warn(`moveToPrinted: リトライ ${attempt}/${MAX_RETRIES} (${e.code || e.message})`);
        sleepSync(RETRY_DELAY_MS);
      } else {
        console.error('moveToPrinted: ファイル移動に失敗', {
          source: filePath,
          destination: destPath,
          error: e.message,
        });
        return false;
      }
    }
  }
  return false;
}

module.exports = { moveToPrinted };
