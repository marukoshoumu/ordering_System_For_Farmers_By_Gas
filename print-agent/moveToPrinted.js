const fs = require('fs');
const path = require('path');

function moveToPrinted(filePath, watchDir, printedDirName) {
  const printedDir = path.join(watchDir, printedDirName);
  if (!fs.existsSync(printedDir)) {
    fs.mkdirSync(printedDir, { recursive: true });
  }
  const basename = path.basename(filePath);
  const destPath = path.join(printedDir, basename);
  try {
    fs.renameSync(filePath, destPath);
    return true;
  } catch (e) {
    console.error('moveToPrinted: ファイル移動に失敗', {
      source: filePath,
      destination: destPath,
      error: e.message,
      stack: e.stack,
    });
    return false;
  }
}

module.exports = { moveToPrinted };
