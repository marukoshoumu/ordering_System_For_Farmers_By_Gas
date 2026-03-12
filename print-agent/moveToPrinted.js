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
  } catch {
    return false;
  }
}

module.exports = { moveToPrinted };
