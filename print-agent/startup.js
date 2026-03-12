const fs = require('fs');
const path = require('path');

function ensureWatchDir(watchDir) {
  if (!fs.existsSync(watchDir)) {
    throw new Error(`Watch directory does not exist: ${watchDir}`);
  }
}

function ensurePrintedDir(watchDir, printedDirName) {
  const printedPath = path.join(watchDir, printedDirName);
  if (!fs.existsSync(printedPath)) {
    fs.mkdirSync(printedPath, { recursive: true });
  }
}

module.exports = { ensureWatchDir, ensurePrintedDir };
