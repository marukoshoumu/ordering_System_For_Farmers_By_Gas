const fs = require('fs');
const path = require('path');

const CONFIG_PATH = path.join(__dirname, 'config.json');
const DEFAULTS = {
  watchDirs: [],
  printedDirName: 'printed',
  printerName: '',
  debounceMs: 60000,
  logDir: path.join(__dirname, 'logs'),
};

/**
 * watchDirs を { path, printerName } の配列に正規化する。
 * 以下の形式をすべてサポート:
 *   - 文字列配列: ["dir1", "dir2"]
 *   - オブジェクト配列: [{ path: "dir1", printerName: "P1" }]
 *   - 混在: ["dir1", { path: "dir2", printerName: "P2" }]
 *   - 旧形式の単一文字列: "dir1"
 */
function normalizeWatchDirs(raw, defaultPrinter) {
  const arr = Array.isArray(raw) ? raw : (raw ? [raw] : []);
  return arr
    .map((item) => {
      if (typeof item === 'string') {
        return item ? { path: item, printerName: defaultPrinter || '' } : null;
      }
      if (item && typeof item === 'object' && item.path) {
        return { path: item.path, printerName: item.printerName || defaultPrinter || '' };
      }
      return null;
    })
    .filter(Boolean);
}

function loadConfig() {
  let fromFile = {};
  if (fs.existsSync(CONFIG_PATH)) {
    try {
      fromFile = JSON.parse(fs.readFileSync(CONFIG_PATH, 'utf8'));
    } catch (err) {
      console.error('config.json の読み込みに失敗しました:', CONFIG_PATH, err.message);
      fromFile = {};
    }
  }

  const fromEnv = {};
  if (process.env.WATCH_DIRS) {
    fromEnv._rawWatchDirs = process.env.WATCH_DIRS.split(',').map(s => s.trim()).filter(Boolean);
  } else if (process.env.WATCH_DIR) {
    fromEnv._rawWatchDirs = [process.env.WATCH_DIR];
  }
  if (process.env.PRINTED_DIR_NAME) fromEnv.printedDirName = process.env.PRINTED_DIR_NAME;
  if (process.env.PRINTER_NAME) fromEnv.printerName = process.env.PRINTER_NAME;
  if (process.env.DEBOUNCE_MS) {
    const parsed = parseInt(process.env.DEBOUNCE_MS, 10);
    if (Number.isFinite(parsed) && parsed >= 0) fromEnv.debounceMs = parsed;
  }
  if (process.env.LOG_DIR) fromEnv.logDir = process.env.LOG_DIR;

  const merged = { ...DEFAULTS, ...fromFile, ...fromEnv };
  const defaultPrinter = merged.printerName || '';

  // watchDirs の解決: env > file > default
  if (fromEnv._rawWatchDirs) {
    merged.watchDirs = normalizeWatchDirs(fromEnv._rawWatchDirs, defaultPrinter);
  } else {
    const fileRaw = fromFile.watchDirs || fromFile.watchDir;
    merged.watchDirs = fileRaw
      ? normalizeWatchDirs(fileRaw, defaultPrinter)
      : DEFAULTS.watchDirs;
  }
  delete merged._rawWatchDirs;
  delete merged.watchDir;
  return merged;
}

module.exports = { loadConfig };
