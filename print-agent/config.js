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

function normalizeWatchDirs(raw) {
  if (Array.isArray(raw)) return raw.filter(Boolean);
  if (typeof raw === 'string' && raw) return [raw];
  return [];
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
  // watchDir (旧形式) と watchDirs (新形式) の両方をサポート
  const fileWatchDirs = normalizeWatchDirs(fromFile.watchDirs || fromFile.watchDir);

  const fromEnv = {};
  if (process.env.WATCH_DIRS) {
    fromEnv.watchDirs = process.env.WATCH_DIRS.split(',').map(s => s.trim()).filter(Boolean);
  } else if (process.env.WATCH_DIR) {
    fromEnv.watchDirs = [process.env.WATCH_DIR];
  }
  if (process.env.PRINTED_DIR_NAME) fromEnv.printedDirName = process.env.PRINTED_DIR_NAME;
  if (process.env.PRINTER_NAME) fromEnv.printerName = process.env.PRINTER_NAME;
  if (process.env.DEBOUNCE_MS) {
    const parsed = parseInt(process.env.DEBOUNCE_MS, 10);
    if (Number.isFinite(parsed) && parsed >= 0) fromEnv.debounceMs = parsed;
  }
  if (process.env.LOG_DIR) fromEnv.logDir = process.env.LOG_DIR;

  const merged = { ...DEFAULTS, ...fromFile, ...fromEnv };
  merged.watchDirs = fromEnv.watchDirs || (fileWatchDirs.length > 0 ? fileWatchDirs : DEFAULTS.watchDirs);
  delete merged.watchDir;
  return merged;
}

module.exports = { loadConfig };
