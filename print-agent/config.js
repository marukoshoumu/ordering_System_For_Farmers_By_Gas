const fs = require('fs');
const path = require('path');

const CONFIG_PATH = path.join(__dirname, 'config.json');
const DEFAULTS = {
  watchDir: '',
  printedDirName: 'printed',
  printerName: '',
  debounceMs: 60000,
  logDir: path.join(__dirname, 'logs'),
};

function loadConfig() {
  let fromFile = {};
  if (fs.existsSync(CONFIG_PATH)) {
    fromFile = JSON.parse(fs.readFileSync(CONFIG_PATH, 'utf8'));
  }
  // env が設定されている場合のみ上書き（未設定時は config.json → DEFAULTS の優先順）
  const fromEnv = {};
  if (process.env.WATCH_DIR) fromEnv.watchDir = process.env.WATCH_DIR;
  if (process.env.PRINTED_DIR_NAME) fromEnv.printedDirName = process.env.PRINTED_DIR_NAME;
  if (process.env.PRINTER_NAME) fromEnv.printerName = process.env.PRINTER_NAME;
  if (process.env.DEBOUNCE_MS) fromEnv.debounceMs = parseInt(process.env.DEBOUNCE_MS, 10);
  if (process.env.LOG_DIR) fromEnv.logDir = process.env.LOG_DIR;
  return { ...DEFAULTS, ...fromFile, ...fromEnv };
}

module.exports = { loadConfig };
