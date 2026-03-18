const path = require('path');
const chokidar = require('chokidar');
const { loadConfig } = require('./config');
const { createLogger } = require('./logger');
const { ensureWatchDir, ensurePrintedDir } = require('./startup');
const { printPdf } = require('./print');
const { moveToPrinted } = require('./moveToPrinted');

const config = loadConfig();
const logger = createLogger(config.logDir);

if (config.watchDirs.length === 0) {
  logger.error('watchDirs が設定されていません。config.json を確認してください。');
  process.exit(1);
}

for (const entry of config.watchDirs) {
  try {
    ensureWatchDir(entry.path);
    ensurePrintedDir(entry.path, config.printedDirName);
  } catch (err) {
    logger.error(err.message);
    process.exit(1);
  }
}

const processed = new Map();

const RETENTION_MS = config.debounceMs * 2;
let cleanupIntervalId = null;
let watcher = null;

function cleanupProcessedMap() {
  const now = Date.now();
  for (const [filePath, timestamp] of processed.entries()) {
    if (now - timestamp >= RETENTION_MS) processed.delete(filePath);
  }
}

cleanupIntervalId = setInterval(cleanupProcessedMap, Math.max(60000, config.debounceMs));

function shutdown() {
  if (cleanupIntervalId) {
    clearInterval(cleanupIntervalId);
    cleanupIntervalId = null;
  }
  if (watcher && typeof watcher.close === 'function') {
    try { watcher.close(); } catch (_) {}
  }
  setTimeout(() => process.exit(0), 5000);
}
process.on('SIGINT', () => shutdown());
process.on('SIGTERM', () => shutdown());

function findWatchEntry(filePath) {
  const normalized = path.normalize(filePath).toLowerCase();
  for (const entry of config.watchDirs) {
    const normalizedEntry = path.normalize(entry.path).toLowerCase();
    if (normalized.startsWith(normalizedEntry + path.sep) || normalized === normalizedEntry) {
      return entry;
    }
  }
  return null;
}

function handlePdf(filePath) {
  const now = Date.now();
  const last = processed.get(filePath);
  if (last != null && now - last < config.debounceMs) {
    return;
  }
  processed.set(filePath, now);

  const entry = findWatchEntry(filePath);
  if (!entry) {
    logger.warn(`Cannot determine watchDir for: ${filePath}`);
    return;
  }

  const basename = path.basename(filePath);
  const printer = entry.printerName || undefined;
  logger.info(`Found ${basename}, printing...` + (printer ? ` (printer: ${printer})` : ''));
  const printSettings = entry.printSettings || config.printSettings || 'fit';
  const result = printPdf(filePath, printer, config.sumatraPdfPath, printSettings);
  if (!result.success) {
    logger.error('Print failed: ' + basename + (result.stderr ? ' stderr: ' + result.stderr : ''));
    return;
  }
  if (result.stderr) logger.warn(`lp stderr: ${result.stderr}`);

  logger.info(`Printed ${basename}, moving to ${config.printedDirName}/`);
  const moved = moveToPrinted(filePath, entry.path, config.printedDirName);
  if (!moved) {
    logger.warn(`Move failed: ${basename}`);
  }
}

const watchPaths = config.watchDirs.map(e => e.path);

const ignored = (p) => {
  for (const entry of config.watchDirs) {
    const rel = path.relative(entry.path, p);
    if (!rel.startsWith('..')) {
      const first = rel.split(path.sep)[0];
      if (first === config.printedDirName) return true;
    }
  }
  return false;
};

watcher = chokidar.watch(watchPaths, {
  ignored,
  persistent: true,
  ignoreInitial: true,
});

watcher.on('add', (filePath) => {
  if (path.extname(filePath).toLowerCase() !== '.pdf') return;
  handlePdf(filePath);
});

watcher.on('error', (err) => logger.error(String(err)));
logger.info(`Watching ${config.watchDirs.length} directories:`);
for (const entry of config.watchDirs) {
  logger.info(`  - ${entry.path}` + (entry.printerName ? ` → printer: ${entry.printerName}` : ''));
}
