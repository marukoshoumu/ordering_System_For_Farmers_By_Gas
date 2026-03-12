const path = require('path');
const chokidar = require('chokidar');
const { loadConfig } = require('./config');
const { createLogger } = require('./logger');
const { ensureWatchDir, ensurePrintedDir } = require('./startup');
const { printPdf } = require('./print');
const { moveToPrinted } = require('./moveToPrinted');

const config = loadConfig();
const logger = createLogger(config.logDir);

try {
  ensureWatchDir(config.watchDir);
  ensurePrintedDir(config.watchDir, config.printedDirName);
} catch (err) {
  logger.error(err.message);
  process.exit(1);
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
  // watcher.close は非同期で完了する場合があるため、一定時間後に強制終了
  setTimeout(() => process.exit(0), 5000);
}
process.on('SIGINT', () => shutdown());
process.on('SIGTERM', () => shutdown());

function handlePdf(filePath) {
  const now = Date.now();
  const last = processed.get(filePath);
  if (last != null && now - last < config.debounceMs) {
    return;
  }
  processed.set(filePath, now);

  const basename = path.basename(filePath);
  logger.info(`Found ${basename}, printing...`);
  const result = printPdf(filePath, config.printerName || undefined);
  if (!result.success) {
    logger.error('Print failed: ' + basename + (result.stderr ? ' stderr: ' + result.stderr : ''));
    return;
  }
  if (result.stderr) logger.warn(`lp stderr: ${result.stderr}`);
  logger.info(`Printed ${basename}, moving to ${config.printedDirName}/`);
  const moved = moveToPrinted(filePath, config.watchDir, config.printedDirName);
  if (!moved) {
    logger.warn(`Move failed: ${basename}`);
  }
}

watcher = chokidar.watch(config.watchDir, {
  ignored: (p) => {
    const rel = path.relative(config.watchDir, p);
    if (rel.startsWith('..')) return false;
    const first = rel.split(path.sep)[0];
    return first === config.printedDirName;
  },
  persistent: true,
  ignoreInitial: true,
});

watcher.on('add', (filePath) => {
  if (path.extname(filePath).toLowerCase() !== '.pdf') return;
  handlePdf(filePath);
});

watcher.on('error', (err) => logger.error(String(err)));
logger.info(`Watching ${config.watchDir}`);
