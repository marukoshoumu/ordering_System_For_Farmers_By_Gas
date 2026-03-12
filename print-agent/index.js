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
}
process.on('SIGINT', shutdown);
process.on('SIGTERM', shutdown);

function handlePdf(filePath) {
  const now = Date.now();
  const last = processed.get(filePath);
  if (last != null && now - last < config.debounceMs) {
    return;
  }
  processed.set(filePath, now);

  const basename = path.basename(filePath);
  logger.info(`Found ${basename}, printing...`);
  const ok = printPdf(filePath, config.printerName || undefined);
  if (!ok) {
    logger.error(`Print failed: ${basename}`);
    return;
  }
  logger.info(`Printed ${basename}, moving to ${config.printedDirName}/`);
  const moved = moveToPrinted(filePath, config.watchDir, config.printedDirName);
  if (!moved) {
    logger.warn(`Move failed: ${basename}`);
  }
}

watcher = chokidar.watch(config.watchDir, {
  ignored: (p) => path.basename(p) === config.printedDirName,
  persistent: true,
  ignoreInitial: true,
});

watcher.on('add', (filePath) => {
  if (path.extname(filePath).toLowerCase() !== '.pdf') return;
  handlePdf(filePath);
});

watcher.on('error', (err) => logger.error(String(err)));
logger.info(`Watching ${config.watchDir}`);
