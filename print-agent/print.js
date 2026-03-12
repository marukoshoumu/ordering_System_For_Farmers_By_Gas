const { spawnSync } = require('child_process');

function printPdf(filePath, printerName) {
  const args = printerName ? ['-d', printerName, filePath] : [filePath];
  const result = spawnSync('lp', args, { encoding: 'utf8' });
  return result.status === 0;
}

module.exports = { printPdf };
