const { spawnSync } = require('child_process');

/**
 * PDF を lp で印刷する
 * @param {string} filePath - PDF ファイルパス
 * @param {string} [printerName] - プリンター名（省略時はデフォルト）
 * @returns {{ success: boolean, stdout: string, stderr: string }}
 */
function printPdf(filePath, printerName) {
  const args = printerName ? ['-d', printerName, filePath] : [filePath];
  const result = spawnSync('lp', args, { encoding: 'utf8' });
  return {
    success: result.status === 0,
    stdout: result.stdout || '',
    stderr: result.stderr || '',
  };
}

module.exports = { printPdf };
