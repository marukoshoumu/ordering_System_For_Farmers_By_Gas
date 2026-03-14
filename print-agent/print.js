const { spawnSync } = require('child_process');
const os = require('os');

const IS_WIN = os.platform() === 'win32';

/**
 * macOS/Linux: lp コマンドで印刷
 */
function printWithLp(filePath, printerName) {
  const args = printerName ? ['-d', printerName, filePath] : [filePath];
  return spawnSync('lp', args, { encoding: 'utf8' });
}

/**
 * Windows: SumatraPDF があれば使用、なければ PowerShell で印刷
 *
 * SumatraPDF は無料・軽量で、サイレント印刷に最適:
 *   choco install sumatrapdf  または  https://www.sumatrapdfreader.org/
 *   パスが通っていれば自動検出される。
 */
function printWithWindows(filePath, printerName) {
  // SumatraPDF を試行
  const sumatraArgs = ['-print-to', printerName || 'default', '-silent', filePath];
  const sumatraResult = spawnSync('SumatraPDF', sumatraArgs, {
    encoding: 'utf8',
    windowsHide: true,
  });
  if (!sumatraResult.error) return sumatraResult;

  // フォールバック: PowerShell
  const psArgs = printerName
    ? ['-NoProfile', '-Command',
      `Start-Process -FilePath '${filePath.replace(/'/g, "''")}' -Verb PrintTo -ArgumentList '${printerName.replace(/'/g, "''")}' -WindowStyle Hidden -Wait`]
    : ['-NoProfile', '-Command',
      `Start-Process -FilePath '${filePath.replace(/'/g, "''")}' -Verb Print -WindowStyle Hidden -Wait`];
  return spawnSync('powershell', psArgs, { encoding: 'utf8', windowsHide: true });
}

/**
 * PDF を印刷する
 * @param {string} filePath - PDF ファイルパス
 * @param {string} [printerName] - プリンター名（省略時はデフォルト）
 * @returns {{ success: boolean, stdout: string, stderr: string }}
 */
function printPdf(filePath, printerName) {
  let result;
  try {
    result = IS_WIN
      ? printWithWindows(filePath, printerName)
      : printWithLp(filePath, printerName);
  } catch (err) {
    return {
      success: false,
      stdout: '',
      stderr: err && (err.message || String(err)) || 'spawnSync threw',
    };
  }
  if (result.error) {
    return {
      success: false,
      stdout: result.stdout || '',
      stderr: result.error.message || String(result.error) || 'spawn error',
    };
  }
  return {
    success: result.status === 0,
    stdout: result.stdout || '',
    stderr: result.stderr || '',
  };
}

module.exports = { printPdf };
