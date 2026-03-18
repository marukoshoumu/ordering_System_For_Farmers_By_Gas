const { spawnSync } = require('child_process');
const path = require('path');
const os = require('os');

const fs = require('fs');
const IS_WIN = os.platform() === 'win32';

/**
 * SumatraPDF の実行ファイルパスを探す
 * @param {string} [configPath] - config.json で指定されたパス
 */
function findSumatraPdf(configPath) {
  if (!IS_WIN) return null;
  const candidates = [
    configPath,
    path.join(process.env.LOCALAPPDATA || '', 'SumatraPDF', 'SumatraPDF.exe'),
    path.join(process.env.PROGRAMFILES || '', 'SumatraPDF', 'SumatraPDF.exe'),
    path.join(process.env['PROGRAMFILES(X86)'] || '', 'SumatraPDF', 'SumatraPDF.exe'),
  ].filter(Boolean);
  for (const p of candidates) {
    if (fs.existsSync(p)) {
      console.log('SumatraPDF found:', p);
      return p;
    }
  }
  console.warn('SumatraPDF not found, will fall back to PowerShell');
  return null;
}

let _sumatraPath;
function getSumatraPath(configPath) {
  if (_sumatraPath === undefined) {
    _sumatraPath = findSumatraPdf(configPath);
  }
  return _sumatraPath;
}

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
function printWithWindows(filePath, printerName, sumatraPdfPath, printSettings) {
  const sumatraExe = getSumatraPath(sumatraPdfPath);
  if (sumatraExe) {
    const settings = printSettings || 'fit';
    const sumatraArgs = printerName
      ? ['-print-to', printerName, '-print-settings', settings, '-silent', filePath]
      : ['-print-to-default', '-print-settings', settings, '-silent', filePath];
    const sumatraResult = spawnSync(sumatraExe, sumatraArgs, {
      encoding: 'utf8',
      windowsHide: true,
    });
    if (!sumatraResult.error) return sumatraResult;
    console.warn('SumatraPDF execution failed:', sumatraResult.error.message);
  }

  // フォールバック: PowerShell（引数を分離して渡しインジェクションを防止）
  const scriptPath = path.join(__dirname, 'scripts', 'print.ps1');
  const psArgs = printerName
    ? ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', scriptPath, filePath, printerName]
    : ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', scriptPath, filePath];
  return spawnSync('powershell', psArgs, { encoding: 'utf8', windowsHide: true });
}

/**
 * PDF を印刷する
 * @param {string} filePath - PDF ファイルパス
 * @param {string} [printerName] - プリンター名（省略時はデフォルト）
 * @param {string} [sumatraPdfPath] - SumatraPDF のフルパス（config.json から）
 * @returns {{ success: boolean, stdout: string, stderr: string }}
 */
function printPdf(filePath, printerName, sumatraPdfPath, printSettings) {
  let result;
  try {
    result = IS_WIN
      ? printWithWindows(filePath, printerName, sumatraPdfPath, printSettings)
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
