/**
 * Drive監視トリガー — 伝票PDFステータス管理
 *
 * 1分間隔の時限トリガーで実行。
 * ヤマト/佐川のDriveフォルダを監視し、
 * description に "pending" を含む新規PDFを検出して
 * ステータスを "pdf_ready" に更新し、ログシートに記録する。
 *
 * 印刷自体は print-agent（ローカルMac上の chokidar + lp）が
 * Google Drive for Desktop の同期フォルダ経由で自動実行する。
 *
 * トリガー設定方法:
 *   GASエディタ → トリガー → 関数: checkSlipStatus → 時間主導型 → 1分おき
 */

/**
 * Workerからのコールバック処理（doPost経由）
 *
 * Worker が処理完了/失敗時に GAS Web App へ POST する。
 * APIキー認証後、伝票処理ログのステータスを即時更新。
 *
 * @param {Object} e - doPost イベントオブジェクト
 * @returns {TextOutput} JSON レスポンス
 */
function handleWorkerCallback(e) {
  try {
    var payload = JSON.parse(e.postData.contents);
    var apiKey = payload.apiKey || '';
    var expectedKey = getWorkerApiKey();

    // APIキー認証（未設定の場合はスキップ）
    if (expectedKey && apiKey !== expectedKey) {
      return ContentService.createTextOutput(
        JSON.stringify({ success: false, message: '認証エラー' })
      ).setMimeType(ContentService.MimeType.JSON);
    }

    var jobId = payload.jobId || '';
    var status = payload.status || '';
    var driveFileId = payload.driveFileId || '';
    var driveWebViewLink = payload.driveWebViewLink || '';
    var errorMessage = payload.error || '';

    if (!jobId || !status) {
      return ContentService.createTextOutput(
        JSON.stringify({ success: false, message: 'jobId と status は必須です' })
      ).setMimeType(ContentService.MimeType.JSON);
    }

    // ステータス更新
    var driveLink = driveWebViewLink || (driveFileId ? 'https://drive.google.com/file/d/' + driveFileId + '/view' : '');
    updateSlipLog(jobId, status, driveLink, errorMessage);

    Logger.log('Workerコールバック受信: jobId=' + jobId + ' status=' + status);

    return ContentService.createTextOutput(
      JSON.stringify({ success: true })
    ).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    Logger.log('コールバック処理エラー: ' + err.message);
    return ContentService.createTextOutput(
      JSON.stringify({ success: false, message: err.message })
    ).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Drive フォルダを監視し、新規PDFのステータスとログを更新
 */
function checkSlipStatus() {
  const folders = [
    { folderId: getYamatoFolderId(), carrier: 'ヤマト運輸' },
    { folderId: getSagawaFolderId(), carrier: '佐川急便' },
  ];

  for (const { folderId, carrier } of folders) {
    if (!folderId) continue;

    try {
      const folder = DriveApp.getFolderById(folderId);
      const files = folder.getFilesByType(MimeType.PDF);

      while (files.hasNext()) {
        const file = files.next();
        const description = file.getDescription() || '';

        // "pending" ステータスのファイルのみ処理
        if (!description.includes('"status":"pending"')) continue;

        let meta;
        try {
          meta = JSON.parse(description);
        } catch {
          continue;
        }

        const fileId = file.getId();
        const jobId = meta.jobId || '';

        Logger.log('新規PDF検出: ' + file.getName() + ' (' + carrier + ')');

        // ステータスを "pdf_ready" に更新
        // print-agent が Google Drive for Desktop 同期で自動印刷する
        meta.status = 'pdf_ready';
        meta.detectedAt = Utilities.formatDate(new Date(), 'JST', 'yyyy/MM/dd HH:mm:ss');
        file.setDescription(JSON.stringify(meta));

        // ログシート更新
        updateSlipLog(jobId, 'pdf_ready', file.getUrl());
        Logger.log('ステータス更新完了: ' + file.getName() + ' → pdf_ready');
      }
    } catch (e) {
      Logger.log('Drive監視エラー (' + carrier + '): ' + e.message);
    }
  }

  // タイムアウト検知: processingのまま10分以上経過したジョブをtimeoutに変更
  checkSlipTimeout();
}

/**
 * processing 状態のまま一定時間経過したジョブを timeout に変更
 */
function checkSlipTimeout() {
  var TIMEOUT_MINUTES = 10;
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('伝票処理ログ');
    if (!sheet) return;

    var data = sheet.getDataRange().getValues();
    var now = new Date();

    for (var i = 1; i < data.length; i++) {
      if (data[i][3] !== 'processing') continue;

      var csvTime = data[i][4];
      if (!csvTime) continue;

      var startTime = (csvTime instanceof Date) ? csvTime : new Date(csvTime);
      if (isNaN(startTime.getTime())) continue;

      var elapsedMin = (now - startTime) / (1000 * 60);
      if (elapsedMin > TIMEOUT_MINUTES) {
        sheet.getRange(i + 1, 4).setValue('timeout');
        sheet.getRange(i + 1, 9).setValue('処理開始から' + Math.floor(elapsedMin) + '分経過（タイムアウト）');
        Logger.log('タイムアウト検出: jobId=' + data[i][2] + ' (' + Math.floor(elapsedMin) + '分)');
      }
    }
  } catch (e) {
    Logger.log('タイムアウト検知エラー: ' + e.message);
  }
}

/**
 * 伝票処理ログを取得（発送伝票画面のステータス一覧用）
 *
 * @param {string} fromDate - 開始日（yyyy-MM-dd）
 * @param {string} toDate - 終了日（yyyy-MM-dd）
 * @returns {string} JSON文字列 [{ shippingDate, carrier, jobId, status, csvTime, pdfTime, printTime, driveLink, error }, ...]
 */
function getSlipLogs(fromDate, toDate) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('伝票処理ログ');
    if (!sheet) return JSON.stringify([]);

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return JSON.stringify([]);

    var from = fromDate ? new Date(fromDate) : null;
    var to = toDate ? new Date(toDate) : null;
    if (from) from.setHours(0, 0, 0, 0);
    if (to) to.setHours(23, 59, 59, 999);

    var logs = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      // A列: 発送日をDateとして比較
      var shipDate = row[0];
      if (shipDate instanceof Date) {
        if (from && shipDate < from) continue;
        if (to && shipDate > to) continue;
        shipDate = Utilities.formatDate(shipDate, 'JST', 'yyyy/MM/dd');
      } else if (typeof shipDate === 'string' && shipDate) {
        var parsed = new Date(shipDate);
        if (!isNaN(parsed.getTime())) {
          if (from && parsed < from) continue;
          if (to && parsed > to) continue;
        }
      }

      logs.push({
        shippingDate: shipDate || '',
        carrier: row[1] || '',
        jobId: row[2] || '',
        status: row[3] || '',
        csvTime: row[4] || '',
        pdfTime: row[5] || '',
        printTime: row[6] || '',
        driveLink: row[7] || '',
        error: row[8] || ''
      });
    }

    // 新しい順にソート
    logs.reverse();
    return JSON.stringify(logs);
  } catch (e) {
    Logger.log('ログ取得エラー: ' + e.message);
    return JSON.stringify([]);
  }
}

/**
 * 伝票処理ログシートのステータスを更新
 *
 * @param {string} jobId - 対象のジョブID
 * @param {string} newStatus - 新しいステータス
 * @param {string} [driveLink] - DriveリンクURL
 * @param {string} [errorDetail] - エラー詳細
 */
function updateSlipLog(jobId, newStatus, driveLink, errorDetail) {
  if (!jobId) return;

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('伝票処理ログ');
    if (!sheet) return;

    const data = sheet.getDataRange().getValues();
    const now = Utilities.formatDate(new Date(), 'JST', 'yyyy/MM/dd HH:mm:ss');

    for (var i = 1; i < data.length; i++) {
      if (data[i][2] === jobId) {
        // D列: ステータス
        sheet.getRange(i + 1, 4).setValue(newStatus);

        if (newStatus === 'pdf_ready') {
          // F列: PDF完了時刻
          if (!data[i][5]) {
            sheet.getRange(i + 1, 6).setValue(now);
          }
        }
        if (driveLink) {
          // H列: Driveリンク
          sheet.getRange(i + 1, 8).setValue(driveLink);
        }
        if (errorDetail) {
          // I列: エラー詳細
          sheet.getRange(i + 1, 9).setValue(errorDetail);
        }
        break;
      }
    }
  } catch (e) {
    Logger.log('ログシート更新エラー: ' + e.message);
  }
}
