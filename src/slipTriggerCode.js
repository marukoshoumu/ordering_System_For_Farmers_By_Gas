/**
 * Drive監視トリガー — 伝票PDF自動印刷
 *
 * 1分間隔の時限トリガーで実行。
 * ヤマト/佐川のDriveフォルダを監視し、
 * description に "pending" を含む新規PDFを検出して Brother プリンタに送信する。
 *
 * トリガー設定方法:
 *   GASエディタ → トリガー → 関数: checkAndPrintNewSlips → 時間主導型 → 1分おき
 */

/**
 * Drive フォルダを監視し、未印刷のPDFを Brother プリンタに送信
 */
function checkAndPrintNewSlips() {
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
        const shippingDate = meta.shippingDate || '';
        const jobId = meta.jobId || '';

        Logger.log('未印刷PDF検出: ' + file.getName() + ' (' + carrier + ')');

        // Brother プリンタに送信
        const printResult = sendSlipToPrinter(fileId, carrier, shippingDate);

        if (printResult.success) {
          // description を "printed" に更新
          meta.status = 'printed';
          meta.printedAt = Utilities.formatDate(new Date(), 'JST', 'yyyy/MM/dd HH:mm:ss');
          file.setDescription(JSON.stringify(meta));

          // ログシート更新
          updateSlipLog(jobId, 'printed', file.getUrl());
          Logger.log('印刷完了: ' + file.getName());
        } else {
          // description を "error" に更新
          meta.status = 'error';
          meta.error = printResult.message;
          file.setDescription(JSON.stringify(meta));

          updateSlipLog(jobId, 'error', '', printResult.message);
          Logger.log('印刷失敗: ' + file.getName() + ' - ' + printResult.message);
        }
      }
    } catch (e) {
      Logger.log('Drive監視エラー (' + carrier + '): ' + e.message);
    }
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

        if (newStatus === 'printed') {
          // G列: 印刷時刻
          sheet.getRange(i + 1, 7).setValue(now);
        }
        if (newStatus === 'pdf_ready' || newStatus === 'printed') {
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
