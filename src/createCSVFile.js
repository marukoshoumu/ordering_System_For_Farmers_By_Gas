function test6() {
  var record = {};
  record['shippingDate'] = '2024-05-09';
  createShip(JSON.stringify(record));
}
function createShip(datas) {
  var data = JSON.parse(datas);
  var target = Utilities.formatDate(new Date(data['shippingDate']), 'JST', 'yyyy/MM/dd');
  var dateNow = Utilities.formatDate(new Date(), 'JST', 'yyyyMMdd_HHmm');
  const yamatoFileName = 'yamato_' + dateNow + '.csv';
  const sagawaFileName = 'sagawa_' + dateNow + '.csv';
  var record = {};
  if (createCsv('ヤマトCSV', target, yamatoFileName, 1)) {
    record['yamato'] = yamatoFileName;
  }
  if (createCsv('佐川CSV', target, sagawaFileName, 2)) {
    record['sagawa'] = sagawaFileName;
  }
  if (Object.keys(record).length) {
    return JSON.stringify(record);
  }
  else {
    return null;
  }

}
function createCsv(sheetName, dateVal, fileName, type) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  const values = sheet.getDataRange().getValues();
  let csv = '';
  var targetDate = Utilities.formatDate(new Date(dateVal), 'JST', 'yyyy/MM/dd');

  for (var value of values) {
    if (value[0] == targetDate) {
      var val = value.slice(1);
      csv += val.join(',') + "\r\n";
    }
  }
  if (csv.length != 0) {
    const blob = createBlob(csv, fileName);
    writeDrive(blob, type);
    return true;
  }
  else {
    return false;
  }
}

function createBlob(csv, fileName) {
  const contentType = 'text/csv';
  const charset = 'utf-8';
  const blob = Utilities.newBlob('', contentType, fileName).setDataFromString(csv, charset);
  return blob;
}

function writeDrive(blob, type) {
  // CSV出力先
  const yamtoDrive = DriveApp.getFolderById(getYamatoFolderId());
  const sagawaDrive = DriveApp.getFolderById(getSagawaFolderId());
  if (type == 1) {
    yamtoDrive.createFile(blob);
  }
  else {
    sagawaDrive.createFile(blob);
  }
}