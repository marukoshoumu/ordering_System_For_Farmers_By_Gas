function getOpenUrl(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  var ss_url = ss.getUrl();
  var sh_id = sheet.getSheetId();

  return ss_url + "#gid=" + sh_id;
}
function getOpenUrlDrive(sheetName) {
  if (sheetName == "見積書") {
    return "https://drive.google.com/drive/folders/##ID##";
  }
  if (sheetName == "請求書") {
    return "https://drive.google.com/drive/folders/##ID##";
  }
  if (sheetName == "納品書") {
    return "https://drive.google.com/drive/folders/##ID##";
  }
  if (sheetName == "ヤマト") {
    return "https://drive.google.com/drive/folders/##ID##";
  }
  if (sheetName == "佐川") {
    return "https://drive.google.com/drive/folders/##ID##";
  }
}
