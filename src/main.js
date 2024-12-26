function doGet(e) {
  const template = HtmlService.createTemplateFromFile("index");
  template.deployURL = ScriptApp.getService().getUrl();
  const htmlOutput = template.evaluate();
  htmlOutput.addMetaTag("viewport", "width=device-width, initial-scale=1");
  htmlOutput.setTitle("ログイン画面");
  return htmlOutput;
}

function getLoginHTML(alert = "") {
  let html = ``;
  html += `<p class="text-danger">${alert}</p>`;
  return html;
}

function doPost(e) {
  Logger.log(e);
  if (e.parameter.login) {
    const items = getAllRecords("pass");
    if (
      items[0]["id"] != e.parameter.loginId ||
      items[0]["pass"] != e.parameter.loginPass
    ) {
      const template = HtmlService.createTemplateFromFile("login");
      template.deployURL = ScriptApp.getService().getUrl();
      const alert = "IDまたはパスワードが間違っています。";
      template.loginHTML = getLoginHTML(alert);
      const htmlOutput = template.evaluate();
      htmlOutput.addMetaTag("viewport", "width=device-width, initial-scale=1");
      htmlOutput.setTitle("ログイン画面");
      return htmlOutput;
    }
    const template = HtmlService.createTemplateFromFile("home");
    template.deployURL = ScriptApp.getService().getUrl();
    const htmlOutput = template.evaluate();
    htmlOutput.addMetaTag("viewport", "width=device-width, initial-scale=1");
    htmlOutput.setTitle("ホーム画面");
    return htmlOutput;
  } else if (e.parameter.main || e.parameter.mainTop) {
    const template = HtmlService.createTemplateFromFile("home");
    template.deployURL = ScriptApp.getService().getUrl();
    const htmlOutput = template.evaluate();
    htmlOutput.addMetaTag("viewport", "width=device-width, initial-scale=1");
    htmlOutput.setTitle("ホーム画面");
    return htmlOutput;
  }
  // 「受注」ボタンが押されたらshipping.htmlを返す
  else if (e.parameter.shipping) {
    const template = HtmlService.createTemplateFromFile("shipping");
    template.deployURL = ScriptApp.getService().getUrl();
    template.shippingHTML = getshippingHTML(e);
    const htmlOutput = template.evaluate();
    htmlOutput.addMetaTag("viewport", "width=device-width, initial-scale=1");
    htmlOutput.setTitle("受注画面");
    return htmlOutput;
  }
  // 「受注確認」ボタンが押されたらconfirm.htmlを返す
  else if (e.parameter.shippingComfirm) {
    if (isZero(e)) {
      const template = HtmlService.createTemplateFromFile("shipping");
      const alert = "少なくとも1個以上注文してください。";
      template.deployURL = ScriptApp.getService().getUrl();
      template.shippingHTML = getshippingHTML(e, alert);
      const htmlOutput = template.evaluate();
      htmlOutput.addMetaTag("viewport", "width=device-width, initial-scale=1");
      htmlOutput.setTitle("受注画面");
      return htmlOutput;
    }

    const template = HtmlService.createTemplateFromFile("shippingComfirm");
    template.deployURL = ScriptApp.getService().getUrl();
    template.confirmHTML = getShippingComfirmHTML(e);
    const htmlOutput = template.evaluate();
    htmlOutput.addMetaTag("viewport", "width=device-width, initial-scale=1");
    htmlOutput.setTitle("確認画面");
    return htmlOutput;
  }
  // 確認画面で「修正する」ボタンが押されたらform.htmlを返す
  else if (e.parameter.shippingModify) {
    const template = HtmlService.createTemplateFromFile("shipping");
    template.deployURL = ScriptApp.getService().getUrl();
    template.shippingHTML = getshippingHTML(e);
    const htmlOutput = template.evaluate();
    htmlOutput.addMetaTag("viewport", "width=device-width, initial-scale=1");
    htmlOutput.setTitle("受注画面");
    return htmlOutput;
  }
  // 確認画面で「受注する」ボタンが押されたらcomplete画面へ
  else if (e.parameter.shippingSubmit) {
    createOrder(e);
    //updateZaiko(e);
    const template = HtmlService.createTemplateFromFile("shippingComplete");
    template.deployURL = ScriptApp.getService().getUrl();
    const htmlOutput = template.evaluate();
    htmlOutput.addMetaTag("viewport", "width=device-width, initial-scale=1");
    htmlOutput.setTitle("受注完了");
    return htmlOutput;
  } else if (e.parameter.createShippingSlips) {
    const template = HtmlService.createTemplateFromFile("createShippingSlips");
    template.deployURL = ScriptApp.getService().getUrl();
    const htmlOutput = template.evaluate();
    htmlOutput.addMetaTag("viewport", "width=device-width, initial-scale=1");
    htmlOutput.setTitle("発注伝票作成");
    return htmlOutput;
  } else if (e.parameter.createBill) {
    const template = HtmlService.createTemplateFromFile("createBill");
    template.deployURL = ScriptApp.getService().getUrl();
    const htmlOutput = template.evaluate();
    htmlOutput.addMetaTag("viewport", "width=device-width, initial-scale=1");
    htmlOutput.setTitle("請求書作成");
    return htmlOutput;
  } else if (e.parameter.csvImport) {
    const template = HtmlService.createTemplateFromFile("csvImport");
    template.deployURL = ScriptApp.getService().getUrl();
    const htmlOutput = template.evaluate();
    htmlOutput.addMetaTag("viewport", "width=device-width, initial-scale=1");
    htmlOutput.setTitle("受注CSV取込");
    return htmlOutput;
  } else if (e.parameter.quotation) {
    const template = HtmlService.createTemplateFromFile("quotation");
    template.deployURL = ScriptApp.getService().getUrl();
    const htmlOutput = template.evaluate();
    htmlOutput.addMetaTag("viewport", "width=device-width, initial-scale=1");
    htmlOutput.setTitle("見積書作成");
    return htmlOutput;
  } else {
    const template = HtmlService.createTemplateFromFile("error");
    template.deployURL = ScriptApp.getService().getUrl();
    const htmlOutput = template.evaluate();
    htmlOutput.addMetaTag("viewport", "width=device-width, initial-scale=1");
    htmlOutput.setTitle("エラー画面");
    return htmlOutput;
  }
}

function getAllRecords(sheetName) {
  return getAllRecords(sheetName, false);
}
// スプレッドシートのシート名からヘッダの文字列の連想配列を返却
function getAllRecords(sheetName, flg) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  const values = sheet.getDataRange().getValues();
  const labels = values.shift();

  const records = [];
  for (const value of values) {
    const record = {};
    labels.forEach((label, index) => {
      record[label] = value[index];
    });
    records.push(record);
  }
  if (flg) {
    return JSON.stringify(records);
  } else {
    return records;
  }
}
// 商品のゼロ入力チェック
function isZero(e) {
  let total = 0;
  var rowNum = 0;
  for (let i = 0; i < 8; i++) {
    rowNum++;
    var quantity = "quantity" + rowNum;
    const count = Number(e.parameter[quantity]);
    if (count) {
      total += count;
    }
  }
  if (total == 0) {
    return true;
  }
  return false;
}
