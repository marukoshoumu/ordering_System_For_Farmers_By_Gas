function doGet(e) {
  // テスト用：?test=components でコンポーネントテストページを表示
  if (e.parameter.test === 'components') {
    try {
      const template = HtmlService.createTemplateFromFile('test-components');
      const htmlOutput = template.evaluate();
      htmlOutput.addMetaTag('viewport', 'width=device-width, initial-scale=1');
      htmlOutput.setTitle('コンポーネントテスト');
      return htmlOutput;
    } catch (error) {
      // エラーが発生した場合、エラー情報を表示
      return HtmlService.createHtmlOutput(
        '<html><body style="padding:20px;font-family:Arial;"><h1>エラー発生</h1><p>' +
        error.message + '</p><pre>' + error.stack + '</pre></body></html>'
      );
    }
  }

  const template = HtmlService.createTemplateFromFile('index');
  template.deployURL = ScriptApp.getService().getUrl();

  // URLパラメータを保持してログイン画面に渡す
  template.tempOrderId = e.parameter.tempOrderId || '';
  template.redirectTo = e.parameter.aiImportList ? 'aiImportList'
                      : e.parameter.shipping ? 'shipping'
                      : '';

  const htmlOutput = template.evaluate();
  htmlOutput.addMetaTag('viewport', 'width=device-width, initial-scale=1');
  htmlOutput.setTitle('ログイン画面');
  return htmlOutput;
}

function getLoginHTML(alert = '') {
  let html = ``;
  html += `<p class="text-danger">${alert}</p>`
  return html;
}

function doPost(e) {
  Logger.log(e);
  if (e.parameter.login) {
    const items = getAllRecords('pass');
    if (items[0]['id'] != e.parameter.loginId || items[0]['pass'] != e.parameter.loginPass) {
      const template = HtmlService.createTemplateFromFile('login');
      template.deployURL = ScriptApp.getService().getUrl();
      const alert = 'IDまたはパスワードが間違っています。';
      template.loginHTML = getLoginHTML(alert);
      const htmlOutput = template.evaluate();
      htmlOutput.addMetaTag('viewport', 'width=device-width, initial-scale=1');
      htmlOutput.setTitle('ログイン画面');
      return htmlOutput;
    }
    // ログイン成功 - 権限情報を取得
    const userRole = items[0]['権限'] || 'admin';  // デフォルトは管理者
    const tempOrderId = e.parameter.tempOrderId || '';
    const redirectTo = e.parameter.redirectTo || '';

    // viewerの場合、AI取込一覧へのアクセスを制限
    if (userRole === 'viewer' && (redirectTo === 'aiImportList' || tempOrderId)) {
      return redirectToHome(userRole);
    }
    
    // tempOrderIdがあれば受注画面へ直接遷移
    if (tempOrderId) {
      const template = HtmlService.createTemplateFromFile('shipping');
      template.deployURL = ScriptApp.getService().getUrl();
      template.isEditMode = false;
      template.editOrderId = '';
      template.tempOrderId = tempOrderId;
      template.shippingHTML = getshippingHTMLForTempOrder(tempOrderId);
      const htmlOutput = template.evaluate();
      htmlOutput.addMetaTag('viewport', 'width=device-width, initial-scale=1');
      htmlOutput.setTitle('受注画面（AI入力）');
      return htmlOutput;
    }
    
    // リダイレクト先パラメータに応じて遷移
    if (redirectTo === 'aiImportList') {
      const template = HtmlService.createTemplateFromFile('aiImportList');
      template.deployURL = ScriptApp.getService().getUrl();
      const htmlOutput = template.evaluate();
      htmlOutput.addMetaTag('viewport', 'width=device-width, initial-scale=1');
      htmlOutput.setTitle('AI取込一覧');
      return htmlOutput;
    }
    
    if (redirectTo === 'shipping') {
      const template = HtmlService.createTemplateFromFile('shipping');
      template.deployURL = ScriptApp.getService().getUrl();
      template.isEditMode = false;
      template.editOrderId = '';
      template.tempOrderId = '';
      template.autoOpenAI = false;
      template.aiAnalysisResult = '';
      template.shippingHTML = getshippingHTML(e);
      const htmlOutput = template.evaluate();
      htmlOutput.addMetaTag('viewport', 'width=device-width, initial-scale=1');
      htmlOutput.setTitle('受注画面');
      return htmlOutput;
    }
    
    const template = HtmlService.createTemplateFromFile('home');
    template.deployURL = ScriptApp.getService().getUrl();
    template.userRole = userRole;
    const htmlOutput = template.evaluate();
    htmlOutput.addMetaTag('viewport', 'width=device-width, initial-scale=1');
    htmlOutput.setTitle('ホーム画面');
    return htmlOutput;
  }
  else if (e.parameter.main || e.parameter.mainTop) {
    // ホーム画面への遷移
    const userRole = e.parameter.userRole || 'admin';  // パラメータから取得、なければadmin
    const template = HtmlService.createTemplateFromFile('home');
    template.deployURL = ScriptApp.getService().getUrl();
    template.userRole = userRole;
    const htmlOutput = template.evaluate();
    htmlOutput.addMetaTag('viewport', 'width=device-width, initial-scale=1');
    htmlOutput.setTitle('ホーム画面');
    return htmlOutput;
  }
  // 電話受注モード
  else if (e.parameter.phoneOrder) {
    const userRole = e.parameter.userRole || 'admin';
    if (userRole === 'viewer') {
      return redirectToHome(userRole);
    }
    const template = HtmlService.createTemplateFromFile('phoneOrder');
    template.deployURL = ScriptApp.getService().getUrl();
    const htmlOutput = template.evaluate();
    htmlOutput.addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no');
    htmlOutput.setTitle('電話受注モード');
    return htmlOutput;
  }
  // 「受注」ボタンが押されたらshipping.htmlを返す
  else if (e.parameter.shipping) {
    const userRole = e.parameter.userRole || 'admin';
    if (userRole === 'viewer') {
      return redirectToHome(userRole);
    }
    const template = HtmlService.createTemplateFromFile('shipping');
    template.deployURL = ScriptApp.getService().getUrl();

    // 編集モードの設定
    const editOrderId = e.parameter.editOrderId || '';
    const editMode = e.parameter.editMode === 'true';
    template.isEditMode = editMode;
    template.editOrderId = editOrderId;

    // AI取込一覧からの遷移チェック
    const tempOrderId = e.parameter.tempOrderId || '';
    template.tempOrderId = tempOrderId;

    if (tempOrderId) {
      // AI取込一覧から遷移した場合
      const tempOrderData = getTempOrderData(tempOrderId);
      if (tempOrderData && tempOrderData.analysisResult) {
        // 解析結果をテンプレートに渡す
        // analysisResult.data に実際の解析データがある場合はそちらを使用
        const analysisData = tempOrderData.analysisResult.data || tempOrderData.analysisResult;
        template.aiAnalysisResult = JSON.stringify(analysisData);
        template.autoOpenAI = true;
      } else {
        template.autoOpenAI = false;
        template.aiAnalysisResult = '';
      }
      template.shippingHTML = getshippingHTML(e);
    } else {
      // 通常の受注画面
      template.autoOpenAI = false;
      template.aiAnalysisResult = '';
      template.shippingHTML = getshippingHTML(e);
    }

    const htmlOutput = template.evaluate();
    htmlOutput.addMetaTag('viewport', 'width=device-width, initial-scale=1');
    htmlOutput.setTitle(editMode ? '受注修正画面' : '受注画面（AI入力）');
    return htmlOutput;
  }
  // 「受注確認」ボタンが押されたらconfirm.htmlを返す
  else if (e.parameter.shippingComfirm) {
    const userRole = e.parameter.userRole || 'admin';
    if (userRole === 'viewer') {
      return redirectToHome(userRole);
    }
    if (isZero(e)) {
      const template = HtmlService.createTemplateFromFile('shipping');
      const alert = '少なくとも1個以上注文してください。';
      template.deployURL = ScriptApp.getService().getUrl();
      template.shippingHTML = getshippingHTML(e, alert);
      const htmlOutput = template.evaluate();
      htmlOutput.addMetaTag('viewport', 'width=device-width, initial-scale=1');
      htmlOutput.setTitle('受注画面');
      return htmlOutput;
    }

    const template = HtmlService.createTemplateFromFile('shippingComfirm');
    template.deployURL = ScriptApp.getService().getUrl();
    // 編集モードの設定を追加
    const editOrderId = e.parameter.editOrderId || '';
    const editMode = e.parameter.editMode === 'true';
    template.isEditMode = editMode;
    template.editOrderId = editOrderId;
    // 仮受注IDを引き継ぎ
    template.tempOrderId = e.parameter.tempOrderId || '';
    template.confirmHTML = getShippingComfirmHTML(e);
    const htmlOutput = template.evaluate();
    htmlOutput.addMetaTag('viewport', 'width=device-width, initial-scale=1');
    htmlOutput.setTitle('確認画面');
    return htmlOutput;
  }
  // 確認画面で「修正する」ボタンが押されたらform.htmlを返す
  else if (e.parameter.shippingModify) {
    const userRole = e.parameter.userRole || 'admin';
    if (userRole === 'viewer') {
      return redirectToHome(userRole);
    }
    const template = HtmlService.createTemplateFromFile('shipping');
    template.deployURL = ScriptApp.getService().getUrl();
    
    // 編集モードの設定を追加
    const editOrderId = e.parameter.editOrderId || '';
    const editMode = e.parameter.editMode === 'true';
    template.isEditMode = editMode;
    template.editOrderId = editOrderId;
    // 仮受注IDを引き継ぎ
    template.tempOrderId = e.parameter.tempOrderId || '';
    
    // fromConfirmフラグを設定（再検索防止）
    e.parameter.fromConfirm = 'true';
    
    template.shippingHTML = getshippingHTML(e);
    const htmlOutput = template.evaluate();
    htmlOutput.addMetaTag('viewport', 'width=device-width, initial-scale=1');
    htmlOutput.setTitle(editMode ? '受注修正画面' : '受注画面');
    return htmlOutput;
  }
  // 確認画面で「受注する」ボタンが押されたらcomplete画面へ
  else if (e.parameter.shippingSubmit) {
    const userRole = e.parameter.userRole || 'admin';
    if (userRole === 'viewer') {
      return redirectToHome(userRole);
    }
    createOrder(e);
    //updateZaiko(e);
    const template = HtmlService.createTemplateFromFile('shippingComplete');
    template.deployURL = ScriptApp.getService().getUrl();
    const htmlOutput = template.evaluate();
    htmlOutput.addMetaTag('viewport', 'width=device-width, initial-scale=1');
    htmlOutput.setTitle('受注完了');
    return htmlOutput;
  }
  else if (e.parameter.createShippingSlips) {
    const userRole = e.parameter.userRole || 'admin';
    if (userRole === 'viewer') {
      return redirectToHome(userRole);
    }
    const template = HtmlService.createTemplateFromFile('createShippingSlips');
    template.deployURL = ScriptApp.getService().getUrl();
    const htmlOutput = template.evaluate();
    htmlOutput.addMetaTag('viewport', 'width=device-width, initial-scale=1');
    htmlOutput.setTitle('発注伝票作成');
    return htmlOutput;
  }
  else if (e.parameter.createBill) {
    const userRole = e.parameter.userRole || 'admin';
    if (userRole === 'viewer') {
      return redirectToHome(userRole);
    }
    const template = HtmlService.createTemplateFromFile('createBill');
    template.deployURL = ScriptApp.getService().getUrl();
    const htmlOutput = template.evaluate();
    htmlOutput.addMetaTag('viewport', 'width=device-width, initial-scale=1');
    htmlOutput.setTitle('請求書作成');
    return htmlOutput;
  }
  else if (e.parameter.csvImport) {
    const userRole = e.parameter.userRole || 'admin';
    if (userRole === 'viewer') {
      return redirectToHome(userRole);
    }
    const template = HtmlService.createTemplateFromFile('csvImport');
    template.deployURL = ScriptApp.getService().getUrl();
    const htmlOutput = template.evaluate();
    htmlOutput.addMetaTag('viewport', 'width=device-width, initial-scale=1');
    htmlOutput.setTitle('受注CSV取込');
    return htmlOutput;
  }
  else if (e.parameter.quotation) {
    const userRole = e.parameter.userRole || 'admin';
    if (userRole === 'viewer') {
      return redirectToHome(userRole);
    }
    const template = HtmlService.createTemplateFromFile('quotation');
    template.deployURL = ScriptApp.getService().getUrl();
    const htmlOutput = template.evaluate();
    htmlOutput.addMetaTag('viewport', 'width=device-width, initial-scale=1');
    htmlOutput.setTitle('見積書作成');
    return htmlOutput;
  }
  else if (e.parameter.orderList) {
    const userRole = e.parameter.userRole || 'admin';
    const template = HtmlService.createTemplateFromFile('HeatmapOrderList');
    template.deployURL = ScriptApp.getService().getUrl();
    template.userRole = userRole;
    return template.evaluate()
      .setTitle('製造数一覧')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
  else if (e.parameter.orderListPage) {
    const userRole = e.parameter.userRole || 'admin';
    const template = HtmlService.createTemplateFromFile('orderList');
    template.deployURL = ScriptApp.getService().getUrl();
    template.userRole = userRole;
    return template.evaluate()
      .setTitle('受注一覧')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
  // AI取込一覧画面（viewerはアクセス不可）
  else if (e.parameter.aiImportList) {
    const userRole = e.parameter.userRole || 'admin';
    // viewerの不正アクセスをブロック
    if (userRole === 'viewer') {
      return redirectToHome(userRole);
    }
    const template = HtmlService.createTemplateFromFile('aiImportList');
    template.deployURL = ScriptApp.getService().getUrl();
    template.userRole = userRole;
    return template.evaluate()
      .setTitle('AI取込一覧')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
  else {
    const template = HtmlService.createTemplateFromFile('error');
    template.deployURL = ScriptApp.getService().getUrl();
    const htmlOutput = template.evaluate();
    htmlOutput.addMetaTag('viewport', 'width=device-width, initial-scale=1');
    htmlOutput.setTitle('エラー画面');
    return htmlOutput;
  }
}

/**
 * 仮受注IDからshippingHTMLを生成
 */
function getshippingHTMLForTempOrder(tempOrderId) {
  // LINE Bot側のスプレッドシートから仮受注データを取得
  const ss = SpreadsheetApp.openById(getLineBotSpreadsheetId());
  const sheet = ss.getSheetByName('仮受注');
  
  if (!sheet) {
    return getshippingHTML({ parameter: {} }, '仮受注データが見つかりません');
  }
  
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  // 仮受注IDで検索
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === tempOrderId) {
      const analysisResult = JSON.parse(data[i][7]); // 解析結果JSON
      
      // eパラメータを模擬して既存関数を利用
      const mockParam = buildMockParameterFromAnalysis(analysisResult);
      return getshippingHTML({
        parameter: mockParam,
        parameters: {}  // 配列パラメータ用（チェックボックスなど）
      });
    }
  }
  
  return getshippingHTML({
    parameter: {},
    parameters: {}
  }, '該当する仮受注が見つかりません');
}

/**
 * AI解析結果からフォームパラメータを生成
 */
function buildMockParameterFromAnalysis(analysis) {
  const param = {
    fromTempOrder: 'true'
  };
  
  // 顧客情報
  if (analysis.customer) {
    param.customerName = analysis.customer.masterData?.displayName || 
                         analysis.customer.rawCompanyName || '';
    param.customerZipcode = analysis.customer.masterData?.zipcode || '';
    param.customerAddress = analysis.customer.masterData?.address || '';
    param.customerTel = analysis.customer.masterData?.tel || '';
  }
  
  // 発送先情報
  if (analysis.shippingTo) {
    param.shippingToName = analysis.shippingTo.masterData?.companyName ||
                           analysis.shippingTo.rawCompanyName || '';
    param.shippingToZipcode = analysis.shippingTo.masterData?.zipcode || 
                              analysis.shippingTo.rawZipcode || '';
    param.shippingToAddress = analysis.shippingTo.masterData?.address ||
                              analysis.shippingTo.rawAddress || '';
    param.shippingToTel = analysis.shippingTo.masterData?.tel ||
                          analysis.shippingTo.rawTel || '';
  }
  
  // 日付
  param.shippingDate = analysis.shippingDate || '';
  param.deliveryDate = analysis.deliveryDate || '';
  
  // 商品情報（最大10件）
  if (analysis.items && analysis.items.length > 0) {
    analysis.items.slice(0, 10).forEach((item, i) => {
      const num = i + 1;
      param['bunrui' + num] = item.category || '';
      param['product' + num] = item.productName || '';
      param['quantity' + num] = item.quantity || '';
      param['price' + num] = item.price || '';
    });
  }
  
  return param;
}

/**
 * 仮受注IDから仮受注データを取得（AI解析結果を含む）
 */
function getTempOrderData(tempOrderId) {
  try {
    const ss = SpreadsheetApp.openById(getLineBotSpreadsheetId());
    const sheet = ss.getSheetByName('仮受注');

    if (!sheet) {
      Logger.log('仮受注シートが見つかりません');
      return null;
    }

    const data = sheet.getDataRange().getValues();
    const headers = data[0];

    // 仮受注IDで検索
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === tempOrderId) {
        const analysisResultJson = data[i][7]; // 解析結果JSON（列H）

        if (!analysisResultJson) {
          Logger.log('解析結果が空です: ' + tempOrderId);
          return null;
        }

        const analysisResult = JSON.parse(analysisResultJson);

        return {
          tempOrderId: tempOrderId,
          registeredAt: data[i][1],
          status: data[i][2],
          customerName: data[i][3],
          shippingToName: data[i][4],
          itemsJson: data[i][5],
          fileName: data[i][6],
          analysisResult: analysisResult,
          fileUrl: data[i][8]
        };
      }
    }

    Logger.log('該当する仮受注が見つかりません: ' + tempOrderId);
    return null;

  } catch (error) {
    Logger.log('getTempOrderData エラー: ' + error.toString());
    return null;
  }
}

// スプレッドシートのシート名からヘッダの文字列の連想配列を返却
function getAllRecords(sheetName) {
  return getAllRecordsInternal(sheetName, false);
}

function getAllRecordsInternal(sheetName, flg) {
  // WebアプリではgetActiveSpreadsheet()がnullになるため、マスタIDを使用
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    const masterSpreadsheetId = getMasterSpreadsheetId();
    if (masterSpreadsheetId) {
      ss = SpreadsheetApp.openById(masterSpreadsheetId);
    } else {
      Logger.log('getAllRecords: スプレッドシートを開けませんでした');
      return flg ? '[]' : [];
    }
  }
  
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    Logger.log('getAllRecords: シート「' + sheetName + '」が見つかりません');
    return flg ? '[]' : [];
  }
  
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

/**
 * ホーム画面にリダイレクト（権限不足時）
 * @param {string} userRole - ユーザー権限（admin/viewer）
 * @returns {HtmlOutput} ホーム画面
 */
function redirectToHome(userRole) {
  const template = HtmlService.createTemplateFromFile('home');
  template.deployURL = ScriptApp.getService().getUrl();
  template.userRole = userRole || 'admin';
  const htmlOutput = template.evaluate();
  htmlOutput.addMetaTag('viewport', 'width=device-width, initial-scale=1');
  htmlOutput.setTitle('ホーム画面');
  return htmlOutput;
}
