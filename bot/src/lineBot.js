/**
 * LINE受注通知Bot - メイン処理
 */

// ========== 設定 ==========
function getLineConfig() {
  const props = PropertiesService.getScriptProperties();
  return {
    channelAccessToken: props.getProperty('LINE_CHANNEL_ACCESS_TOKEN'),
    channelSecret: props.getProperty('LINE_CHANNEL_SECRET')
  };
}

// ========== LINE送信 ==========

/**
 * LINEにテキストメッセージを送信（テスト用）
 */
function sendLineMessage(to, message) {
  const config = getLineConfig();
  const url = 'https://api.line.me/v2/bot/message/push';
  
  const payload = {
    to: to,
    messages: [
      {
        type: 'text',
        text: message
      }
    ]
  };
  
  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'Authorization': 'Bearer ' + config.channelAccessToken
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  
  const response = UrlFetchApp.fetch(url, options);
  console.log('LINE送信結果:', response.getContentText());
  return response;
}

/**
 * 受注確認用のFlex Messageを送信
 */
function sendOrderConfirmation(to, orderData) {
  const config = getLineConfig();
  const url = 'https://api.line.me/v2/bot/message/push';
  
  // Flex Message組み立て
  const flexMessage = buildOrderFlexMessage(orderData);
  
  const payload = {
    to: to,
    messages: [flexMessage]
  };
  
  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'Authorization': 'Bearer ' + config.channelAccessToken
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  
  const response = UrlFetchApp.fetch(url, options);
  console.log('Flex Message送信結果:', response.getContentText());
  return response;
}

/**
 * 受注確認用Flex Messageを組み立て
 */
function buildOrderFlexMessage(orderData) {
  // 商品リストを生成
  const itemContents = orderData.items.map(item => ({
    type: 'box',
    layout: 'horizontal',
    contents: [
      {
        type: 'text',
        text: item.productName,
        size: 'sm',
        color: '#555555',
        flex: 3
      },
      {
        type: 'text',
        text: item.quantity + item.unit,
        size: 'sm',
        color: '#111111',
        align: 'end',
        flex: 1
      }
    ]
  }));
  
  return {
    type: 'flex',
    altText: '新しい受注があります',
    contents: {
      type: 'bubble',
      header: {
        type: 'box',
        layout: 'vertical',
        contents: [
          {
            type: 'text',
            text: '📦 新しい受注',
            weight: 'bold',
            size: 'lg',
            color: '#ffffff'
          }
        ],
        backgroundColor: '#27AE60'
      },
      body: {
        type: 'box',
        layout: 'vertical',
        contents: [
          // 顧客情報
          {
            type: 'box',
            layout: 'horizontal',
            contents: [
              { type: 'text', text: '顧客', size: 'sm', color: '#888888', flex: 1 },
              { type: 'text', text: orderData.customerName + (orderData.customerMatch ? ' ✓' : ' ?'), size: 'sm', flex: 3 }
            ]
          },
          // 発送先
          {
            type: 'box',
            layout: 'horizontal',
            margin: 'md',
            contents: [
              { type: 'text', text: '発送先', size: 'sm', color: '#888888', flex: 1 },
              { type: 'text', text: orderData.shippingTo || '同上', size: 'sm', flex: 3, wrap: true }
            ]
          },
          // 区切り線
          {
            type: 'separator',
            margin: 'lg'
          },
          // 商品リスト
          {
            type: 'box',
            layout: 'vertical',
            margin: 'lg',
            spacing: 'sm',
            contents: itemContents
          },
          // 信頼度
          {
            type: 'text',
            text: '解析信頼度: ' + (orderData.confidence || '中'),
            size: 'xs',
            color: '#888888',
            margin: 'lg'
          }
        ]
      },
      footer: {
        type: 'box',
        layout: 'horizontal',
        spacing: 'sm',
        contents: [
          {
            type: 'button',
            style: 'primary',
            color: '#27AE60',
            action: {
              type: 'postback',
              label: '登録',
              data: 'action=confirm&orderId=' + orderData.tempOrderId
            }
          },
          {
            type: 'button',
            style: 'secondary',
            action: {
              type: 'uri',
              label: '修正',
              uri: orderData.editFormUrl || 'https://example.com'
            }
          }
        ]
      }
    }
  };
}

// ========== Webhook受信 ==========

/**
 * LINEからのWebhookを受信
 */
function doPost(e) {
  const lineBotSpreadsheetId = PropertiesService.getScriptProperties().getProperty('LINE_BOT_SPREADSHEET_ID') || '';
  const ss = SpreadsheetApp.openById(lineBotSpreadsheetId);
  const sheet = ss.getSheetByName("ログ");
  const now = new Date();
  const contents = e.postData.contents;
  
  // ログ記録
  sheet.appendRow([now, "LINE受信", contents]);
  
  try {
    const json = JSON.parse(contents);
    const events = json.events;
    
    for (const event of events) {
      if (event.type === 'postback') {
        // ボタン押下時
        const data = event.postback.data;
        
        // URLSearchParamsの代わりに自前でパース
        const params = {};
        data.split('&').forEach(pair => {
          const [key, value] = pair.split('=');
          params[key] = decodeURIComponent(value || '');
        });
        
        const action = params['action'];
        const orderId = params['orderId'];
        
        sheet.appendRow([now, "Postback", action + " / " + orderId]);
        
        if (action === 'confirm') {
          // 登録ボタン押下 → 返信
          replyMessage(event.replyToken, '✅ 受注を登録しました（ID: ' + orderId + '）');
        }
      }
    }
  } catch (error) {
    sheet.appendRow([now, "エラー", error.toString()]);
  }
  
  return ContentService.createTextOutput(JSON.stringify({status: "ok"}))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * メッセージイベント処理
 */
function handleMessage(event) {
  // 必要に応じてテキスト受信処理を追加
  console.log('メッセージ受信:', event.message);
}

/**
 * 返信メッセージ送信
 */
function replyMessage(replyToken, message) {
  const config = getLineConfig();
  const url = 'https://api.line.me/v2/bot/message/reply';
  
  const payload = {
    replyToken: replyToken,
    messages: [
      {
        type: 'text',
        text: message
      }
    ]
  };
  
  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'Authorization': 'Bearer ' + config.channelAccessToken
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  
  UrlFetchApp.fetch(url, options);
}

/**
 * 受注確定処理
 */
function confirmOrder(tempOrderId) {
  // TODO: 仮登録データを本登録に変更
  console.log('受注確定:', tempOrderId);
  
  // スプレッドシートの該当レコードのステータスを「確定」に更新
  // 実装は後ほど
}

// ========== テスト用 ==========

/**
 * 接続テスト - 自分のLINEにメッセージを送信
 * 実行前に U0b0bef6e2d88d7d62c3f2f47c1351205 を設定してください
 */
function testSendMessage() {
  // LINE公式アカウントを友だち追加した後、
  // Webhook経由で取得したユーザーIDを設定
  const testUserId = 'U0b0bef6e2d88d7d62c3f2f47c1351205';
  
  sendLineMessage(testUserId, 'GASからのテストメッセージです！');
}

/**
 * Flex Messageのテスト送信
 */
function testSendFlexMessage() {
  const testUserId = 'U0b0bef6e2d88d7d62c3f2f47c1351205';
  
  const testOrderData = {
    tempOrderId: 'TEST001',
    customerName: '田中農園',
    customerMatch: true,
    shippingTo: '東京都新宿区1-2-3',
    items: [
      { productName: '青首大根', quantity: 10, unit: 'ケース' },
      { productName: 'ほうれん草', quantity: 5, unit: '袋' }
    ],
    confidence: '高',
    editFormUrl: 'https://docs.google.com/forms/d/xxxxx/viewform'
  };
  
  sendOrderConfirmation(testUserId, testOrderData);
}
function checkConfig() {
  const props = PropertiesService.getScriptProperties();
  const token = props.getProperty('LINE_CHANNEL_ACCESS_TOKEN');
  const secret = props.getProperty('LINE_CHANNEL_SECRET');
  
  console.log('トークン設定あり:', !!token);
  console.log('トークン長さ:', token ? token.length : 0);
  console.log('トークン先頭20文字:', token ? token.substring(0, 20) + '...' : 'なし');
  console.log('シークレット設定あり:', !!secret);
}

/**
 * 設定（ScriptPropertiesから取得）
 */
function getLineConfig() {
  return {
    MASTER_SPREADSHEET_ID: PropertiesService.getScriptProperties().getProperty('MASTER_SPREADSHEET_ID') || '',
    LINE_USER_ID: PropertiesService.getScriptProperties().getProperty('LINE_USER_ID') || ''
  };
}

/**
 * 受注テキストを解析してLINE通知を送信
 */
function processOrder(orderText) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("ログ");
  const now = new Date();
  const lineConfig = getLineConfig();
  
  try {
    // 1. Gemini解析
    sheet.appendRow([now, "解析開始", orderText.substring(0, 100) + "..."]);
    const analysisResult = OrderSystem.analyzeOrderTextById(lineConfig.MASTER_SPREADSHEET_ID, orderText);
    
    if (!analysisResult || analysisResult.error) {
      throw new Error(analysisResult?.error || '解析失敗');
    }
    
    // 2. 仮受注IDを生成
    const tempOrderId = 'TMP' + Utilities.formatDate(now, 'JST', 'yyyyMMddHHmmss');
    
    // 3. 仮登録データを保存
    saveTempOrder(tempOrderId, analysisResult, orderText);
    
    // 4. LINE通知用データを作成
    const orderData = {
      tempOrderId: tempOrderId,
      customerName: analysisResult.customer?.rawCompanyName || '不明',
      customerMatch: analysisResult.customer?.masterMatch === 'exact',
      shippingTo: analysisResult.shippingTo?.rawCompanyName || analysisResult.shippingTo?.rawAddress || '',
      items: (analysisResult.items || []).map(item => ({
        productName: item.productName || item.originalText,
        quantity: item.quantity || 0,
        unit: item.unit || 'ケース'
      })),
      confidence: analysisResult.overallConfidence || 'medium',
      editFormUrl: createEditFormUrl(tempOrderId)
    };
    
    // 5. LINE送信
    sendOrderConfirmation(lineConfig.LINE_USER_ID, orderData);
    
    sheet.appendRow([now, "LINE送信完了", tempOrderId]);
    return { success: true, tempOrderId: tempOrderId };
    
  } catch (error) {
    sheet.appendRow([now, "エラー", error.message]);
    sendLineMessage(lineConfig.LINE_USER_ID, '❌ 解析エラー: ' + error.message);
    return { success: false, error: error.message };
  }
}

/**
 * 仮受注データを保存
 */
function saveTempOrder(tempOrderId, analysisResult, originalText) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("仮受注");
  
  // シートがなければ作成
  if (!sheet) {
    sheet = ss.insertSheet("仮受注");
    sheet.appendRow(["仮受注ID", "登録日時", "ステータス", "顧客名", "発送先", "商品データ", "原文", "解析結果"]);
  }
  
  sheet.appendRow([
    tempOrderId,
    new Date(),
    "未確定",
    analysisResult.customer?.rawCompanyName || '',
    analysisResult.shippingTo?.rawCompanyName || '',
    JSON.stringify(analysisResult.items || []),
    originalText,
    JSON.stringify(analysisResult)
  ]);
}

/**
 * 修正用フォームURLを生成（後で実装）
 */
function createEditFormUrl(tempOrderId) {
  // TODO: Googleフォーム連携
  return 'https://example.com/edit?id=' + tempOrderId;
}

/**
 * 受注確定処理
 */
function confirmOrder(tempOrderId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("仮受注");
  
  if (!sheet) return { success: false, error: '仮受注シートがありません' };
  
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === tempOrderId) {
      // ステータスを「確定」に更新
      sheet.getRange(i + 1, 3).setValue("確定");
      
      // TODO: 本番の受注シートに登録
      
      return { success: true, message: '受注を確定しました' };
    }
  }
  
  return { success: false, error: '該当する仮受注が見つかりません' };
}

/**
 * テスト: 受注処理の全体フロー
 */
function testProcessOrder() {
  const testText = `
お世話になっております。
以下の通り注文します。

大根 10ケース
人参 5ケース

12/30着でお願いします。

田中商店
  `;
  
  const result = processOrder(testText);
  console.log('処理結果:', result);
}