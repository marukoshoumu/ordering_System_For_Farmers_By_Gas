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

/**
 * 通知先LINEユーザーIDのリストを取得
 * スクリプトプロパティ LINE_USER_IDS にカンマ区切りで設定
 * 例: U59910f24d0338fe4953bf6490e5d2715,Uxxxxxxxxxxxxx
 */
function getLineUserIds() {
  const props = PropertiesService.getScriptProperties();
  const idsString = props.getProperty('LINE_USER_IDS') || '';
  
  if (!idsString) {
    return [];
  }
  
  // カンマ区切りで分割し、空白を除去
  return idsString.split(',').map(id => id.trim()).filter(id => id);
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
  let sheet = ss.getSheetByName("ログ");
  
  // ログシートがなければ作成
  if (!sheet) {
    sheet = createLogSheet(ss);
  }
  
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

// ========== シート自動作成 ==========

/**
 * ログシートを作成
 */
function createLogSheet(spreadsheet) {
  const sheet = spreadsheet.insertSheet('ログ');
  
  // ヘッダー行を設定
  const headers = ['日時', '種別', '内容'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // ヘッダー行のスタイル設定
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#6aa84f');
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');
  
  // 列幅を調整
  sheet.setColumnWidth(1, 180);  // 日時
  sheet.setColumnWidth(2, 100);  // 種別
  sheet.setColumnWidth(3, 500);  // 内容
  
  // 行を固定
  sheet.setFrozenRows(1);
  
  Logger.log('ログシートを作成しました');
  return sheet;
}

// ========== テスト用 ==========

/**
 * 接続テスト - 自分のLINEにメッセージを送信
 * 実行前にスクリプトプロパティに TEST_LINE_USER_ID を設定してください
 */
function testSendMessage() {
  // スクリプトプロパティからテスト用ユーザーIDを取得
  const testUserId = PropertiesService.getScriptProperties().getProperty('TEST_LINE_USER_ID');
  
  if (!testUserId) {
    Logger.log('エラー: TEST_LINE_USER_ID がスクリプトプロパティに設定されていません');
    return;
  }
  
  sendLineMessage(testUserId, 'GASからのテストメッセージです！');
}

/**
 * Flex Messageのテスト送信
 */
function testSendFlexMessage() {
  const testUserId = PropertiesService.getScriptProperties().getProperty('TEST_LINE_USER_ID');
  
  if (!testUserId) {
    Logger.log('エラー: TEST_LINE_USER_ID がスクリプトプロパティに設定されていません');
    return;
  }
  
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

// ========== 定期リマインド通知 ==========

/**
 * 毎日17時に実行：未処理があればLINE通知を送信
 * トリガーから自動実行される
 * LINE_USER_IDS にカンマ区切りで複数ユーザーIDを設定可能
 */
function sendDailyReminder() {
  try {
    const props = PropertiesService.getScriptProperties();
    const lineBotSpreadsheetId = props.getProperty('LINE_BOT_SPREADSHEET_ID') || '';
    const lineUserIds = getLineUserIds();
    
    if (!lineBotSpreadsheetId || lineUserIds.length === 0) {
      Logger.log('スクリプトプロパティが未設定です');
      return;
    }
    
    // 未処理件数をカウント
    const unprocessedCount = countUnprocessedOrders(lineBotSpreadsheetId);
    
    if (unprocessedCount === 0) {
      Logger.log('未処理なし - 通知スキップ');
      return;
    }
    
    // リマインド通知を全員に送信
    const message = buildReminderFlexMessage(unprocessedCount);
    lineUserIds.forEach(userId => {
      sendFlexMessage(userId, message);
    });
    
    Logger.log(`リマインド通知送信完了: 未処理${unprocessedCount}件 → ${lineUserIds.length}名に送信`);
    
  } catch (error) {
    Logger.log('リマインド通知エラー: ' + error.toString());
  }
}

/**
 * 未処理件数をカウント
 */
function countUnprocessedOrders(spreadsheetId) {
  const ss = SpreadsheetApp.openById(spreadsheetId);
  const sheet = ss.getSheetByName('仮受注');
  
  if (!sheet) {
    Logger.log('仮受注シートが見つかりません');
    return 0;
  }
  
  const data = sheet.getDataRange().getValues();
  let count = 0;
  
  // C列がステータス（インデックス2）
  for (let i = 1; i < data.length; i++) {
    const status = data[i][2];
    if (status === '未処理') {
      count++;
    }
  }
  
  return count;
}

/**
 * リマインド用Flex Messageを組み立て
 */
function buildReminderFlexMessage(count) {
  const props = PropertiesService.getScriptProperties();
  const deployUrl = props.getProperty('DEPLOY_URL') || 'https://script.google.com/macros/s/xxx/exec';
  
  return {
    type: 'flex',
    altText: `未処理が${count}件あります`,
    contents: {
      type: 'bubble',
      body: {
        type: 'box',
        layout: 'vertical',
        contents: [
          {
            type: 'text',
            text: `⏰ 未処理が${count}件あります`,
            weight: 'bold',
            size: 'md',
            wrap: true
          },
          {
            type: 'text',
            text: '確認をお願いします',
            size: 'sm',
            color: '#888888',
            margin: 'md'
          }
        ]
      },
      footer: {
        type: 'box',
        layout: 'vertical',
        contents: [
          {
            type: 'button',
            style: 'primary',
            color: '#27AE60',
            action: {
              type: 'uri',
              label: '確認する',
              uri: deployUrl + '?aiImportList=true'
            }
          }
        ]
      }
    }
  };
}

/**
 * Flex Messageを送信
 */
function sendFlexMessage(to, flexMessage) {
  const config = getLineConfig();
  const url = 'https://api.line.me/v2/bot/message/push';
  
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
  Logger.log('Flex Message送信結果: ' + response.getContentText());
  return response;
}

// ========== トリガー設定 ==========

/**
 * 定期リマインドトリガーを設定（手動実行）
 * 毎日17時にsendDailyReminderを実行
 */
function setupDailyReminderTrigger() {
  // 既存のトリガーを削除
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'sendDailyReminder') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  // 新しいトリガーを作成（毎日17時）
  ScriptApp.newTrigger('sendDailyReminder')
    .timeBased()
    .atHour(17)
    .everyDays(1)
    .create();
  
  Logger.log('トリガーを設定しました: sendDailyReminder を毎日17時に実行');
}

/**
 * 定期リマインドトリガーを削除（手動実行）
 */
function removeDailyReminderTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'sendDailyReminder') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  Logger.log('定期リマインドトリガーを削除しました');
}

/**
 * テスト: リマインド通知を手動送信
 */
function testDailyReminder() {
  sendDailyReminder();
}

// ========== まとめ通知 ==========

/**
 * 5分ごとに実行：未通知の仮受注をまとめて通知
 * トリガーから自動実行される
 * LINE_USER_IDS にカンマ区切りで複数ユーザーIDを設定可能
 */
function checkAndSendBatchNotification() {
  try {
    const props = PropertiesService.getScriptProperties();
    const lineBotSpreadsheetId = props.getProperty('LINE_BOT_SPREADSHEET_ID') || '';
    const lineUserIds = getLineUserIds();
    
    if (!lineBotSpreadsheetId || lineUserIds.length === 0) {
      Logger.log('スクリプトプロパティが未設定です');
      return;
    }
    
    // 未通知データを取得
    const queue = getNotificationQueue(lineBotSpreadsheetId);
    
    if (queue.length === 0) {
      Logger.log('未通知データなし - スキップ');
      return;
    }
    
    // 最初の受信から5分以上経過しているかチェック
    const now = new Date();
    const firstTimestamp = new Date(queue[0].createdAt);
    const diffMinutes = (now - firstTimestamp) / 1000 / 60;
    
    // 5分以上経過、または10件以上たまっていれば送信
    if (diffMinutes >= 5 || queue.length >= 10) {
      Logger.log(`まとめ通知送信: ${queue.length}件 (経過時間: ${diffMinutes.toFixed(1)}分)`);
      
      // まとめ通知を全員に送信
      const message = buildBatchFlexMessage(queue);
      lineUserIds.forEach(userId => {
        sendFlexMessage(userId, message);
      });
      
      // 通知済みにマーク
      markAsSent(lineBotSpreadsheetId, queue);
      
      Logger.log(`まとめ通知完了 → ${lineUserIds.length}名に送信`);
    } else {
      Logger.log(`待機中: ${queue.length}件 (経過時間: ${diffMinutes.toFixed(1)}分)`);
    }
    
  } catch (error) {
    Logger.log('まとめ通知エラー: ' + error.toString());
  }
}

/**
 * 未通知の仮受注データを取得
 */
function getNotificationQueue(spreadsheetId) {
  const ss = SpreadsheetApp.openById(spreadsheetId);
  const sheet = ss.getSheetByName('仮受注');
  
  if (!sheet) {
    Logger.log('仮受注シートが見つかりません');
    return [];
  }
  
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  // 列インデックスを取得
  const idxTempOrderId = 0;   // A列: 仮受注ID
  const idxCreatedAt = 1;     // B列: 登録日時
  const idxStatus = 2;        // C列: ステータス
  const idxCustomerName = 3;  // D列: 顧客名
  
  // 通知済みフラグ列を確認（なければ追加）
  let idxNotified = headers.indexOf('通知済み');
  if (idxNotified === -1) {
    // 列を追加
    const lastCol = sheet.getLastColumn();
    sheet.getRange(1, lastCol + 1).setValue('通知済み');
    sheet.getRange(1, lastCol + 2).setValue('通知日時');
    idxNotified = lastCol; // 0-indexed
  }
  
  const queue = [];
  
  // 未処理かつ未通知のデータを取得
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const status = row[idxStatus];
    const notified = row[idxNotified];
    
    // 未処理かつ未通知（空 or FALSE）
    if (status === '未処理' && !notified) {
      queue.push({
        rowIndex: i + 1,  // 1-indexed for sheet operations
        tempOrderId: row[idxTempOrderId] || '',
        createdAt: row[idxCreatedAt],
        customerName: row[idxCustomerName] || '不明'
      });
    }
  }
  
  // 登録日時でソート（古い順）
  queue.sort((a, b) => new Date(a.createdAt) - new Date(b.createdAt));
  
  return queue;
}

/**
 * 通知済みにマーク
 */
function markAsSent(spreadsheetId, queue) {
  const ss = SpreadsheetApp.openById(spreadsheetId);
  const sheet = ss.getSheetByName('仮受注');
  
  if (!sheet) return;
  
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const idxNotified = headers.indexOf('通知済み') + 1;  // 1-indexed
  const idxNotifiedAt = headers.indexOf('通知日時') + 1;
  
  const now = new Date();
  
  queue.forEach(item => {
    sheet.getRange(item.rowIndex, idxNotified).setValue(true);
    sheet.getRange(item.rowIndex, idxNotifiedAt).setValue(now);
  });
}

/**
 * まとめ通知用Flex Messageを組み立て
 */
function buildBatchFlexMessage(queue) {
  const props = PropertiesService.getScriptProperties();
  const deployUrl = props.getProperty('DEPLOY_URL') || 'https://script.google.com/macros/s/xxx/exec';
  
  // 顧客名リストを作成（最大5件表示）
  const customerNames = queue.slice(0, 5).map(item => `・${item.customerName}様`);
  if (queue.length > 5) {
    customerNames.push(`...他${queue.length - 5}件`);
  }
  
  return {
    type: 'flex',
    altText: `新しい注文が${queue.length}件届きました`,
    contents: {
      type: 'bubble',
      header: {
        type: 'box',
        layout: 'vertical',
        contents: [
          {
            type: 'text',
            text: `📦 新しい注文が${queue.length}件`,
            weight: 'bold',
            size: 'lg',
            color: '#ffffff'
          },
          {
            type: 'text',
            text: '届きました',
            size: 'md',
            color: '#ffffff'
          }
        ],
        backgroundColor: '#27AE60'
      },
      body: {
        type: 'box',
        layout: 'vertical',
        contents: customerNames.map(name => ({
          type: 'text',
          text: name,
          size: 'sm',
          color: '#555555',
          margin: 'sm'
        }))
      },
      footer: {
        type: 'box',
        layout: 'vertical',
        contents: [
          {
            type: 'button',
            style: 'primary',
            color: '#27AE60',
            action: {
              type: 'uri',
              label: '一覧を確認',
              uri: deployUrl + '?aiImportList=true'
            }
          }
        ]
      }
    }
  };
}

// ========== まとめ通知トリガー設定 ==========

/**
 * まとめ通知トリガーを設定（手動実行）
 * 5分ごとにcheckAndSendBatchNotificationを実行
 */
function setupBatchNotificationTrigger() {
  // 既存のトリガーを削除
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'checkAndSendBatchNotification') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  // 新しいトリガーを作成（5分ごと）
  ScriptApp.newTrigger('checkAndSendBatchNotification')
    .timeBased()
    .everyMinutes(5)
    .create();
  
  Logger.log('トリガーを設定しました: checkAndSendBatchNotification を5分ごとに実行');
}

/**
 * まとめ通知トリガーを削除（手動実行）
 */
function removeBatchNotificationTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'checkAndSendBatchNotification') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  Logger.log('まとめ通知トリガーを削除しました');
}

/**
 * テスト: まとめ通知を手動実行
 */
function testBatchNotification() {
  checkAndSendBatchNotification();
}