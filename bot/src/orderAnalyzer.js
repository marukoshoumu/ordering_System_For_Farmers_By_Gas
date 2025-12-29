
// 機密情報は GAS の ScriptProperties で管理
const MASTER_SPREADSHEET_ID = PropertiesService.getScriptProperties().getProperty('MASTER_SPREADSHEET_ID') || '';
// CONFIG に受注システムのURLを追加
const CONFIG = {
  LINE_USER_ID: PropertiesService.getScriptProperties().getProperty('LINE_USER_ID') || '',
  ORDER_SYSTEM_URL: PropertiesService.getScriptProperties().getProperty('ORDER_SYSTEM_URL') || ''
};
/**
 * ライブラリ接続テスト
 */
function testLibraryConnection() {
  
  try {
    // ライブラリ経由でマスタデータ取得
    const masterData = OrderSystem.getMasterDataById(MASTER_SPREADSHEET_ID);
    
    console.log('✅ ライブラリ接続成功！');
    console.log('商品数:', masterData.productList.length);
    console.log('顧客数:', masterData.customerList.length);
    console.log('発送先数:', masterData.shippingToList.length);
    
    if (masterData.productList.length > 0) {
      console.log('商品例:', masterData.productList.slice(0, 3).map(p => p.name));
    }
    
    return { success: true };
  } catch (error) {
    console.log('❌ エラー:', error.message);
    return { success: false, error: error.message };
  }
}
/**
 * Gemini解析テスト
 */
function testGeminiAnalysis() {
  
  const testText = `
田中農園様

いつもお世話になっております。
下記の通り注文いたします。

青首大根 10ケース
ほうれん草 5袋

納品日：12月30日
よろしくお願いいたします。

山田商店
  `;
  
  try {
    const result = OrderSystem.analyzeOrderTextById(MASTER_SPREADSHEET_ID, testText);
    
    console.log('✅ 解析成功！');
    console.log('顧客:', result.customer?.rawCompanyName);
    console.log('商品数:', result.items?.length);
    
    if (result.items) {
      result.items.forEach((item, i) => {
        console.log(`商品${i+1}: ${item.productName} × ${item.quantity}`);
      });
    }
    
    return result;
  } catch (error) {
    console.log('❌ エラー:', error.message);
    return { error: error.message };
  }
}

/**
 * 受注解析モジュール - Gemini API連携
 */

// ========== 設定 ==========
function getGeminiApiKey() {
  return PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
}

// ========== Gemini API呼び出し ==========

/**
 * テキストをGeminiで解析して受注情報を抽出
 */
function analyzeOrderWithGemini(text) {
  const apiKey = getGeminiApiKey();
  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash-exp:generateContent?key=${apiKey}`;
  
  const prompt = buildAnalysisPrompt(text);
  
  const payload = {
    contents: [
      {
        parts: [
          { text: prompt }
        ]
      }
    ],
    generationConfig: {
      temperature: 0.1,
      maxOutputTokens: 2048
    }
  };
  
  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  
  try {
    const response = UrlFetchApp.fetch(url, options);
    const json = JSON.parse(response.getContentText());
    
    if (json.candidates && json.candidates[0]) {
      const resultText = json.candidates[0].content.parts[0].text;
      return parseGeminiResponse(resultText);
    } else {
      console.error('Gemini応答エラー:', json);
      return { success: false, error: 'Gemini応答が不正です' };
    }
  } catch (error) {
    console.error('Gemini API呼び出しエラー:', error);
    return { success: false, error: error.toString() };
  }
}

/**
 * 解析用プロンプトを生成
 */
function buildAnalysisPrompt(text) {
  // マスタデータを取得
  const customers = getCustomerList();
  const products = getProductList();
  const shippingTos = getShippingToList();
  
  return `あなたは受注データ解析のエキスパートです。
以下の受注テキストから情報を抽出し、JSON形式で出力してください。

【マスタデータ】
■顧客リスト:
${customers.map(c => c.name).join(', ')}

■商品リスト:
${products.map(p => p.name + '(' + p.price + '円)').join(', ')}

■発送先リスト:
${shippingTos.map(s => s.name).join(', ')}

【受注テキスト】
${text}

【出力形式】
以下のJSON形式で出力してください。マスタに近い項目があれば正規化してください。
\`\`\`json
{
  "customerName": "顧客名（マスタから選択）",
  "customerMatch": true/false,
  "shippingTo": "発送先名",
  "shippingToMatch": true/false,
  "shippingDate": "YYYY-MM-DD形式または空文字",
  "deliveryDate": "YYYY-MM-DD形式または空文字",
  "items": [
    {
      "productName": "商品名（マスタから選択）",
      "productMatch": true/false,
      "quantity": 数量,
      "unit": "単位（ケース/本/袋/kg等）",
      "price": 単価
    }
  ],
  "memo": "その他特記事項",
  "confidence": "high/medium/low"
}
\`\`\`

JSONのみを出力し、それ以外の文章は含めないでください。`;
}

/**
 * Geminiの応答をパース
 */
function parseGeminiResponse(text) {
  try {
    // ```json ... ``` を除去
    let jsonStr = text.replace(/```json\n?/g, '').replace(/```\n?/g, '').trim();
    const data = JSON.parse(jsonStr);
    data.success = true;
    return data;
  } catch (error) {
    console.error('JSONパースエラー:', error, text);
    return { success: false, error: 'JSONパースに失敗', rawText: text };
  }
}

// ========== マスタデータ取得 ==========

/**
 * 顧客マスタを取得
 */
function getCustomerList() {
  try {
    const ss = SpreadsheetApp.openById(docName);
    const sheet = ss.getSheetByName('顧客情報');
    if (!sheet) return [];
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const nameIdx = headers.indexOf('顧客名');
    
    if (nameIdx === -1) return [];
    
    return data.slice(1).map(row => ({
      name: row[nameIdx] || ''
    })).filter(c => c.name);
  } catch (e) {
    console.log('顧客マスタ取得エラー:', e);
    return [];
  }
}

/**
 * 商品マスタを取得
 */
function getProductList() {
  try {
    const ss = SpreadsheetApp.openById(docName);
    const sheet = ss.getSheetByName('商品');
    if (!sheet) return [];
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const nameIdx = headers.indexOf('商品名');
    const priceIdx = headers.indexOf('価格');
    
    if (nameIdx === -1) return [];
    
    return data.slice(1).map(row => ({
      name: row[nameIdx] || '',
      price: row[priceIdx] || 0
    })).filter(p => p.name);
  } catch (e) {
    console.log('商品マスタ取得エラー:', e);
    return [];
  }
}

/**
 * 発送先マスタを取得
 */
function getShippingToList() {
  try {
    const ss = SpreadsheetApp.openById(docName);
    const sheet = ss.getSheetByName('発送先情報');
    if (!sheet) return [];
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const nameIdx = headers.indexOf('発送先名');
    
    if (nameIdx === -1) return [];
    
    return data.slice(1).map(row => ({
      name: row[nameIdx] || ''
    })).filter(s => s.name);
  } catch (e) {
    console.log('発送先マスタ取得エラー:', e);
    return [];
  }
}

// ========== テスト ==========

/**
 * Gemini解析のテスト
 */
function testGeminiAnalysis() {
  const testText = `
田中農園様

いつもお世話になっております。
下記の通り注文いたします。

青首大根 10ケース
ほうれん草 5袋
にんじん 3kg

納品日：12月30日
よろしくお願いいたします。

山田商店
  `;
  
  const result = analyzeOrderWithGemini(testText);
  console.log('解析結果:', JSON.stringify(result, null, 2));
  return result;
}



/**
 * 修正用URLを生成
 */
function createEditFormUrl(tempOrderId) {
  return CONFIG.ORDER_SYSTEM_URL + '?tempOrderId=' + tempOrderId;
}