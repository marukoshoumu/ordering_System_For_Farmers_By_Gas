/**
 * AI受注入力アシスタント（v7 - 完全版）
 * - v6の全機能を維持
 * - PDF/画像対応（Gemini Vision）
 * - 商品マッピングの確認・修正後に学習
 * - 顧客別の表記ゆれを記憶
 * - 価格情報対応（顧客指定価格 or マスタ価格）
 */

// Gemini API設定
const GEMINI_API_KEY = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
const GEMINI_MODEL = 'gemini-2.0-flash-exp';
const GEMINI_API_URL = `https://generativelanguage.googleapis.com/v1beta/models/${GEMINI_MODEL}:generateContent`;

// ====================================
// UI用：商品マスタ取得（価格付き）
// ====================================

/**
 * UI用に商品マスタをカテゴリ別で取得（価格情報付き）
 */
function getProductMasterForUI() {
  try {
    const products = getAllRecords('商品');
    const productsByCategory = {};
    const productPrices = {};
    
    products.forEach(p => {
      if (!p['商品名']) return;
      const category = p['商品分類'] || '未分類';
      const productName = p['商品名'];
      const price = p['価格（P)'] || 0;
      
      if (!productsByCategory[category]) {
        productsByCategory[category] = [];
      }
      productsByCategory[category].push(productName);
      productPrices[productName] = price;
    });
    
    return { 
      productsByCategory,
      productPrices
    };
    
  } catch (error) {
    Logger.log('Error in getProductMasterForUI: ' + error.message);
    return { productsByCategory: {}, productPrices: {} };
  }
}

// ====================================
// 商品マッピング自動学習
// ====================================

/**
 * 商品マッピングを登録（UIから呼び出し）
 * @param {string} originalText - 原文の商品表記
 * @param {string} productName - マッピング先の商品名（マスタ）
 * @param {string} customerName - 顧客名（任意）
 * @param {string} spec - 規格情報（任意）
 */
function registerProductMapping(originalText, productName, customerName, spec) {
  try {
    if (!originalText || !productName) {
      return { success: false, error: '原文と商品名は必須です' };
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('商品マッピング');
    
    // シートがなければ作成
    if (!sheet) {
      sheet = ss.insertSheet('商品マッピング');
      sheet.appendRow([
        '顧客表記', '商品名', '商品分類', '規格', '顧客名', 
        '登録日', '使用回数', '最終使用日', '登録者', '備考'
      ]);
      sheet.getRange(1, 1, 1, 10).setBackground('#4a90d9').setFontColor('#ffffff').setFontWeight('bold');
      sheet.setFrozenRows(1);
    }
    
    // 既存マッピングを確認
    const data = sheet.getDataRange().getValues();
    const normalizedOriginal = normalizeForMapping(originalText);
    const normalizedCustomer = customerName ? normalizeForMapping(customerName) : '';
    
    for (let i = 1; i < data.length; i++) {
      const existingOriginal = normalizeForMapping(data[i][0] || '');
      const existingCustomer = normalizeForMapping(data[i][4] || '');
      
      // 同じ表記の場合は更新
      if (existingOriginal === normalizedOriginal && 
          (!normalizedCustomer || existingCustomer === normalizedCustomer || !existingCustomer)) {
        // 商品名を更新（修正された場合）
        if (data[i][1] !== productName) {
          sheet.getRange(i + 1, 2).setValue(productName);
          // 商品分類も更新
          const products = getAllRecords('商品');
          const matchedProduct = products.find(p => p['商品名'] === productName);
          if (matchedProduct) {
            sheet.getRange(i + 1, 3).setValue(matchedProduct['商品分類'] || '');
          }
        }
        // 規格を更新
        if (spec && data[i][3] !== spec) {
          sheet.getRange(i + 1, 4).setValue(spec);
        }
        // 使用回数をインクリメント
        const currentCount = data[i][6] || 0;
        sheet.getRange(i + 1, 7).setValue(currentCount + 1);
        sheet.getRange(i + 1, 8).setValue(new Date());
        
        return { 
          success: true, 
          action: 'updated', 
          message: `マッピングを更新しました（使用回数: ${currentCount + 1}）`
        };
      }
    }
    
    // 新規登録
    let category = '';
    const products = getAllRecords('商品');
    const matchedProduct = products.find(p => p['商品名'] === productName);
    if (matchedProduct) {
      category = matchedProduct['商品分類'] || '';
    }
    
    const newRow = [
      originalText,
      productName,
      category,
      spec || '',
      customerName || '',
      new Date(),
      1,
      new Date(),
      Session.getActiveUser().getEmail() || 'system',
      ''
    ];
    
    sheet.appendRow(newRow);
    
    return { 
      success: true, 
      action: 'created', 
      message: `マッピングを登録しました`
    };
    
  } catch (error) {
    Logger.log('Error in registerProductMapping: ' + error.message);
    return { success: false, error: error.message };
  }
}

/**
 * 複数の商品マッピングを一括登録
 */
function registerProductMappingBulk(mappings) {
  const results = [];
  
  for (const mapping of mappings) {
    const result = registerProductMapping(
      mapping.originalText,
      mapping.productName,
      mapping.customerName,
      mapping.spec
    );
    results.push({
      originalText: mapping.originalText,
      ...result
    });
  }
  
  return results;
}

/**
 * マッピング用の正規化（あいまい検索用）
 */
function normalizeForMapping(str) {
  if (!str) return '';
  return String(str)
    .toLowerCase()
    .replace(/[\s　]+/g, '')         // スペース除去
    .replace(/[_＿－\-]+/g, '')       // 区切り文字除去
    .replace(/[（()）]/g, '')         // 括弧除去
    .replace(/株式会社|有限会社|合同会社/g, '')
    .replace(/（株）|\(株\)|㈱/g, '')
    .trim();
}

/**
 * 既存マッピングから商品を検索
 */
function findProductFromMapping(originalText, customerName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('商品マッピング');
    
    if (!sheet) return null;
    
    const data = sheet.getDataRange().getValues();
    const normalizedOriginal = normalizeForMapping(originalText);
    const normalizedCustomer = customerName ? normalizeForMapping(customerName) : '';
    
    let bestMatch = null;
    let bestScore = 0;
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const mappingOriginal = normalizeForMapping(row[0] || '');
      const mappingProduct = row[1];
      const mappingCategory = row[2];
      const mappingSpec = row[3];
      const mappingCustomer = normalizeForMapping(row[4] || '');
      const usageCount = row[6] || 0;
      
      // 完全一致
      if (mappingOriginal === normalizedOriginal) {
        let score = 100;
        
        // 顧客一致でボーナス
        if (normalizedCustomer && mappingCustomer === normalizedCustomer) {
          score += 50;
        }
        
        // 使用回数でボーナス
        score += Math.min(usageCount * 2, 20);
        
        if (score > bestScore) {
          bestScore = score;
          bestMatch = {
            productName: mappingProduct,
            category: mappingCategory,
            spec: mappingSpec,
            customerSpecific: !!mappingCustomer,
            usageCount: usageCount,
            matchType: 'exact',
            rowIndex: i + 1
          };
        }
      }
      // 部分一致
      else if (mappingOriginal.includes(normalizedOriginal) || normalizedOriginal.includes(mappingOriginal)) {
        let score = 50;
        
        if (normalizedCustomer && mappingCustomer === normalizedCustomer) {
          score += 30;
        }
        
        score += Math.min(usageCount, 10);
        
        if (score > bestScore) {
          bestScore = score;
          bestMatch = {
            productName: mappingProduct,
            category: mappingCategory,
            spec: mappingSpec,
            customerSpecific: !!mappingCustomer,
            usageCount: usageCount,
            matchType: 'partial',
            rowIndex: i + 1
          };
        }
      }
    }
    
    return bestMatch;
    
  } catch (error) {
    Logger.log('Error in findProductFromMapping: ' + error.message);
    return null;
  }
}

// ====================================
// テキスト解析
// ====================================

/**
 * テキストを解析して受注情報を抽出
 */
function analyzeOrderText(text) {
  try {
    const { productList, customerList, shippingToList, mappingList } = getMasterData();
    const prompt = buildTextAnalysisPrompt(text, productList, mappingList, customerList, shippingToList);
    const result = callGeminiAPI([{ text: prompt }]);
    const parsedResult = JSON.parse(result);
    
    // マスタデータで補完
    const enhancedResult = enhanceWithMasterData(parsedResult, customerList, shippingToList);
    
    // 商品マッピングで補完
    enhancedResult.items = enhanceItemsWithMapping(enhancedResult.items, enhancedResult.customer?.rawCompanyName);
    
    // 価格情報を追加
    enhancedResult.items = addPriceToItems(enhancedResult.items, productList);
    
    return JSON.stringify(enhancedResult);
  } catch (error) {
    Logger.log('Error in analyzeOrderText: ' + error.message);
    return JSON.stringify({ success: false, error: error.message });
  }
}

/**
 * Base64エンコードされたファイルを解析
 */
function analyzeOrderFileBase64(base64Data, mimeType, fileName) {
  try {
    Logger.log('Analyzing uploaded file: ' + fileName + ' (' + mimeType + ')');
    
    const { productList, customerList, shippingToList, mappingList } = getMasterData();
    const prompt = buildPDFAnalysisPrompt(productList, mappingList);
    
    const parts = [
      { text: prompt },
      { inline_data: { mime_type: mimeType, data: base64Data } }
    ];
    
    const result = callGeminiAPI(parts);
    const parsedResult = JSON.parse(result);
    
    // マスタデータで補完
    const enhancedResult = enhanceWithMasterData(parsedResult, customerList, shippingToList);
    
    // 商品マッピングで補完
    enhancedResult.items = enhanceItemsWithMapping(enhancedResult.items, enhancedResult.customer?.rawCompanyName);
    
    // 価格情報を追加
    enhancedResult.items = addPriceToItems(enhancedResult.items, productList);
    
    return JSON.stringify(enhancedResult);
    
  } catch (error) {
    Logger.log('Error in analyzeOrderFileBase64: ' + error.message);
    return JSON.stringify({ success: false, error: error.message });
  }
}

/**
 * 商品に価格情報を追加
 * - AIが抽出した価格があればそれを使用
 * - なければマスタ価格を使用
 */
function addPriceToItems(items, productList) {
  if (!items || !Array.isArray(items)) return items;
  
  // 商品名→価格のマップを作成
  const priceMap = {};
  productList.forEach(p => {
    priceMap[p.name] = p.price || 0;
  });
  
  return items.map(item => {
    // AIが価格を抽出している場合はそれを使用（0や空でなければ）
    if (item.price && item.price > 0) {
      item.priceSource = 'ai';  // 価格の出所を記録
    } 
    // マスタから価格を取得
    else if (item.productName && priceMap[item.productName]) {
      item.price = priceMap[item.productName];
      item.priceSource = 'master';
    }
    // 価格が取得できない場合
    else {
      item.price = 0;
      item.priceSource = 'none';
    }
    
    return item;
  });
}

/**
 * 商品リストをマッピングで補完
 */
function enhanceItemsWithMapping(items, customerName) {
  if (!items || !Array.isArray(items)) return items;
  
  return items.map(item => {
    // 既にマッピング情報があればスキップ
    if (item.mappingMatch) return item;
    
    const mapping = findProductFromMapping(item.originalText, customerName);
    
    if (mapping) {
      item.mappingMatch = {
        found: true,
        matchType: mapping.matchType,
        usageCount: mapping.usageCount,
        customerSpecific: mapping.customerSpecific,
        rowIndex: mapping.rowIndex
      };
      
      // 完全一致の場合、マッピングの情報で上書き
      if (mapping.matchType === 'exact') {
        if (item.confidence !== 'high') {
          item.productName = mapping.productName;
          item.category = mapping.category;
          if (mapping.spec) item.spec = mapping.spec;
          item.confidence = 'high';
          item.matchSource = 'mapping';
        }
      } else if (mapping.matchType === 'partial' && item.confidence === 'low') {
        // 部分一致は提案として追加
        item.suggestedProduct = mapping.productName;
        item.suggestedCategory = mapping.category;
      }
    } else {
      item.mappingMatch = { found: false };
      // 学習可能フラグ
      if (item.confidence === 'high' || item.confidence === 'medium') {
        item.canLearn = true;
      }
    }
    
    return item;
  });
}

// ====================================
// マスタデータ取得
// ====================================

function getMasterData() {
  // 商品マスタ
  const products = getAllRecords('商品');
  const productList = products
    .filter(p => p['商品名'])
    .map(p => ({
      name: p['商品名'],
      category: p['商品分類'],
      price: p['価格（P)'] || 0,
      invoiceName: p['送り状品名'] || ''
    }));

  // 商品マッピング（存在する場合）
  let mappingList = [];
  try {
    const mappings = getAllRecords('商品マッピング');
    mappingList = mappings.map(m => ({
      variant: m['顧客表記'] || '',
      productName: m['商品名'],
      category: m['商品分類'],
      spec: m['規格'],
      customerName: m['顧客名']
    }));
  } catch (e) {
    Logger.log('商品マッピングシートなし（初回は自動作成されます）');
  }

  // 顧客マスタ
  const customers = getAllRecords('顧客情報');
  const customerList = customers
    .filter(c => c['会社名'] || c['氏名'])
    .map(c => ({
      displayName: c['表示名'] || '',
      furigana: c['フリガナ'] || '',
      companyName: c['会社名'] || '',
      department: c['部署'] || '',
      personName: c['氏名'] || '',
      zipcode: String(c['郵便番号'] || ''),
      address1: c['住所１'] || '',
      address2: c['住所２'] || '',
      tel: String(c['TEL'] || ''),
      mobile: String(c['携帯電話'] || ''),
      fax: String(c['FAX'] || ''),
      email: c['メールアドレス'] || '',
      category: c['顧客分類'] || ''
    }));

  // 発送先マスタ
  const shippingTos = getAllRecords('発送先情報');
  const shippingToList = shippingTos
    .filter(s => s['会社名'] || s['氏名'])
    .map(s => ({
      companyName: s['会社名'] || '',
      department: s['部署'] || '',
      personName: s['氏名'] || '',
      zipcode: String(s['郵便番号'] || ''),
      address1: s['住所１'] || '',
      address2: s['住所２'] || '',
      tel: String(s['TEL'] || ''),
      email: s['メールアドレス'] || ''
    }));

  return { productList, customerList, shippingToList, mappingList };
}

// ====================================
// プロンプト構築
// ====================================

function buildPDFAnalysisPrompt(productList, mappingList) {
  // 商品マスタをカテゴリ別に整理
  const productsByCategory = {};
  productList.forEach(p => {
    if (!productsByCategory[p.category]) {
      productsByCategory[p.category] = [];
    }
    productsByCategory[p.category].push(`${p.name}（¥${p.price || 0}）`);
  });
  
  let productMasterText = '';
  for (const category in productsByCategory) {
    productMasterText += `【${category}】\n`;
    productMasterText += productsByCategory[category].join('、');
    productMasterText += '\n\n';
  }

  // 学習済みマッピングがあれば追加
  let mappingText = '';
  if (mappingList && mappingList.length > 0) {
    mappingText = '\n═══════════════════════════════════════════════════════════\n';
    mappingText += '【学習済みマッピング】※過去に登録した表記ゆれ\n';
    mappingText += '═══════════════════════════════════════════════════════════\n';
    mappingList.slice(0, 50).forEach(m => {
      mappingText += `「${m.variant}」→「${m.productName}」\n`;
    });
  }

  const today = new Date();
  const year = today.getFullYear();

  return `あなたは受注データ解析の専門家です。
この画像/PDFは発注書、注文書、出荷指示書、FAXのいずれかです。
内容を正確に読み取り、以下のJSON形式で返してください。

═══════════════════════════════════════════════════════════
【商品マスタ】※可能な限りこの中から選んでください
═══════════════════════════════════════════════════════════
${productMasterText}
${mappingText}

═══════════════════════════════════════════════════════════
【本日】${year}年
═══════════════════════════════════════════════════════════

【解析ルール】

■ 文書全体
- 縦書き・横書き両方に対応
- 表形式、リスト形式に対応
- 手書き文字も可能な限り読み取る

■ 日付の解析
- 「納期」「納品日」「着日」「届け日」→ deliveryDate
- 「出荷日」「発送日」「発日」→ shippingDate
- 令和7年 = 2025年、R7 = 2025年、25年 = 2025年

■ 発注元（顧客）
- FAX送信元、発注者欄の会社名/担当者
- 「御中」「様」の前が会社名や個人名

■ 発送先
- 「届け先」「配送先」「納入先」「センター」の記載

■ 商品（重要）
- originalText: 原文の表記をそのまま（学習用に重要）
- productName: 商品マスタから最も近い商品名を選択
- category: 商品マスタの分類
- quantity: 数量（数値のみ）
- price: 単価が記載されていれば抽出（なければ0）
- confidence: マスタとの一致度
  - high: 完全一致 or マッピング一致
  - medium: 表記ゆれあるが特定可能
  - low: 不明確

【出力形式】※JSONのみ、説明文不要

{
  "success": true,
  "documentType": "発注書/FAX/出荷指示書",
  "shippingDate": "YYYY-MM-DD",
  "deliveryDate": "YYYY-MM-DD",
  "deliveryTime": "午前中/14-16時 など",
  
  "customer": {
    "rawCompanyName": "発注元の会社名",
    "rawDepartment": "部署名",
    "rawPersonName": "担当者名",
    "rawTel": "電話番号",
    "rawFax": "FAX番号"
  },
  
  "shippingTo": {
    "rawCompanyName": "発送先の会社名/施設名",
    "rawPersonName": "宛名/担当者",
    "rawZipcode": "郵便番号（7桁数字）",
    "rawAddress": "住所",
    "rawTel": "電話番号"
  },
  
  "items": [
    {
      "originalText": "原文の商品表記をそのまま",
      "productName": "マスタの商品名",
      "category": "商品分類",
      "quantity": 数量,
      "price": 単価,
      "spec": "規格（10kg, 2Lなど）",
      "confidence": "high/medium/low"
    }
  ],
  
  "unknownItems": [
    {
      "originalText": "マスタにない商品の原文",
      "reason": "不明な理由"
    }
  ],
  
  "memo": "備考/連絡事項",
  "alerts": ["注意事項があれば配列で"],
  "overallConfidence": "high/medium/low"
}`;
}

function buildTextAnalysisPrompt(text, productList, mappingList, customerList, shippingToList) {
  // 商品マスタをカテゴリ別に整理
  const productsByCategory = {};
  productList.forEach(p => {
    if (!productsByCategory[p.category]) {
      productsByCategory[p.category] = [];
    }
    productsByCategory[p.category].push(`${p.name}（¥${p.price || 0}）`);
  });
  
  let productMasterText = '';
  for (const category in productsByCategory) {
    productMasterText += `【${category}】\n`;
    productMasterText += productsByCategory[category].join('、');
    productMasterText += '\n\n';
  }

  // 学習済みマッピングがあれば追加
  let mappingText = '';
  if (mappingList && mappingList.length > 0) {
    mappingText = '\n【学習済みマッピング】\n';
    mappingList.slice(0, 50).forEach(m => {
      mappingText += `「${m.variant}」→「${m.productName}」\n`;
    });
  }

  const today = new Date();
  const year = today.getFullYear();

  return `あなたは受注データ解析の専門家です。
以下のテキストはメール、チャット、またはFAXの内容です。
注文内容を抽出してJSON形式で返してください。

【商品マスタ】
${productMasterText}
${mappingText}

【本日】${year}年

【注文内容】
${text}

【出力形式】※JSONのみ
{
  "success": true,
  "shippingDate": "YYYY-MM-DD",
  "deliveryDate": "YYYY-MM-DD",
  "deliveryTime": "時間指定",
  "customer": {
    "rawCompanyName": "会社名",
    "rawPersonName": "氏名",
    "rawTel": "電話番号"
  },
  "shippingTo": {
    "rawCompanyName": "発送先名",
    "rawPersonName": "宛名",
    "rawZipcode": "郵便番号",
    "rawAddress": "住所",
    "rawTel": "電話番号"
  },
  "items": [
    {
      "originalText": "原文",
      "productName": "商品マスタの商品名",
      "category": "商品分類",
      "quantity": 数量,
      "price": 単価,
      "spec": "規格",
      "confidence": "high/medium/low"
    }
  ],
  "unknownItems": [],
  "memo": "備考",
  "alerts": [],
  "overallConfidence": "high/medium/low"
}`;
}

// ====================================
// Gemini API呼び出し
// ====================================

function callGeminiAPI(parts) {
  const payload = {
    contents: [{ parts: parts }],
    generationConfig: {
      temperature: 0.1,
      topK: 1,
      topP: 0.95,
      maxOutputTokens: 8192,
    }
  };

  const options = {
    method: 'POST',
    contentType: 'application/json',
    headers: { 'x-goog-api-key': GEMINI_API_KEY },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(GEMINI_API_URL, options);
  const responseCode = response.getResponseCode();
  const responseText = response.getContentText();

  if (responseCode !== 200) {
    Logger.log('Gemini API Error: ' + responseText);
    throw new Error('Gemini API エラー: ' + responseCode);
  }

  const responseJson = JSON.parse(responseText);
  
  if (!responseJson.candidates || responseJson.candidates.length === 0) {
    throw new Error('Gemini APIからの応答が空です');
  }

  let resultText = responseJson.candidates[0].content.parts[0].text;
  
  // JSONブロックを抽出
  const jsonMatch = resultText.match(/```json\s*([\s\S]*?)\s*```/);
  if (jsonMatch) {
    resultText = jsonMatch[1];
  }
  
  resultText = resultText.trim();
  
  // JSON検証
  try {
    JSON.parse(resultText);
    return resultText;
  } catch (e) {
    Logger.log('JSON parse error. Raw text: ' + resultText);
    throw new Error('AIの応答をJSONとして解析できませんでした');
  }
}

// ====================================
// マスタ照合
// ====================================

/**
 * AI取込一覧から遷移した場合の解析結果を強化
 * クライアントから呼び出して顧客・発送先マスタと照合
 */
function enhanceAnalysisResultFromTempOrder(analysisResultJson) {
  try {
    const result = typeof analysisResultJson === 'string' ? JSON.parse(analysisResultJson) : analysisResultJson;

    // 顧客マスタ取得
    const customerList = getAllRecords('顧客').map(c => ({
      displayName: c['表示名'] || c['会社名'],
      companyName: c['会社名'],
      personName: c['氏名'],
      zipcode: c['郵便番号'],
      address1: c['住所１'],
      address2: c['住所２'],
      tel: c['電話番号'],
      fax: c['ＦＡＸ'],
      email: c['メールアドレス']
    }));

    // 発送先マスタ取得
    const shippingToList = getAllRecords('発送先').map(s => ({
      companyName: s['会社名'],
      personName: s['氏名'],
      zipcode: s['郵便番号'],
      address1: s['住所１'],
      address2: s['住所２'],
      tel: s['電話番号']
    }));

    // マスタデータで強化
    enhanceWithMasterData(result, customerList, shippingToList);

    return JSON.stringify(result);

  } catch (error) {
    Logger.log('enhanceAnalysisResultFromTempOrder エラー: ' + error.toString());
    return analysisResultJson; // エラー時は元のデータを返す
  }
}

function enhanceWithMasterData(result, customerList, shippingToList) {
  if (!result.alerts) result.alerts = [];

  // 顧客照合
  if (result.customer) {
    const customerMatch = findBestMatch(result.customer, customerList, 'customer');
    if (customerMatch.match === 'exact' || customerMatch.match === 'partial') {
      const matched = customerList[customerMatch.index];
      result.customer.masterData = {
        displayName: matched.displayName,
        companyName: matched.companyName,
        personName: matched.personName,
        zipcode: matched.zipcode,
        address: (matched.address1 || '') + (matched.address2 || ''),
        tel: matched.tel,
        fax: matched.fax,
        email: matched.email
      };
      result.customer.masterMatch = customerMatch.match;
      result.customer.matchedBy = customerMatch.matchedBy;
      result.customer.isNewCustomer = false;
      
      if (customerMatch.match === 'partial') {
        result.alerts.push(`顧客「${customerMatch.matchedBy}」で部分一致しました`);
      }
    } else {
      result.customer.masterMatch = 'none';
      result.customer.isNewCustomer = true;
      result.alerts.push('⚠️ 新規顧客の可能性があります');
    }
  }

  // 発送先照合
  if (result.shippingTo) {
    const shippingToMatch = findBestMatch(result.shippingTo, shippingToList, 'shippingTo');
    if (shippingToMatch.match === 'exact' || shippingToMatch.match === 'partial') {
      const matched = shippingToList[shippingToMatch.index];
      result.shippingTo.masterData = {
        companyName: matched.companyName,
        personName: matched.personName,
        zipcode: matched.zipcode,
        address: (matched.address1 || '') + (matched.address2 || ''),
        tel: matched.tel
      };
      result.shippingTo.masterMatch = shippingToMatch.match;
      result.shippingTo.matchedBy = shippingToMatch.matchedBy;
      result.shippingTo.isNewShippingTo = false;
      
      if (shippingToMatch.match === 'partial') {
        result.alerts.push(`発送先「${shippingToMatch.matchedBy}」で部分一致しました`);
      }
    } else {
      result.shippingTo.masterMatch = 'none';
      result.shippingTo.isNewShippingTo = true;
      result.alerts.push('⚠️ 新規発送先の可能性があります');
    }
  }

  return normalizeResult(result);
}

function findBestMatch(rawData, masterList, type) {
  const rawCompany = normalizeString(rawData.rawCompanyName || '');
  const rawPerson = normalizeString(rawData.rawPersonName || '');
  const rawTel = normalizeTel(rawData.rawTel || '');
  const rawZipcode = normalizeZipcode(rawData.rawZipcode || '');
  const rawFax = normalizeTel(rawData.rawFax || '');

  let bestMatch = { match: 'none', index: null, score: 0, matchedBy: '' };

  masterList.forEach((master, index) => {
    let score = 0;
    let matchedBy = [];

    const masterCompany = normalizeString(master.companyName || '');
    const masterPerson = normalizeString(master.personName || '');
    const masterTel = normalizeTel(master.tel || '');
    const masterZipcode = normalizeZipcode(master.zipcode || '');
    const masterFax = normalizeTel(master.fax || '');
    const masterDisplayName = normalizeString(master.displayName || '');

    // 電話番号一致（高信頼度）
    if (rawTel && masterTel && rawTel === masterTel) {
      score += 100;
      matchedBy.push('電話番号');
    }

    // FAX番号一致
    if (rawFax && masterFax && rawFax === masterFax) {
      score += 90;
      matchedBy.push('FAX');
    }

    // 郵便番号一致
    if (rawZipcode && masterZipcode && rawZipcode === masterZipcode) {
      score += 50;
      matchedBy.push('郵便番号');
    }

    // 会社名一致
    if (rawCompany && masterCompany) {
      if (rawCompany === masterCompany) {
        score += 80;
        matchedBy.push('会社名');
      } else if (rawCompany.includes(masterCompany) || masterCompany.includes(rawCompany)) {
        score += 40;
        matchedBy.push('会社名(部分)');
      }
    }

    // 表示名一致
    if (rawCompany && masterDisplayName && rawCompany === masterDisplayName) {
      score += 70;
      matchedBy.push('表示名');
    }

    // 氏名一致
    if (rawPerson && masterPerson) {
      if (rawPerson === masterPerson) {
        score += 60;
        matchedBy.push('氏名');
      } else if (rawPerson.includes(masterPerson) || masterPerson.includes(rawPerson)) {
        score += 30;
        matchedBy.push('氏名(部分)');
      }
    }

    if (score > bestMatch.score) {
      bestMatch = {
        match: score >= 80 ? 'exact' : (score >= 40 ? 'partial' : 'none'),
        index: index,
        score: score,
        matchedBy: matchedBy.join(', ')
      };
    }
  });

  if (bestMatch.score < 40) {
    bestMatch.match = 'none';
    bestMatch.index = null;
  }

  return bestMatch;
}

// ====================================
// ユーティリティ
// ====================================

function normalizeString(str) {
  if (!str) return '';
  return String(str)
    .toLowerCase()
    .replace(/\s+/g, '')
    .replace(/[　]/g, '')
    .replace(/株式会社|有限会社|合同会社/g, '')
    .replace(/（株）|\(株\)|㈱/g, '')
    .replace(/様|御中/g, '')
    .trim();
}

function normalizeZipcode(zipcode) {
  if (!zipcode) return '';
  const normalized = String(zipcode).replace(/[^0-9]/g, '');
  return normalized.length === 7 ? normalized : '';
}

function normalizeTel(tel) {
  if (!tel) return '';
  const normalized = String(tel).replace(/[^0-9]/g, '');
  return (normalized.length >= 10 && normalized.length <= 11) ? normalized : '';
}

function normalizeResult(result) {
  // 日付正規化
  if (result.shippingDate) {
    result.shippingDate = normalizeDate(result.shippingDate);
  }
  if (result.deliveryDate) {
    result.deliveryDate = normalizeDate(result.deliveryDate);
  }
  
  // 商品の数量正規化
  if (result.items) {
    result.items = result.items.map(item => {
      if (item.quantity) {
        item.quantity = parseFloat(item.quantity);
        if (isNaN(item.quantity)) item.quantity = null;
      }
      if (item.price) {
        item.price = parseFloat(item.price);
        if (isNaN(item.price)) item.price = 0;
      }
      return item;
    });
  }
  
  return result;
}

function normalizeDate(dateStr) {
  if (!dateStr) return null;
  if (/^\d{4}-\d{2}-\d{2}$/.test(dateStr)) return dateStr;
  
  try {
    const date = new Date(dateStr);
    if (!isNaN(date.getTime())) {
      return Utilities.formatDate(date, 'JST', 'yyyy-MM-dd');
    }
  } catch (e) {
    // ignore
  }
  
  return null;
}

// ====================================
// API設定確認
// ====================================

function checkGeminiAPIKey() {
  const key = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  return { 
    isSet: !!key, 
    message: key ? 'APIキー設定済み' : 'APIキー未設定。スクリプトプロパティに GEMINI_API_KEY を設定してください' 
  };
}

/**
 * スプレッドシートIDを指定してレコード取得（ライブラリ用）
 */
function getAllRecordsById(spreadsheetId, sheetName) {
  const ss = SpreadsheetApp.openById(spreadsheetId);
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];
  
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
  return records;
}

/**
 * マスタデータ取得（ライブラリ用：スプレッドシートID指定）
 */
function getMasterDataById(spreadsheetId) {
  // 商品マスタ
  const products = getAllRecordsById(spreadsheetId, '商品');
  const productList = products
    .filter(p => p['商品名'])
    .map(p => ({
      name: p['商品名'],
      category: p['商品分類'],
      price: p['価格（P)'] || 0,
      invoiceName: p['送り状品名'] || ''
    }));

  // 商品マッピング
  let mappingList = [];
  try {
    const mappings = getAllRecordsById(spreadsheetId, '商品マッピング');
    mappingList = mappings.map(m => ({
      variant: m['顧客表記'] || '',
      productName: m['商品名'],
      category: m['商品分類'],
      spec: m['規格'],
      customerName: m['顧客名']
    }));
  } catch (e) {
    // マッピングシートがない場合は空配列
  }

  // 顧客マスタ
  const customers = getAllRecordsById(spreadsheetId, '顧客情報');
  const customerList = customers
    .filter(c => c['会社名'] || c['氏名'])
    .map(c => ({
      displayName: c['表示名'] || '',
      companyName: c['会社名'] || '',
      personName: c['氏名'] || '',
      zipcode: String(c['郵便番号'] || ''),
      address1: c['住所１'] || '',
      address2: c['住所２'] || '',
      tel: String(c['TEL'] || ''),
      fax: String(c['FAX'] || '')
    }));

  // 発送先マスタ
  const shippingTos = getAllRecordsById(spreadsheetId, '発送先情報');
  const shippingToList = shippingTos
    .filter(s => s['会社名'] || s['氏名'])
    .map(s => ({
      companyName: s['会社名'] || '',
      personName: s['氏名'] || '',
      zipcode: String(s['郵便番号'] || ''),
      address1: s['住所１'] || '',
      address2: s['住所２'] || '',
      tel: String(s['TEL'] || '')
    }));

  return { productList, customerList, shippingToList, mappingList };
}
/**
 * テキスト解析（ライブラリ用：スプレッドシートID指定）
 */
function analyzeOrderTextById(spreadsheetId, text) {
  try {
    const { productList, customerList, shippingToList, mappingList } = getMasterDataById(spreadsheetId);
    const prompt = buildTextAnalysisPrompt(text, productList, mappingList, customerList, shippingToList);
    const result = callGeminiAPI([{ text: prompt }]);
    const parsedResult = JSON.parse(result);
    
    // マスタデータで補完
    const enhancedResult = enhanceWithMasterData(parsedResult, customerList, shippingToList);
    
    // 商品マッピングで補完
    enhancedResult.items = enhanceItemsWithMapping(enhancedResult.items, enhancedResult.customer?.rawCompanyName);
    
    // 価格情報を追加
    enhancedResult.items = addPriceToItems(enhancedResult.items, productList);
    
    return enhancedResult;
  } catch (error) {
    Logger.log('Error in analyzeOrderTextById: ' + error.message);
    return { success: false, error: error.message };
  }
}