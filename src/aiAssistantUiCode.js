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
    // マスタスプレッドシートをIDで直接開く（WebアプリではgetActiveSpreadsheetがnullになるため）
    const masterSpreadsheetId = getMasterSpreadsheetId();
    if (!masterSpreadsheetId) {
      Logger.log('MASTER_SPREADSHEET_ID が未設定です');
      return { productsByCategory: {}, productPrices: {} };
    }
    
    const ss = SpreadsheetApp.openById(masterSpreadsheetId);
    const sheet = ss.getSheetByName('商品');
    
    if (!sheet) {
      Logger.log('商品シートが見つかりません');
      return { productsByCategory: {}, productPrices: {} };
    }
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const productsByCategory = {};
    const productPrices = {};
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const record = {};
      headers.forEach((h, idx) => { record[h] = row[idx]; });
      
      if (!record['商品名']) continue;
      const category = record['商品分類'] || '未分類';
      const productName = record['商品名'];
      const price = record['価格（P)'] || 0;
      
      if (!productsByCategory[category]) {
        productsByCategory[category] = [];
      }
      productsByCategory[category].push(productName);
      productPrices[productName] = price;
    }
    
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
 * @param {string} shippingToName - 発送先名（優先）
 * @param {string} customerName - 顧客名（後方互換用）
 * @param {string} spec - 規格情報（任意）
 */
function registerProductMapping(originalText, productName, shippingToName, customerName, spec) {
  try {
    if (!originalText || !productName) {
      return { success: false, error: '原文と商品名は必須です' };
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('商品マッピング');
    
    // シートがなければ作成（新カラム構成）
    if (!sheet) {
      sheet = ss.insertSheet('商品マッピング');
      sheet.appendRow([
        '顧客表記', '商品名', '商品分類', '規格', '発送先名', '顧客名', 
        '登録日', '使用回数', '最終使用日', '登録者', '備考'
      ]);
      sheet.getRange(1, 1, 1, 11).setBackground('#4a90d9').setFontColor('#ffffff').setFontWeight('bold');
      sheet.setFrozenRows(1);
    }
    
    // カラム構成を確認（旧形式の場合は発送先名カラムを追加）
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const hasShippingToColumn = headers.includes('発送先名');
    if (!hasShippingToColumn && headers.length >= 5) {
      // 旧形式: 顧客名の前に発送先名カラムを挿入
      sheet.insertColumnAfter(4);
      sheet.getRange(1, 5).setValue('発送先名').setBackground('#4a90d9').setFontColor('#ffffff').setFontWeight('bold');
    }
    
    // カラムインデックスを動的に取得
    const currentHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const colIdx = {
      originalText: 0,
      productName: 1,
      category: 2,
      spec: 3,
      shippingTo: currentHeaders.indexOf('発送先名'),
      customer: currentHeaders.indexOf('顧客名'),
      regDate: currentHeaders.indexOf('登録日'),
      usageCount: currentHeaders.indexOf('使用回数'),
      lastUsed: currentHeaders.indexOf('最終使用日'),
      registrant: currentHeaders.indexOf('登録者'),
      memo: currentHeaders.indexOf('備考')
    };
    
    // 既存マッピングを確認
    const data = sheet.getDataRange().getValues();
    const normalizedOriginal = normalizeForMapping(originalText);
    const normalizedShippingTo = shippingToName ? normalizeForMapping(shippingToName) : '';
    const normalizedCustomer = customerName ? normalizeForMapping(customerName) : '';
    
    for (let i = 1; i < data.length; i++) {
      const existingOriginal = normalizeForMapping(data[i][colIdx.originalText] || '');
      const existingShippingTo = colIdx.shippingTo >= 0 ? normalizeForMapping(data[i][colIdx.shippingTo] || '') : '';
      const existingCustomer = colIdx.customer >= 0 ? normalizeForMapping(data[i][colIdx.customer] || '') : '';
      
      // 同じ表記 + 発送先名一致の場合は更新
      const shippingToMatch = normalizedShippingTo && existingShippingTo === normalizedShippingTo;
      const customerMatch = !normalizedShippingTo && normalizedCustomer && existingCustomer === normalizedCustomer;
      const noSpecificMatch = !normalizedShippingTo && !normalizedCustomer && !existingShippingTo && !existingCustomer;
      
      if (existingOriginal === normalizedOriginal && (shippingToMatch || customerMatch || noSpecificMatch)) {
        // 商品名を更新（修正された場合）
        if (data[i][colIdx.productName] !== productName) {
          sheet.getRange(i + 1, colIdx.productName + 1).setValue(productName);
          // 商品分類も更新
          const products = getAllRecords('商品');
          const matchedProduct = products.find(p => p['商品名'] === productName);
          if (matchedProduct) {
            sheet.getRange(i + 1, colIdx.category + 1).setValue(matchedProduct['商品分類'] || '');
          }
        }
        // 規格を更新
        if (spec && data[i][colIdx.spec] !== spec) {
          sheet.getRange(i + 1, colIdx.spec + 1).setValue(spec);
        }
        // 使用回数をインクリメント
        const currentCount = colIdx.usageCount >= 0 ? (data[i][colIdx.usageCount] || 0) : 0;
        if (colIdx.usageCount >= 0) sheet.getRange(i + 1, colIdx.usageCount + 1).setValue(currentCount + 1);
        if (colIdx.lastUsed >= 0) sheet.getRange(i + 1, colIdx.lastUsed + 1).setValue(new Date());
        
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
    
    // 新カラム構成で登録
    const newRow = [
      originalText,
      productName,
      category,
      spec || '',
      shippingToName || '',  // 発送先名
      customerName || '',     // 顧客名（後方互換）
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

// ====================================
// 顧客マスタ FAX番号更新
// ====================================

/**
 * 顧客マスタのFAX番号を更新
 * @param {string} customerDisplayName - 顧客の表示名（会社名＋氏名 or displayName）
 * @param {string} customerCompanyName - 顧客の会社名
 * @param {string} customerPersonName - 顧客の氏名
 * @param {string} faxNumber - 登録するFAX番号
 * @returns {Object} - 処理結果
 */
function updateCustomerFax(customerDisplayName, customerCompanyName, customerPersonName, faxNumber) {
  try {
    if (!faxNumber) {
      return { success: false, error: 'FAX番号が指定されていません' };
    }

    // 正規化
    const normalizedFax = String(faxNumber).replace(/[-ー−‐\s]/g, '');
    if (!/^\d{10,11}$/.test(normalizedFax)) {
      return { success: false, error: 'FAX番号の形式が正しくありません' };
    }

    const masterSpreadsheetId = getMasterSpreadsheetId();
    if (!masterSpreadsheetId) {
      return { success: false, error: 'MASTER_SPREADSHEET_ID が未設定です' };
    }

    const ss = SpreadsheetApp.openById(masterSpreadsheetId);
    const sheet = ss.getSheetByName('顧客情報');
    if (!sheet) {
      return { success: false, error: '顧客情報シートが見つかりません' };
    }

    const data = sheet.getDataRange().getValues();
    const headers = data[0];

    // 列インデックスを取得
    const displayNameCol = headers.indexOf('表示名');
    const companyNameCol = headers.indexOf('会社名');
    const personNameCol = headers.indexOf('氏名');
    const faxCol = headers.indexOf('FAX');

    if (faxCol === -1) {
      return { success: false, error: 'FAX列が見つかりません' };
    }

    // 顧客を検索（会社名＋氏名 or 表示名で照合）
    let foundRow = -1;
    for (let i = 1; i < data.length; i++) {
      const rowDisplayName = data[i][displayNameCol] || '';
      const rowCompanyName = data[i][companyNameCol] || '';
      const rowPersonName = data[i][personNameCol] || '';

      // 会社名＋氏名で照合
      if (customerCompanyName && customerCompanyName === rowCompanyName) {
        if (!customerPersonName || customerPersonName === rowPersonName) {
          foundRow = i + 1; // 1-indexed for GAS
          break;
        }
      }
      // 表示名で照合
      if (customerDisplayName && customerDisplayName === rowDisplayName) {
        foundRow = i + 1;
        break;
      }
    }

    if (foundRow === -1) {
      return { success: false, error: '該当する顧客が見つかりません' };
    }

    // 既存のFAX番号を確認
    const existingFax = data[foundRow - 1][faxCol] || '';
    if (existingFax) {
      return { 
        success: false, 
        error: '既にFAX番号が登録されています: ' + existingFax,
        existingFax: existingFax
      };
    }

    // FAX番号を更新
    sheet.getRange(foundRow, faxCol + 1).setValue(normalizedFax);

    // バックアップシートも更新
    const bkSheet = ss.getSheetByName('顧客情報BK');
    if (bkSheet) {
      const bkData = bkSheet.getDataRange().getValues();
      for (let i = 1; i < bkData.length; i++) {
        const rowCompanyName = bkData[i][companyNameCol] || '';
        const rowPersonName = bkData[i][personNameCol] || '';
        if (customerCompanyName && customerCompanyName === rowCompanyName) {
          if (!customerPersonName || customerPersonName === rowPersonName) {
            bkSheet.getRange(i + 1, faxCol + 1).setValue(normalizedFax);
            break;
          }
        }
      }
    }

    Logger.log(`顧客FAX番号を更新: ${customerCompanyName || customerDisplayName} → ${normalizedFax}`);
    return { 
      success: true, 
      message: 'FAX番号を登録しました',
      faxNumber: normalizedFax
    };

  } catch (error) {
    Logger.log('Error in updateCustomerFax: ' + error.message);
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
      mapping.shippingToName || '',  // 発送先名優先
      mapping.customerName || '',     // 顧客名（後方互換）
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
 * @param {string} originalText - 原文の商品表記
 * @param {string} shippingToName - 発送先名（優先）
 * @param {string} customerName - 顧客名（後方互換）
 */
function findProductFromMapping(originalText, shippingToName, customerName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('商品マッピング');
    
    if (!sheet) return null;
    
    // カラムインデックスを動的に取得
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const colIdx = {
      originalText: 0,
      productName: 1,
      category: 2,
      spec: 3,
      shippingTo: headers.indexOf('発送先名'),
      customer: headers.indexOf('顧客名'),
      usageCount: headers.indexOf('使用回数')
    };
    
    const data = sheet.getDataRange().getValues();
    const normalizedOriginal = normalizeForMapping(originalText);
    const normalizedShippingTo = shippingToName ? normalizeForMapping(shippingToName) : '';
    const normalizedCustomer = customerName ? normalizeForMapping(customerName) : '';
    
    let bestMatch = null;
    let bestScore = 0;
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const mappingOriginal = normalizeForMapping(row[colIdx.originalText] || '');
      const mappingProduct = row[colIdx.productName];
      const mappingCategory = row[colIdx.category];
      const mappingSpec = row[colIdx.spec];
      const mappingShippingTo = colIdx.shippingTo >= 0 ? normalizeForMapping(row[colIdx.shippingTo] || '') : '';
      const mappingCustomer = colIdx.customer >= 0 ? normalizeForMapping(row[colIdx.customer] || '') : '';
      const usageCount = colIdx.usageCount >= 0 ? (row[colIdx.usageCount] || 0) : 0;
      
      // 完全一致
      if (mappingOriginal === normalizedOriginal) {
        let score = 100;
        
        // 発送先一致で最大ボーナス（優先）
        if (normalizedShippingTo && mappingShippingTo === normalizedShippingTo) {
          score += 70;
        }
        // 顧客一致でボーナス（後方互換）
        else if (normalizedCustomer && mappingCustomer === normalizedCustomer) {
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
            shippingToSpecific: !!mappingShippingTo,
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
        
        if (normalizedShippingTo && mappingShippingTo === normalizedShippingTo) {
          score += 40;
        } else if (normalizedCustomer && mappingCustomer === normalizedCustomer) {
          score += 30;
        }
        
        score += Math.min(usageCount, 10);
        
        if (score > bestScore) {
          bestScore = score;
          bestMatch = {
            productName: mappingProduct,
            category: mappingCategory,
            spec: mappingSpec,
            shippingToSpecific: !!mappingShippingTo,
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
    
    // 発送先名と顧客名を取得（発送先優先）
    const shippingToName = enhancedResult.shippingTo?.rawCompanyName || 
                           enhancedResult.shippingTo?.masterData?.companyName || '';
    const customerName = enhancedResult.customer?.rawCompanyName || '';
    
    // 商品マッピングで補完
    enhancedResult.items = enhanceItemsWithMapping(enhancedResult.items, shippingToName, customerName);
    
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
    const prompt = buildPDFAnalysisPrompt(productList, mappingList, customerList, shippingToList);
    
    const parts = [
      { text: prompt },
      { inline_data: { mime_type: mimeType, data: base64Data } }
    ];
    
    const result = callGeminiAPI(parts);
    const parsedResult = JSON.parse(result);
    
    // マスタデータで補完
    const enhancedResult = enhanceWithMasterData(parsedResult, customerList, shippingToList);
    
    // 発送先名と顧客名を取得（発送先優先）
    const shippingToName = enhancedResult.shippingTo?.rawCompanyName || 
                           enhancedResult.shippingTo?.masterData?.companyName || '';
    const customerName = enhancedResult.customer?.rawCompanyName || '';
    
    // 商品マッピングで補完
    enhancedResult.items = enhanceItemsWithMapping(enhancedResult.items, shippingToName, customerName);
    
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
 * @param {Array} items - 商品リスト
 * @param {string} shippingToName - 発送先名（優先）
 * @param {string} customerName - 顧客名（後方互換）
 */
function enhanceItemsWithMapping(items, shippingToName, customerName) {
  if (!items || !Array.isArray(items)) return items;
  
  return items.map(item => {
    // 既にマッピング情報があればスキップ
    if (item.mappingMatch) return item;
    
    const mapping = findProductFromMapping(item.originalText, shippingToName, customerName);
    
    if (mapping) {
      item.mappingMatch = {
        found: true,
        matchType: mapping.matchType,
        usageCount: mapping.usageCount,
        shippingToSpecific: mapping.shippingToSpecific,
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
      shippingToName: m['発送先名'] || '',  // 発送先名優先
      customerName: m['顧客名'] || ''            // 後方互換
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

/**
 * 発送先別マッピングセクションを構築
 * テキスト内の発送先名を検出し、その発送先のマッピングを優先表示
 * @param {Array} mappingList - マッピング一覧
 * @param {string} text - 解析対象テキスト（オプション）
 * @param {Array} customerList - 顧客一覧（後方互換用）
 * @param {Array} shippingToList - 発送先一覧（オプション）
 * @returns {string} プロンプト用のマッピングセクション
 */
function buildCustomerMappingSection(mappingList, text, customerList, shippingToList) {
  if (!mappingList || mappingList.length === 0) {
    return '';
  }

  // 発送先別・顧客別にマッピングをグループ化
  const mappingByShippingTo = {};
  const mappingByCustomer = {};
  const generalMappings = [];

  mappingList.forEach(m => {
    // 発送先名優先
    if (m.shippingToName && m.shippingToName.trim()) {
      if (!mappingByShippingTo[m.shippingToName]) {
        mappingByShippingTo[m.shippingToName] = [];
      }
      mappingByShippingTo[m.shippingToName].push(m);
    }
    // 顧客名（後方互換）
    else if (m.customerName && m.customerName.trim()) {
      if (!mappingByCustomer[m.customerName]) {
        mappingByCustomer[m.customerName] = [];
      }
      mappingByCustomer[m.customerName].push(m);
    }
    // 汎用
    else {
      generalMappings.push(m);
    }
  });

  // テキストから発送先候補を検出（優先表示用）
  let priorityShippingTos = [];
  if (text && shippingToList) {
    shippingToList.forEach(s => {
      const names = [s.companyName, s.personName].filter(n => n);
      names.forEach(name => {
        if (name && text.includes(name) && mappingByShippingTo[name]) {
          if (!priorityShippingTos.includes(name)) {
            priorityShippingTos.push(name);
          }
        }
      });
    });
  }

  // 顧客候補も検出（後方互換）
  let priorityCustomers = [];
  if (text && customerList) {
    customerList.forEach(c => {
      const names = [c.companyName, c.displayName, c.personName].filter(n => n);
      names.forEach(name => {
        if (name && text.includes(name) && mappingByCustomer[name]) {
          if (!priorityCustomers.includes(name)) {
            priorityCustomers.push(name);
          }
        }
      });
    });
  }

  let mappingText = '\n═══════════════════════════════════════════════════════════\n';
  mappingText += '【学習済みマッピング】※過去に登録した表記ゆれ\n';
  mappingText += '═══════════════════════════════════════════════════════════\n';

  // 優先発送先のマッピングを先に表示
  if (priorityShippingTos.length > 0) {
    mappingText += '\n★ この発送先で過去に使用されたマッピング ★\n';
    priorityShippingTos.forEach(shippingToName => {
      mappingText += `\n【${shippingToName}】\n`;
      mappingByShippingTo[shippingToName].slice(0, 20).forEach(m => {
        mappingText += `  「${m.variant}」→「${m.productName}」\n`;
      });
    });
    mappingText += '\n---\n';
  }

  // 顧客別マッピング（後方互換）
  if (priorityCustomers.length > 0) {
    mappingText += '\n【顧客別マッピング】\n';
    priorityCustomers.forEach(customerName => {
      mappingText += `\n【${customerName}】\n`;
      mappingByCustomer[customerName].slice(0, 10).forEach(m => {
        mappingText += `  「${m.variant}」→「${m.productName}」\n`;
      });
    });
  }

  // 発送先名付きのマッピング（優先発送先以外）
  const otherShippingTos = Object.keys(mappingByShippingTo)
    .filter(name => !priorityShippingTos.includes(name));
  
  if (otherShippingTos.length > 0) {
    mappingText += '\n【発送先別マッピング】\n';
    otherShippingTos.slice(0, 10).forEach(shippingToName => {
      mappingText += `  ${shippingToName}: `;
      const items = mappingByShippingTo[shippingToName].slice(0, 5)
        .map(m => `「${m.variant}」→「${m.productName}」`);
      mappingText += items.join('、') + '\n';
    });
  }

  // 汎用マッピング（発送先名・顧客名なし）
  if (generalMappings.length > 0) {
    mappingText += '\n【共通マッピング】\n';
    generalMappings.slice(0, 30).forEach(m => {
      mappingText += `「${m.variant}」→「${m.productName}」\n`;
    });
  }

  return mappingText;
}

function buildPDFAnalysisPrompt(productList, mappingList, customerList, shippingToList) {
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

  // 発送先別マッピングセクションを構築
  const mappingText = buildCustomerMappingSection(mappingList, null, customerList, shippingToList);

  // 自社名を取得
  const companyName = getCompanyDisplayName();

  const today = new Date();
  const year = today.getFullYear();

  return `あなたは受注データ解析の専門家です。
この画像/PDFは発注書、注文書、出荷指示書、FAXのいずれかです。
内容を正確に読み取り、以下のJSON形式で返してください。

═══════════════════════════════════════════════════════════
【重要: 顧客（注文者）の特定ルール】
═══════════════════════════════════════════════════════════

1. **「○○御中」「○○様」は受注会社（あなた）であり、顧客ではありません**
   - 特に「${companyName}」は受注会社（自社）です。絶対に顧客として抽出しないでください。
   - 「御中」「様」が付いている企業名は必ず受信先（自社）です。

2. **顧客（注文者）を特定する方法**（優先順位順）:
   a) 「依頼先」「発注元」「送信元」として明記されている企業
   b) FAXの送信元ヘッダー（「From: ○○」など）
   c) 本文の最初の挨拶文に登場する企業（例: 「株式会社○○です」「○○よりご注文」）
   d) 署名欄・名刺情報に記載されている企業
   e) 連絡先（TEL/FAX）とセットで記載されている企業名

3. **顧客として扱わないもの**:
   - 受注会社（${companyName}）
   - 「御中」「様」が付いている企業名

4. **発送先の扱い**:
   - 「届け先」「配送先」「納入先」「納品先」「送り先」「センター」「倉庫」はすべて shippingTo に入れる
   - 倉庫名や配送センター名も発送先として正しい情報です

5. **顧客情報が不明な場合**:
   - customer フィールドは null にする
   - alerts に「顧客情報が特定できませんでした」と記載

【受注会社（自社）= 絶対に顧客にしないこと】
${companyName}

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
    "rawCompanyName": "注文者の会社名（送信元企業）",
    "rawDepartment": "部署名",
    "rawPersonName": "担当者名",
    "rawTel": "電話番号",
    "rawFax": "FAX番号"
  },

  "shippingTo": {
    "rawCompanyName": "発送先の会社名・店舗名・倉庫名",
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
}

【最終確認】
- customer.rawCompanyName に「${companyName}」が入っていないか確認
- 「御中」「様」付きの企業名が customer に入っていないか確認

上記に該当する場合は、customer を null に修正してください。`;
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

  // 発送先別マッピングセクションを構築（テキスト内の発送先名を検出して優先表示）
  const mappingText = buildCustomerMappingSection(mappingList, text, customerList, shippingToList);

  // 自社名を取得
  const companyName = getCompanyDisplayName();

  const today = new Date();
  const year = today.getFullYear();

  return `あなたは受注データ解析の専門家です。
以下のテキストはメール、チャット、またはFAXの内容です。
注文内容を抽出してJSON形式で返してください。

【重要: 顧客（注文者）の特定ルール】

1. **「○○御中」「○○様」は受注会社（あなた）であり、顧客ではありません**
   - 特に「${companyName}」は受注会社（自社）です。絶対に顧客として抽出しないでください。
   - 「御中」「様」が付いている企業名は必ず受信先（自社）です。

2. **顧客（注文者）を特定する方法**（優先順位順）:
   a) 「依頼先」「発注元」「送信元」として明記されている企業
   b) FAXの送信元ヘッダー（「From: ○○」など）
   c) 本文の最初の挨拶文に登場する企業（例: 「株式会社○○です」「○○よりご注文」）
   d) 署名欄・名刺情報に記載されている企業
   e) 連絡先（TEL/FAX）とセットで記載されている企業名

3. **顧客として扱わないもの**:
   - 受注会社（${companyName}）
   - 「御中」「様」が付いている企業名

4. **発送先の扱い**:
   - 「届け先」「配送先」「納入先」「納品先」「送り先」「センター」「倉庫」はすべて shippingTo に入れる
   - 倉庫名や配送センター名も発送先として正しい情報です

5. **顧客情報が不明な場合**:
   - customer フィールドは null にする
   - alerts に「顧客情報が特定できませんでした」と記載

【受注会社（自社）= 絶対に顧客にしないこと】
${companyName}

【商品マスタ】
${productMasterText}
${mappingText}

【本日】${year}年

【注文内容】
${text}

【出力形式】※JSONのみ、説明文不要
{
  "success": true,
  "shippingDate": "YYYY-MM-DD",
  "deliveryDate": "YYYY-MM-DD",
  "deliveryTime": "時間指定",
  "customer": {
    "rawCompanyName": "注文者の会社名（送信元企業）",
    "rawPersonName": "注文者の氏名（担当者名）",
    "rawTel": "電話番号",
    "rawFax": "FAX番号"
  },
  "shippingTo": {
    "rawCompanyName": "発送先の会社名・店舗名・倉庫名",
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
  "alerts": ["警告メッセージ"],
  "overallConfidence": "high/medium/low"
}

【最終確認】
- customer.rawCompanyName に「${companyName}」が入っていないか確認
- 「御中」「様」付きの企業名が customer に入っていないか確認

上記に該当する場合は、customer を null に修正してください。`;
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

    // マスタスプレッドシートをIDで直接開く（WebアプリではgetActiveSpreadsheetがnullになるため）
    const masterSpreadsheetId = getMasterSpreadsheetId();
    if (!masterSpreadsheetId) {
      Logger.log('MASTER_SPREADSHEET_ID が未設定です');
      return analysisResultJson;
    }
    
    const ss = SpreadsheetApp.openById(masterSpreadsheetId);
    
    // 顧客マスタ取得
    const customerSheet = ss.getSheetByName('顧客情報');
    const customerList = [];
    if (customerSheet) {
      const customerData = customerSheet.getDataRange().getValues();
      const headers = customerData[0];
      for (let i = 1; i < customerData.length; i++) {
        const row = customerData[i];
        const record = {};
        headers.forEach((h, idx) => { record[h] = row[idx]; });
        customerList.push({
          displayName: record['表示名'] || record['会社名'],
          companyName: record['会社名'],
          personName: record['氏名'],
          zipcode: record['郵便番号'],
          address1: record['住所１'],
          address2: record['住所２'],
          tel: record['TEL'] || record['電話番号'],
          fax: record['FAX'] || record['ＦＡＸ'],
          email: record['メールアドレス']
        });
      }
    }

    // 発送先マスタ取得
    const shippingToSheet = ss.getSheetByName('発送先情報');
    const shippingToList = [];
    if (shippingToSheet) {
      const shippingToData = shippingToSheet.getDataRange().getValues();
      const headers = shippingToData[0];
      for (let i = 1; i < shippingToData.length; i++) {
        const row = shippingToData[i];
        const record = {};
        headers.forEach((h, idx) => { record[h] = row[idx]; });
        shippingToList.push({
          companyName: record['会社名'],
          personName: record['氏名'],
          zipcode: record['郵便番号'],
          address1: record['住所１'],
          address2: record['住所２'],
          tel: record['TEL'] || record['電話番号']
        });
      }
    }

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

  // 自社名を取得
  const companyName = getCompanyDisplayName();
  const normalizedCompanyName = normalizeString(companyName);

  // 顧客照合
  if (result.customer) {
    // まず自社名チェック（最終防御層）
    const rawCompany = normalizeString(result.customer.rawCompanyName || '');
    if (rawCompany === normalizedCompanyName ||
        rawCompany === normalizeString(companyName + '様') ||
        rawCompany === normalizeString(companyName + '御中')) {
      // 自社名が顧客として抽出されている場合
      result.customer = null;
      result.customerMatch = 'none';
      result.customerMatchIndex = null;
      result.customerMatchScore = 0;
      result.alerts.push('⚠️ AIが自社名（' + companyName + '）を顧客として抽出しました。これは誤りです。FAXの送信元情報を確認してください。');
    } else {
      // 通常のマッチング処理
      const customerMatch = findBestMatch(result.customer, customerList, 'customer');
      if (customerMatch.match === 'exact' || customerMatch.match === 'partial') {
        const matched = customerList[customerMatch.index];

        // マッチング結果も自社名でないか再チェック
        if (normalizeString(matched.companyName) === normalizedCompanyName) {
          result.customer = null;
          result.customerMatch = 'none';
          result.customerMatchIndex = null;
          result.customerMatchScore = 0;
          result.alerts.push('⚠️ 顧客マッチング結果が自社名（' + companyName + '）でした。マッチングを無効化しました。');
        } else {
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
        }
      } else {
        result.customer.masterMatch = 'none';
        result.customer.isNewCustomer = true;
        result.alerts.push('⚠️ 新規顧客の可能性があります');
      }
    }
  }

  // 🤖 顧客推定（Phase 4: 常に実行するように変更）
  // 発送先情報がある場合、発送先→顧客マッピングから顧客を推定
  // LLMが特定した顧客と比較して、矛盾があればアラートを出す
  if (result.shippingTo) {
    try {
      const shippingToName = result.shippingTo.rawCompanyName || result.shippingTo.rawPersonName || '';
      if (shippingToName) {
        const estimation = estimateCustomerFromShippingTo(shippingToName);
        if (estimation && estimation.customer) {
          result.customerEstimation = estimation;
          Logger.log('顧客推定成功: ' + shippingToName + ' → ' + estimation.customer + ' (' + estimation.confidence + '%)');
          
          // LLMが特定した顧客と発送先マッピングからの推定を比較
          if (result.customer && result.customer.masterData) {
            const llmCustomer = normalizeString(result.customer.masterData.companyName || result.customer.masterData.displayName || '');
            const estimatedCustomer = normalizeString(estimation.customer);
            
            // 顧客名が異なる場合は警告（発送先マッピングの方が信頼性高い可能性）
            if (llmCustomer && estimatedCustomer && llmCustomer !== estimatedCustomer) {
              // 発送先マッピングの信頼度が高い場合（70%以上）
              if (estimation.confidence >= 70) {
                result.alerts.push(`⚠️ 発送先「${shippingToName}」の過去の取引先は「${estimation.customer}」です（現在の顧客: ${result.customer.masterData.companyName || result.customer.masterData.displayName}）`);
                result.customerConflict = {
                  llmCustomer: result.customer.masterData.companyName || result.customer.masterData.displayName,
                  estimatedCustomer: estimation.customer,
                  estimationConfidence: estimation.confidence
                };
              }
            }
          }
          
          // 顧客がnullまたは新規の場合は推定結果を提案
          if (!result.customer || result.customer.isNewCustomer) {
            if (estimation.confidence >= 50) {
              result.alerts.push(`💡 発送先「${shippingToName}」から顧客「${estimation.customer}」を推定しました（信頼度: ${estimation.confidence}%）`);
            }
          }
        }
      }
    } catch (error) {
      Logger.log('顧客推定エラー（処理続行）: ' + error.message);
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
  
  // 🏠 発送先がない場合は顧客を発送先として使用（依頼先=発送先のパターン）
  if (!result.shippingTo && result.customer && result.customer.masterData) {
    result.shippingTo = {
      rawCompanyName: result.customer.masterData.companyName || result.customer.rawCompanyName,
      rawPersonName: result.customer.masterData.personName || result.customer.rawPersonName,
      rawZipcode: result.customer.masterData.zipcode,
      rawAddress: result.customer.masterData.address,
      rawTel: result.customer.masterData.tel || result.customer.rawTel,
      masterMatch: 'customer_fallback',
      matchedBy: '顧客情報から設定',
      isNewShippingTo: false,
      isCustomerFallback: true
    };
    result.alerts.push('📍 発送先が指定されていないため、依頼先（顧客）を発送先として設定しました');
    Logger.log('発送先フォールバック: 顧客情報を発送先に使用');
  }

  return normalizeResult(result);
}

function findBestMatch(rawData, masterList, type) {
  const rawCompany = normalizeString(rawData.rawCompanyName || '');
  const rawPerson = normalizeString(rawData.rawPersonName || '');
  const rawTel = normalizeTel(rawData.rawTel || '');
  const rawFax = normalizeTel(rawData.rawFax || '');
  const rawZipcode = normalizeZipcode(rawData.rawZipcode || '');
  const rawAddress = normalizeString(rawData.rawAddress || '');

  // 自社名と「御中」「様」付きを除外リストに追加
  const companyName = getCompanyDisplayName();
  const excludeCompanies = [
    normalizeString(companyName),
    normalizeString(companyName + '様'),
    normalizeString(companyName + '御中')
  ];

  let bestMatch = { match: 'none', index: null, score: 0, matchedBy: '' };

  masterList.forEach((master, index) => {
    const masterCompany = normalizeString(master.companyName || '');

    // 除外リストに該当する場合はスキップ（顧客マッチングの場合のみ）
    if (type === 'customer' && excludeCompanies.some(exclude =>
      exclude && (masterCompany === exclude || masterCompany.includes(exclude))
    )) {
      return; // このマスタをスキップ
    }

    let score = 0;
    let matchedBy = [];

    const masterPerson = normalizeString(master.personName || '');
    const masterTel = normalizeTel(master.tel || '');
    const masterFax = normalizeTel(master.fax || '');
    const masterDisplayName = normalizeString(master.displayName || '');
    const masterZipcode = normalizeZipcode(master.zipcode || '');
    const masterAddress = normalizeString((master.address1 || '') + (master.address2 || ''));

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
    
    // 🆕 宛名とマスタ会社名の照合（FAXでは宛名が実際の会社名であることが多い）
    if (rawPerson && masterCompany && !matchedBy.includes('会社名') && !matchedBy.includes('会社名(部分)')) {
      if (rawPerson === masterCompany) {
        score += 75;
        matchedBy.push('宛名→会社名');
      } else if (rawPerson.includes(masterCompany) || masterCompany.includes(rawPerson)) {
        score += 35;
        matchedBy.push('宛名→会社名(部分)');
      }
    }
    
    // 🆕 会社名がマスタの住所に含まれている（川崎青果プロセスセンターのケース）
    if (rawCompany && masterAddress && masterAddress.includes(rawCompany)) {
      score += 30;
      matchedBy.push('会社名→住所');
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
    
    // 🆕 郵便番号一致（高信頼度）
    if (rawZipcode && masterZipcode && rawZipcode === masterZipcode) {
      score += 70;
      matchedBy.push('郵便番号');
    }
    
    // 🆕 住所一致・部分一致
    if (rawAddress && masterAddress) {
      if (rawAddress === masterAddress) {
        score += 60;
        matchedBy.push('住所');
      } else if (rawAddress.includes(masterAddress) || masterAddress.includes(rawAddress)) {
        score += 35;
        matchedBy.push('住所(部分)');
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

/**
 * 自社名を取得
 * @returns {string} 自社の表示名
 */
function getCompanyDisplayName() {
  // スクリプトプロパティから取得
  const companyName = PropertiesService.getScriptProperties().getProperty('COMPANY_NAME');
  if (!companyName) {
    Logger.log('警告: COMPANY_NAME プロパティが未設定です。スクリプトプロパティに設定してください。');
    return '';  // 空文字を返す（プロンプトには何も挿入されない）
  }
  return companyName;
}

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
    
    // 発送先名と顧客名を取得（発送先優先）
    const shippingToName = enhancedResult.shippingTo?.rawCompanyName || 
                           enhancedResult.shippingTo?.masterData?.companyName || '';
    const customerName = enhancedResult.customer?.rawCompanyName || '';
    
    // 商品マッピングで補完
    enhancedResult.items = enhanceItemsWithMapping(enhancedResult.items, shippingToName, customerName);
    
    // 価格情報を追加
    enhancedResult.items = addPriceToItems(enhancedResult.items, productList);
    
    return enhancedResult;
  } catch (error) {
    Logger.log('Error in analyzeOrderTextById: ' + error.message);
    return { success: false, error: error.message };
  }
}

// ====================================
// 移行関数
// ====================================

/**
 * 商品マッピングの移行：発送先名カラムを追加
 * 既存のシートに「発送先名」カラムがなければ追加する
 * 
 * 使用方法：GASエディタで migrateProductMappingAddShippingToColumn() を実行
 */
function migrateProductMappingAddShippingToColumn() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('商品マッピング');
    
    if (!sheet) {
      Logger.log('商品マッピングシートが存在しません');
      return { success: false, message: '商品マッピングシートが存在しません' };
    }
    
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    // 既に発送先名カラムがあるか確認
    if (headers.includes('発送先名')) {
      Logger.log('発送先名カラムは既に存在します');
      return { success: true, message: '発送先名カラムは既に存在します', migrated: false };
    }
    
    // 顧客名カラムの位置を確認
    const customerColIndex = headers.indexOf('顧客名');
    if (customerColIndex === -1) {
      Logger.log('顧客名カラムが見つかりません');
      return { success: false, message: '顧客名カラムが見つかりません' };
    }
    
    // 顧客名の前に発送先名カラムを挿入
    sheet.insertColumnBefore(customerColIndex + 1);
    sheet.getRange(1, customerColIndex + 1)
      .setValue('発送先名')
      .setBackground('#4a90d9')
      .setFontColor('#ffffff')
      .setFontWeight('bold');
    
    Logger.log('発送先名カラムを追加しました（列 ' + (customerColIndex + 1) + '）');
    
    return { 
      success: true, 
      message: '発送先名カラムを追加しました', 
      migrated: true,
      newColumnIndex: customerColIndex + 1
    };
    
  } catch (error) {
    Logger.log('Error in migrateProductMappingAddShippingToColumn: ' + error.message);
    return { success: false, error: error.message };
  }
}

/**
 * 商品マッピングの移行：顧客名を発送先名にコピー
 * 発送先マッピングを参照して、顧客名→発送先名の変換を試みる
 * 
 * 使用方法：GASエディタで migrateCustomerToShippingTo() を実行
 */
function migrateCustomerToShippingTo() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('商品マッピング');
    
    if (!sheet) {
      return { success: false, message: '商品マッピングシートが存在しません' };
    }
    
    // まず発送先名カラムを追加
    const addResult = migrateProductMappingAddShippingToColumn();
    if (!addResult.success) {
      return addResult;
    }
    
    // 発送先マッピングを取得
    const shippingMappingSheet = ss.getSheetByName('発送先マッピング');
    const shippingMappings = {};
    
    if (shippingMappingSheet) {
      const mappingData = shippingMappingSheet.getDataRange().getValues();
      const mappingHeaders = mappingData[0];
      const shippingColIdx = mappingHeaders.indexOf('発送先会社名');
      const customerColIdx = mappingHeaders.indexOf('顧客名');
      
      if (shippingColIdx >= 0 && customerColIdx >= 0) {
        for (let i = 1; i < mappingData.length; i++) {
          const customerName = mappingData[i][customerColIdx];
          const shippingToName = mappingData[i][shippingColIdx];
          if (customerName && shippingToName) {
            // 顧客名→発送先名の配列を作成
            if (!shippingMappings[customerName]) {
              shippingMappings[customerName] = [];
            }
            shippingMappings[customerName].push(shippingToName);
          }
        }
      }
    }
    
    // 商品マッピングを更新
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const shippingToColIdx = headers.indexOf('発送先名');
    const customerColIdx = headers.indexOf('顧客名');
    
    let migratedCount = 0;
    let skippedCount = 0;
    
    for (let i = 1; i < data.length; i++) {
      const existingShippingTo = data[i][shippingToColIdx];
      const customerName = data[i][customerColIdx];
      
      // 既に発送先名がある場合はスキップ
      if (existingShippingTo) {
        skippedCount++;
        continue;
      }
      
      // 顧客名から発送先名を推測
      if (customerName && shippingMappings[customerName]) {
        // 発送先が1つだけの場合のみ自動マッピング
        if (shippingMappings[customerName].length === 1) {
          sheet.getRange(i + 1, shippingToColIdx + 1).setValue(shippingMappings[customerName][0]);
          migratedCount++;
        } else {
          // 複数発送先がある場合は手動確認が必要
          skippedCount++;
        }
      } else {
        skippedCount++;
      }
    }
    
    Logger.log(`移行完了: ${migratedCount}件を発送先名に変換, ${skippedCount}件はスキップ`);
    
    return {
      success: true,
      message: `移行完了: ${migratedCount}件を発送先名に変換, ${skippedCount}件はスキップ（手動確認が必要）`,
      migratedCount,
      skippedCount
    };
    
  } catch (error) {
    Logger.log('Error in migrateCustomerToShippingTo: ' + error.message);
    return { success: false, error: error.message };
  }
}