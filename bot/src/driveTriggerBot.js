/**
 * Google Drive監視トリガー Bot
 *
 * 指定フォルダにPDF/画像が追加されたら自動でAI解析して仮受注シートに登録
 */

// ========== 設定 ==========
// 機密情報は GAS の ScriptProperties で管理してください

// 監視対象のGoogle DriveフォルダID
const WATCH_FOLDER_ID = PropertiesService.getScriptProperties().getProperty('WATCH_FOLDER_ID') || 'YOUR_FOLDER_ID_HERE';

// 処理済みファイルを移動するバックアップフォルダID
const BACKUP_FOLDER_ID = PropertiesService.getScriptProperties().getProperty('BACKUP_FOLDER_ID') || '';

// 処理失敗ファイルを移動するエラーフォルダID
const ERROR_FOLDER_ID = PropertiesService.getScriptProperties().getProperty('ERROR_FOLDER_ID') || '';

// LINE Botプロジェクトのスプレッドシート
const LINE_BOT_SPREADSHEET_ID = PropertiesService.getScriptProperties().getProperty('LINE_BOT_SPREADSHEET_ID') || '';

// マスタスプレッドシート（商品・顧客・発送先マスタがあるシート）
const MASTER_SPREADSHEET_ID = PropertiesService.getScriptProperties().getProperty('MASTER_SPREADSHEET_ID') || '';

// Gemini API設定
const GEMINI_API_KEY = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
const GEMINI_VISION_API_URL = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash-exp:generateContent';

// ========== トリガー設定用関数 ==========

/**
 * 時間ベーストリガーを設定（手動実行）
 * 5分ごとに checkNewFiles を実行
 */
function setupDriveTrigger() {
  // 既存のトリガーを削除
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'checkNewFiles') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // 新しいトリガーを作成（5分ごと）
  ScriptApp.newTrigger('checkNewFiles')
    .timeBased()
    .everyMinutes(5)
    .create();

  Logger.log('トリガーを設定しました: checkNewFiles を5分ごとに実行');
}

/**
 * トリガーを削除（手動実行）
 */
function removeDriveTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'checkNewFiles') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  Logger.log('トリガーを削除しました');
}

// ========== メイン処理 ==========

/**
 * 新規ファイルをチェックして処理
 * トリガーから5分ごとに実行される
 * 処理済みファイルはバックアップフォルダに移動
 */
function checkNewFiles() {
  try {
    const folder = DriveApp.getFolderById(WATCH_FOLDER_ID);
    
    // バックアップフォルダ（未設定の場合は監視フォルダ内に作成）
    const backupFolder = getOrCreateBackupFolder();
    const errorFolder = getOrCreateErrorFolder();

    // フォルダ内の全ファイルを処理（処理後は移動するので残っているのは未処理）
    const files = folder.getFiles();
    let processedCount = 0;
    let errorCount = 0;

    while (files.hasNext()) {
      const file = files.next();
      const mimeType = file.getMimeType();
      const fileName = file.getName();

      // PDF or 画像ファイルのみ処理
      if (isSupportedFileType(mimeType)) {
        Logger.log(`ファイル検出: ${fileName}`);
        
        const success = processFile(file);
        
        if (success) {
          // 処理成功: バックアップフォルダに移動
          moveFileToFolder(file, backupFolder);
          Logger.log(`バックアップに移動: ${fileName}`);
          processedCount++;
        } else {
          // 処理失敗: エラーフォルダに移動
          moveFileToFolder(file, errorFolder);
          Logger.log(`エラーフォルダに移動: ${fileName}`);
          errorCount++;
        }
      }
    }

    if (processedCount > 0 || errorCount > 0) {
      Logger.log(`処理完了: 成功${processedCount}件, 失敗${errorCount}件`);
    }

  } catch (error) {
    Logger.log('checkNewFilesエラー: ' + error.toString());
  }
}

/**
 * バックアップフォルダを取得または作成
 */
function getOrCreateBackupFolder() {
  if (BACKUP_FOLDER_ID) {
    try {
      return DriveApp.getFolderById(BACKUP_FOLDER_ID);
    } catch (e) {
      Logger.log('BACKUP_FOLDER_ID が無効です: ' + e.toString());
    }
  }
  
  // 監視フォルダ内に「処理済み」フォルダを作成
  const watchFolder = DriveApp.getFolderById(WATCH_FOLDER_ID);
  const folders = watchFolder.getFoldersByName('処理済み');
  
  if (folders.hasNext()) {
    return folders.next();
  } else {
    Logger.log('「処理済み」フォルダを作成しました');
    return watchFolder.createFolder('処理済み');
  }
}

/**
 * エラーフォルダを取得または作成
 */
function getOrCreateErrorFolder() {
  if (ERROR_FOLDER_ID) {
    try {
      return DriveApp.getFolderById(ERROR_FOLDER_ID);
    } catch (e) {
      Logger.log('ERROR_FOLDER_ID が無効です: ' + e.toString());
    }
  }
  
  // 監視フォルダ内に「エラー」フォルダを作成
  const watchFolder = DriveApp.getFolderById(WATCH_FOLDER_ID);
  const folders = watchFolder.getFoldersByName('エラー');
  
  if (folders.hasNext()) {
    return folders.next();
  } else {
    Logger.log('「エラー」フォルダを作成しました');
    return watchFolder.createFolder('エラー');
  }
}

/**
 * ファイルを指定フォルダに移動
 * 
 * 注意: DriveApp.removeFile/addFileは共有ドライブや権限設定によって
 * Access deniedエラーが発生することがある。
 * file.moveTo()を使用するか、エラー時はコピー＋削除にフォールバック。
 */
function moveFileToFolder(file, targetFolder) {
  try {
    // 方法1: moveTo()を使用（GAS の新しいAPI）
    file.moveTo(targetFolder);
  } catch (error) {
    Logger.log('ファイル移動エラー: ' + error.toString());
    
    // 方法2: フォールバック - コピーして元ファイルをゴミ箱に移動
    try {
      const newFile = file.makeCopy(file.getName(), targetFolder);
      file.setTrashed(true);
      Logger.log('フォールバック方法でファイルを移動しました: ' + file.getName());
    } catch (fallbackError) {
      Logger.log('フォールバック移動も失敗: ' + fallbackError.toString());
    }
  }
}

/**
 * ファイルを処理（AI解析 → 仮受注登録）
 * @param {File} file - 処理対象のファイル
 * @returns {boolean} 処理成功ならtrue、失敗ならfalse
 */
function processFile(file) {
  try {
    const fileName = file.getName();
    const mimeType = file.getMimeType();

    // ファイルをBase64エンコード
    const blob = file.getBlob();
    const base64 = Utilities.base64Encode(blob.getBytes());

    // AI解析実行
    const analysisResult = analyzeFileWithGemini(base64, mimeType, fileName);

    if (!analysisResult || !analysisResult.success) {
      Logger.log(`AI解析失敗: ${fileName}`);
      return false;
    }

    // 仮受注シートに登録
    registerTempOrder(analysisResult, fileName, file.getUrl());

    Logger.log(`処理完了: ${fileName}`);
    return true;

  } catch (error) {
    Logger.log(`ファイル処理エラー(${file.getName()}): ${error.toString()}`);
    return false;
  }
}

/**
 * Gemini AIでファイルを解析（マスタデータ参照版）
 */
function analyzeFileWithGemini(base64, mimeType, fileName) {
  try {
    // マスタデータを取得
    const masterData = getMasterDataForTrigger();
    
    // 詳細なプロンプトを構築
    const prompt = buildTriggerAnalysisPrompt(masterData);

    const payload = {
      contents: [{
        parts: [
          { text: prompt },
          {
            inline_data: {
              mime_type: mimeType,
              data: base64
            }
          }
        ]
      }],
      generationConfig: {
        temperature: 0.1,
        topK: 1,
        topP: 0.95,
        maxOutputTokens: 8192,
      }
    };

    const options = {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(
      `${GEMINI_VISION_API_URL}?key=${GEMINI_API_KEY}`,
      options
    );

    const result = JSON.parse(response.getContentText());

    if (!result.candidates || !result.candidates[0]) {
      Logger.log('AI解析レスポンスが不正: ' + JSON.stringify(result));
      return { success: false };
    }

    const text = result.candidates[0].content.parts[0].text;

    // JSON部分を抽出（マークダウンコードブロックを除去）
    let jsonText = text.replace(/```json\n?/g, '').replace(/```\n?/g, '').trim();
    const analysisData = JSON.parse(jsonText);

    // マスタ照合を実行
    const enhancedData = enhanceWithMasterMatch(analysisData, masterData);

    return {
      success: true,
      data: enhancedData,
      rawResponse: text
    };

  } catch (error) {
    Logger.log('Gemini API エラー: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

/**
 * トリガー用のマスタデータを取得
 */
function getMasterDataForTrigger() {
  try {
    if (!MASTER_SPREADSHEET_ID) {
      Logger.log('MASTER_SPREADSHEET_ID が未設定です');
      return { productList: [], customerList: [], shippingToList: [], mappingList: [] };
    }

    const ss = SpreadsheetApp.openById(MASTER_SPREADSHEET_ID);

    // 商品マスタ
    const productSheet = ss.getSheetByName('商品');
    const productList = [];
    if (productSheet) {
      const productData = productSheet.getDataRange().getValues();
      const productHeaders = productData[0];
      const nameIdx = productHeaders.indexOf('商品名');
      const categoryIdx = productHeaders.indexOf('商品分類');
      const priceIdx = productHeaders.indexOf('価格（P)');

      for (let i = 1; i < productData.length; i++) {
        const row = productData[i];
        if (row[nameIdx]) {
          productList.push({
            name: row[nameIdx],
            category: row[categoryIdx] || '未分類',
            price: row[priceIdx] || 0
          });
        }
      }
    }

    // 商品マッピング
    const mappingList = [];
    const mappingSheet = ss.getSheetByName('商品マッピング');
    if (mappingSheet) {
      const mappingData = mappingSheet.getDataRange().getValues();
      const mappingHeaders = mappingData[0];
      const variantIdx = mappingHeaders.indexOf('顧客表記');
      const prodNameIdx = mappingHeaders.indexOf('商品名');
      const catIdx = mappingHeaders.indexOf('商品分類');
      const shippingToIdx = mappingHeaders.indexOf('発送先名');
      const customerIdx = mappingHeaders.indexOf('顧客名');

      for (let i = 1; i < mappingData.length; i++) {
        const row = mappingData[i];
        if (row[variantIdx]) {
          mappingList.push({
            variant: row[variantIdx],
            productName: row[prodNameIdx],
            category: row[catIdx],
            shippingToName: shippingToIdx >= 0 ? (row[shippingToIdx] || '') : '',
            customerName: customerIdx >= 0 ? (row[customerIdx] || '') : ''
          });
        }
      }
    }

    // 顧客マスタ
    const customerSheet = ss.getSheetByName('顧客情報');
    const customerList = [];
    if (customerSheet) {
      const customerData = customerSheet.getDataRange().getValues();
      const customerHeaders = customerData[0];
      const companyIdx = customerHeaders.indexOf('会社名');
      const personIdx = customerHeaders.indexOf('氏名');
      const telIdx = customerHeaders.indexOf('TEL');
      const faxIdx = customerHeaders.indexOf('FAX');
      const zipIdx = customerHeaders.indexOf('郵便番号');

      for (let i = 1; i < customerData.length; i++) {
        const row = customerData[i];
        if (row[companyIdx] || row[personIdx]) {
          customerList.push({
            companyName: row[companyIdx] || '',
            personName: row[personIdx] || '',
            tel: String(row[telIdx] || ''),
            fax: String(row[faxIdx] || ''),
            zipcode: String(row[zipIdx] || '')
          });
        }
      }
    }

    // 発送先マスタ
    const shippingToSheet = ss.getSheetByName('発送先情報');
    const shippingToList = [];
    if (shippingToSheet) {
      const shippingToData = shippingToSheet.getDataRange().getValues();
      const stHeaders = shippingToData[0];
      const stCompanyIdx = stHeaders.indexOf('会社名');
      const stPersonIdx = stHeaders.indexOf('氏名');
      const stTelIdx = stHeaders.indexOf('TEL');
      const stZipIdx = stHeaders.indexOf('郵便番号');

      for (let i = 1; i < shippingToData.length; i++) {
        const row = shippingToData[i];
        if (row[stCompanyIdx] || row[stPersonIdx]) {
          shippingToList.push({
            companyName: row[stCompanyIdx] || '',
            personName: row[stPersonIdx] || '',
            tel: String(row[stTelIdx] || ''),
            zipcode: String(row[stZipIdx] || '')
          });
        }
      }
    }

    Logger.log(`マスタ取得: 商品${productList.length}件, マッピング${mappingList.length}件, 顧客${customerList.length}件, 発送先${shippingToList.length}件`);

    return { productList, customerList, shippingToList, mappingList };

  } catch (error) {
    Logger.log('マスタデータ取得エラー: ' + error.toString());
    return { productList: [], customerList: [], shippingToList: [], mappingList: [] };
  }
}

/**
 * 発送先別マッピングセクションを構築
 * @param {Array} mappingList - マッピング一覧
 * @param {Array} shippingToList - 発送先一覧
 * @returns {string} プロンプト用のマッピングセクション
 */
function buildCustomerMappingSection(mappingList, shippingToList) {
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

  let mappingText = '\n═══════════════════════════════════════════════════════════\n';
  mappingText += '【学習済みマッピング】※過去に登録した表記ゆれ\n';
  mappingText += '═══════════════════════════════════════════════════════════\n';

  // 発送先名付きのマッピング（優先）
  const shippingToNames = Object.keys(mappingByShippingTo);
  if (shippingToNames.length > 0) {
    mappingText += '\n【発送先別マッピング】★この発送先のFAXを解析する際は優先的に使用★\n';
    shippingToNames.slice(0, 15).forEach(shippingToName => {
      mappingText += `\n【${shippingToName}】\n`;
      mappingByShippingTo[shippingToName].slice(0, 15).forEach(m => {
        mappingText += `  「${m.variant}」→「${m.productName}」\n`;
      });
    });
  }

  // 顧客名付きのマッピング（後方互換）
  const customerNames = Object.keys(mappingByCustomer);
  if (customerNames.length > 0) {
    mappingText += '\n【顧客別マッピング】\n';
    customerNames.slice(0, 10).forEach(customerName => {
      mappingText += `  ${customerName}: `;
      const items = mappingByCustomer[customerName].slice(0, 5)
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

/**
 * トリガー用の詳細プロンプトを構築
 */
function buildTriggerAnalysisPrompt(masterData) {
  const { productList, mappingList, customerList, shippingToList } = masterData;

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
  const mappingText = buildCustomerMappingSection(mappingList, shippingToList);

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

/**
 * 解析結果にマスタ照合を追加
 */
function enhanceWithMasterMatch(analysisData, masterData) {
  const { customerList, shippingToList } = masterData;

  // 警告配列の初期化
  if (!analysisData.alerts) {
    analysisData.alerts = [];
  }

  // 自社名を取得
  const companyName = getCompanyDisplayName();
  const normalizedCompanyName = normalizeStringForMatch(companyName);

  // 顧客照合
  if (analysisData.customer) {
    // まず自社名チェック（最終防御層）
    const rawCompany = normalizeStringForMatch(analysisData.customer.rawCompanyName || '');
    if (rawCompany === normalizedCompanyName ||
        rawCompany === normalizeStringForMatch(companyName + '様') ||
        rawCompany === normalizeStringForMatch(companyName + '御中')) {
      // 自社名が顧客として抽出されている場合
      analysisData.customer = null;
      analysisData.alerts.push('⚠️ AIが自社名（' + companyName + '）を顧客として抽出しました。これは誤りです。FAXの送信元情報を確認してください。');
    } else {
      // 通常のマッチング処理
      const customerMatch = findBestCustomerMatch(analysisData.customer, customerList);

      // マッチング結果も自社名でないか再チェック
      if (customerMatch.match !== 'none' && customerList.length > 0) {
        const matched = customerList[Math.floor(customerMatch.score / 80)]; // スコアからインデックス推定
        if (matched && normalizeStringForMatch(matched.companyName || '') === normalizedCompanyName) {
          analysisData.customer = null;
          analysisData.alerts.push('⚠️ 顧客マッチング結果が自社名（' + companyName + '）でした。マッチングを無効化しました。');
        } else {
          analysisData.customer.masterMatch = customerMatch.match;
          analysisData.customer.matchedBy = customerMatch.matchedBy;
        }
      } else {
        analysisData.customer.masterMatch = customerMatch.match;
        analysisData.customer.matchedBy = customerMatch.matchedBy;
      }
    }
  }

  // 発送先照合
  if (analysisData.shippingTo) {
    const shippingToMatch = findBestShippingToMatch(analysisData.shippingTo, shippingToList);
    analysisData.shippingTo.masterMatch = shippingToMatch.match;
    analysisData.shippingTo.matchedBy = shippingToMatch.matchedBy;
  }

  return analysisData;
}

/**
 * 顧客マスタとの照合（受注画面と統一）
 */
function findBestCustomerMatch(rawCustomer, customerList) {
  const rawCompany = normalizeStringForMatch(rawCustomer.rawCompanyName || '');
  const rawPerson = normalizeStringForMatch(rawCustomer.rawPersonName || '');
  const rawTel = normalizeTelForMatch(rawCustomer.rawTel || '');
  const rawFax = normalizeTelForMatch(rawCustomer.rawFax || '');

  // 自社名と「御中」「様」付きを除外リストに追加
  const companyName = getCompanyDisplayName();
  const excludeCompanies = [
    normalizeStringForMatch(companyName),
    normalizeStringForMatch(companyName + '様'),
    normalizeStringForMatch(companyName + '御中')
  ];

  let bestMatch = { match: 'none', score: 0, matchedBy: '' };

  customerList.forEach((master) => {
    const masterCompany = normalizeStringForMatch(master.companyName || '');

    // 除外リストに該当する場合はスキップ
    if (excludeCompanies.some(exclude =>
      exclude && (masterCompany === exclude || masterCompany.includes(exclude))
    )) {
      return; // このマスタをスキップ
    }

    let score = 0;
    let matchedBy = [];

    const masterPerson = normalizeStringForMatch(master.personName || '');
    const masterTel = normalizeTelForMatch(master.tel || '');
    const masterFax = normalizeTelForMatch(master.fax || '');
    const masterDisplayName = normalizeStringForMatch(master.displayName || '');

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
        score: score,
        matchedBy: matchedBy.join(', ')
      };
    }
  });

  if (bestMatch.score < 40) {
    bestMatch.match = 'none';
  }

  return bestMatch;
}

/**
 * 発送先マスタとの照合
 */
function findBestShippingToMatch(rawShippingTo, shippingToList) {
  const rawCompany = normalizeStringForMatch(rawShippingTo.rawCompanyName || '');
  const rawPerson = normalizeStringForMatch(rawShippingTo.rawPersonName || '');
  const rawTel = normalizeTelForMatch(rawShippingTo.rawTel || '');
  const rawFax = normalizeTelForMatch(rawShippingTo.rawFax || '');
  const rawZipcode = normalizeZipcodeForMatch(rawShippingTo.rawZipcode || '');
  const rawAddress = normalizeStringForMatch(rawShippingTo.rawAddress || '');

  let bestMatch = { match: 'none', score: 0, matchedBy: '' };

  shippingToList.forEach((master) => {
    let score = 0;
    let matchedBy = [];

    const masterCompany = normalizeStringForMatch(master.companyName || '');
    const masterPerson = normalizeStringForMatch(master.personName || '');
    const masterTel = normalizeTelForMatch(master.tel || '');
    const masterFax = normalizeTelForMatch(master.fax || '');
    const masterDisplayName = normalizeStringForMatch(master.displayName || '');
    const masterZipcode = normalizeZipcodeForMatch(master.zipcode || '');
    const masterAddress = normalizeStringForMatch((master.address1 || '') + (master.address2 || ''));

    // 電話番号一致
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
        score: score,
        matchedBy: matchedBy.join(', ')
      };
    }
  });

  if (bestMatch.score < 40) {
    bestMatch.match = 'none';
  }

  return bestMatch;
}

/**
 * 郵便番号正規化（照合用）
 */
function normalizeZipcodeForMatch(zipcode) {
  if (!zipcode) return '';
  const normalized = String(zipcode).replace(/[^0-9]/g, '');
  return normalized.length === 7 ? normalized : '';
}

/**
 * 文字列正規化（照合用）
 */
function normalizeStringForMatch(str) {
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

/**
 * 電話番号正規化
 */
function normalizeTelForMatch(tel) {
  if (!tel) return '';
  const normalized = String(tel).replace(/[^0-9]/g, '');
  return (normalized.length >= 10 && normalized.length <= 11) ? normalized : '';
}

/**
 * 郵便番号正規化
 */
function normalizeZipcodeForMatch(zipcode) {
  if (!zipcode) return '';
  const normalized = String(zipcode).replace(/[^0-9]/g, '');
  return normalized.length === 7 ? normalized : '';
}

/**
 * 仮受注シートに登録
 */
function registerTempOrder(analysisResult, fileName, fileUrl) {
  try {
    if (!LINE_BOT_SPREADSHEET_ID) {
      Logger.log('エラー: LINE_BOT_SPREADSHEET_ID が未設定です');
      throw new Error('LINE_BOT_SPREADSHEET_ID が未設定です');
    }

    const ss = SpreadsheetApp.openById(LINE_BOT_SPREADSHEET_ID);
    
    if (!ss) {
      Logger.log('エラー: スプレッドシートを開けませんでした');
      throw new Error('スプレッドシートを開けませんでした');
    }

    let sheet = ss.getSheetByName('仮受注');

    // シートがなければ作成
    if (!sheet) {
      sheet = createTempOrderSheet(ss);
      Logger.log('仮受注シートを作成しました');
    }

    const data = analysisResult.data;
    const tempOrderId = 'TMP' + Utilities.formatDate(new Date(), 'JST', 'yyyyMMddHHmmss');

    // 商品データをJSON化
    const itemsJson = JSON.stringify(data.items || []);

    // 顧客名抽出（rawCompanyName も確認）
    const customerName = data.customer?.rawCompanyName || data.customer?.companyName || '';

    // 発送先名抽出（rawCompanyName も確認）
    const shippingToName = data.shippingTo?.rawCompanyName || data.shippingTo?.companyName || '';

    // 解析結果全体をJSON化
    const analysisJson = JSON.stringify(analysisResult);

    // 新しい行を追加（通知済みフラグも初期化）
    sheet.appendRow([
      tempOrderId,                    // A: 仮受注ID
      new Date(),                     // B: 登録日時
      '未処理',                       // C: ステータス
      customerName,                   // D: 顧客名
      shippingToName,                 // E: 発送先
      itemsJson,                      // F: 商品データ
      fileName,                       // G: 原文（ファイル名）
      analysisJson,                   // H: 解析結果
      fileUrl,                        // I: ファイルURL
      '',                             // J: 処理日時
      false,                          // K: 通知済み（まとめ通知用）
      '',                             // L: 通知日時（まとめ通知用）
    ]);

    Logger.log(`仮受注登録完了: ${tempOrderId}`);

  } catch (error) {
    Logger.log('仮受注登録エラー: ' + error.toString());
  }
}

/**
 * 仮受注シートを作成
 */
function createTempOrderSheet(spreadsheet) {
  const sheet = spreadsheet.insertSheet('仮受注');
  
  // ヘッダー行を設定
  const headers = [
    '仮受注ID',     // A
    '登録日時',     // B
    'ステータス',   // C
    '顧客名',       // D
    '発送先',       // E
    '商品データ',   // F
    '原文',         // G
    '解析結果',     // H
    'ファイルURL',  // I
    '処理日時',     // J
    '通知済み',     // K
    '通知日時'      // L
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // ヘッダー行のスタイル設定
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#4a86e8');
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');
  
  // 列幅を調整
  sheet.setColumnWidth(1, 180);  // 仮受注ID
  sheet.setColumnWidth(2, 150);  // 登録日時
  sheet.setColumnWidth(3, 80);   // ステータス
  sheet.setColumnWidth(4, 150);  // 顧客名
  sheet.setColumnWidth(5, 150);  // 発送先
  sheet.setColumnWidth(6, 300);  // 商品データ
  sheet.setColumnWidth(7, 150);  // 原文
  sheet.setColumnWidth(8, 300);  // 解析結果
  sheet.setColumnWidth(9, 200);  // ファイルURL
  sheet.setColumnWidth(10, 150); // 処理日時
  sheet.setColumnWidth(11, 80);  // 通知済み
  sheet.setColumnWidth(12, 150); // 通知日時
  
  // 行を固定
  sheet.setFrozenRows(1);
  
  return sheet;
}

// ========== ユーティリティ関数 ==========

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

/**
 * 対応ファイルタイプか確認
 */
function isSupportedFileType(mimeType) {
  const supported = [
    'application/pdf',
    'image/png',
    'image/jpeg',
    'image/jpg',
    'image/gif',
    'image/webp'
  ];
  return supported.includes(mimeType);
}

/**
 * 最終チェック時刻を取得
 */
function getLastCheckTime() {
  const props = PropertiesService.getScriptProperties();
  const lastCheck = props.getProperty('LAST_CHECK_TIME');

  if (lastCheck) {
    return new Date(lastCheck);
  } else {
    // 初回は現在時刻の10分前
    return new Date(Date.now() - 10 * 60 * 1000);
  }
}

/**
 * 最終チェック時刻を保存
 */
function saveLastCheckTime(time) {
  const props = PropertiesService.getScriptProperties();
  props.setProperty('LAST_CHECK_TIME', time.toISOString());
}

/**
 * 手動テスト用: 指定ファイルを処理
 */
function testProcessFile(fileId) {
  const file = DriveApp.getFileById(fileId);
  processFile(file);
}

// ========== データ移行用関数 ==========

/**
 * 既存の「未確定」ステータスを「未処理」に一括変換
 * ※初回セットアップ時に1回だけ実行してください
 */
function migrateStatusToUnprocessed() {
  try {
    const ss = SpreadsheetApp.openById(LINE_BOT_SPREADSHEET_ID);
    const sheet = ss.getSheetByName('仮受注');

    if (!sheet) {
      Logger.log('仮受注シートが見つかりません');
      return;
    }

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const idxStatus = headers.indexOf('ステータス');

    if (idxStatus === -1) {
      Logger.log('ステータス列が見つかりません');
      return;
    }

    let updateCount = 0;

    // データ行をチェック（ヘッダー行をスキップ）
    for (let i = 1; i < data.length; i++) {
      const currentStatus = data[i][idxStatus];

      // 「未確定」または空白を「未処理」に変換
      if (currentStatus === '未確定' || currentStatus === '' || !currentStatus) {
        sheet.getRange(i + 1, idxStatus + 1).setValue('未処理');
        updateCount++;
      }
    }

    Logger.log(`ステータス移行完了: ${updateCount}件を「未処理」に更新しました`);
    return { success: true, count: updateCount };

  } catch (error) {
    Logger.log('ステータス移行エラー: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}
