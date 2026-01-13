/**
 * 発送先マッピングシステム
 * AI解析で認識した発送先名から顧客を自動推定するための学習データを管理
 *
 * データ構造:
 * - AI認識発送先名: FAXから読み取った発送先名（表記ゆれを含む）
 * - マスタ発送先名: 正規化された発送先名
 * - 顧客名: この発送先に紐づく顧客
 * - 信頼度: この組み合わせの出現回数（使用頻度）
 * - 最終使用日: 最後に使用された日付
 * - 作成日時: レコード作成日時
 */

/**
 * 発送先マッピングシートを作成（存在しない場合のみ）
 *
 * AI自動学習システム（Phase 1）の基盤となるシートを作成します。
 * 発送先名と顧客の対応関係を学習データとして蓄積します。
 *
 * シート構造:
 * - AI認識発送先名: FAXから読み取った発送先名（表記ゆれを含む）
 * - マスタ発送先名: 正規化された発送先名
 * - 顧客名: この発送先に紐づく顧客
 * - 信頼度: 使用回数（頻度が高いほど信頼性が高い）
 * - 最終使用日: 最後に使用された日付（古いデータの減衰に使用）
 * - 作成日時: レコード作成日時
 *
 * スタイル:
 * - ヘッダー: 青背景（#4285f4）、白文字、太字、中央揃え
 * - ヘッダー行を固定（スクロール時も表示）
 *
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} - 作成または既存のシート
 * @see recordShippingMapping() - マッピングデータ記録関数
 */
function createShippingMappingSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = '発送先マッピング';

  // 既にシートが存在する場合はそのまま返す
  let sheet = ss.getSheetByName(sheetName);
  if (sheet) {
    Logger.log('発送先マッピングシートは既に存在します');
    return sheet;
  }

  // 新規シートを作成
  sheet = ss.insertSheet(sheetName);

  // ヘッダー行を設定
  const headers = [
    'AI認識発送先名',
    'マスタ発送先名',
    '顧客名',
    '信頼度',
    '最終使用日',
    '作成日時'
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // ヘッダー行のスタイル設定
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#4285f4');
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');
  headerRange.setHorizontalAlignment('center');

  // 列幅を調整
  sheet.setColumnWidth(1, 200); // AI認識発送先名
  sheet.setColumnWidth(2, 200); // マスタ発送先名
  sheet.setColumnWidth(3, 150); // 顧客名
  sheet.setColumnWidth(4, 80);  // 信頼度
  sheet.setColumnWidth(5, 120); // 最終使用日
  sheet.setColumnWidth(6, 150); // 作成日時

  // シートを固定
  sheet.setFrozenRows(1);

  Logger.log('発送先マッピングシートを作成しました');
  return sheet;
}

/**
 * 発送先マッピングデータを記録または更新（Phase 1-2: 学習データ蓄積）
 *
 * 受注完了時に自動的に呼び出され、発送先名と顧客の対応を学習します。
 * 同じ組み合わせが再度使用されると信頼度が自動的に上昇します。
 *
 * 処理ロジック:
 * 1. 既存レコード検索（AI認識発送先名 + マスタ発送先名 + 顧客名）
 * 2. 見つかった場合:
 *    - 信頼度を+1（使用頻度カウント）
 *    - 最終使用日を現在日時に更新
 * 3. 見つからない場合:
 *    - 新規レコードを作成（初期信頼度: 1）
 *
 * 自動呼び出し元:
 * - orderCode.js の受注完了処理（約1942行目）
 *
 * データ例:
 * ```
 * AI認識発送先名: "東京青果市場 太田支店"
 * マスタ発送先名: "東京青果市場太田支店"
 * 顧客名: "山田商店"
 * 信頼度: 5 (5回使用)
 * ```
 *
 * @param {string} aiRecognizedName - AI認識した発送先名（表記ゆれを含む）
 * @param {string} masterName - マスタの発送先名（正規化済み）
 * @param {string} customerName - 顧客名
 * @returns {Object} - 処理結果 {success, action: 'created'|'updated', confidence}
 * @see createShippingMappingSheet() - シート作成
 * @see estimateCustomerFromShippingTo() - 推定時に使用
 */
function recordShippingMapping(aiRecognizedName, masterName, customerName) {
  if (!aiRecognizedName || !masterName || !customerName) {
    return { success: false, error: '必須パラメータが不足しています' };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('発送先マッピング');

  // シートが存在しない場合は作成
  if (!sheet) {
    sheet = createShippingMappingSheet();
  }

  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  // 列インデックス
  const aiNameCol = 0;
  const masterNameCol = 1;
  const customerNameCol = 2;
  const confidenceCol = 3;
  const lastUsedCol = 4;
  const createdCol = 5;

  const now = new Date();

  // 既存レコードを検索（AI認識発送先名 + マスタ発送先名 + 顧客名の組み合わせ）
  let foundRow = -1;
  for (let i = 1; i < data.length; i++) {
    if (data[i][aiNameCol] === aiRecognizedName &&
        data[i][masterNameCol] === masterName &&
        data[i][customerNameCol] === customerName) {
      foundRow = i + 1; // 1-indexed
      break;
    }
  }

  if (foundRow > 0) {
    // 既存レコードを更新（信頼度+1、最終使用日更新）
    const currentConfidence = Number(data[foundRow - 1][confidenceCol]) || 0;
    sheet.getRange(foundRow, confidenceCol + 1).setValue(currentConfidence + 1);
    sheet.getRange(foundRow, lastUsedCol + 1).setValue(now);

    Logger.log(`発送先マッピング更新: ${aiRecognizedName} → ${customerName} (信頼度: ${currentConfidence + 1})`);
    return { success: true, action: 'updated', confidence: currentConfidence + 1 };
  } else {
    // 新規レコードを追加
    const newRow = [
      aiRecognizedName,
      masterName,
      customerName,
      1, // 初期信頼度
      now, // 最終使用日
      now  // 作成日時
    ];

    sheet.appendRow(newRow);

    Logger.log(`発送先マッピング新規作成: ${aiRecognizedName} → ${customerName}`);
    return { success: true, action: 'created', confidence: 1 };
  }
}

/**
 * 発送先名から顧客を推定（マッピングデータから検索、Phase 3）
 *
 * 学習済みのマッピングデータから、発送先名に対応する顧客候補を検索します。
 * 正規化+柔軟なマッチングにより、表記ゆれに対応します。
 *
 * マッチングロジック:
 * 1. 入力発送先名を正規化（空白除去、全角→半角、小文字化）
 * 2. マッピングシートの各レコードと比較:
 *    - 完全一致: normalizedAiName === normalizedInput
 *    - 部分一致: どちらかが他方を含む
 * 3. 候補を信頼度と最終使用日でソート
 *
 * 正規化処理:
 * - 空白除去: "東京　青果" → "東京青果"
 * - 全角→半角: "ＡＢＣ" → "ABC"
 * - ハイフン統一: "‐－―−" → "-"
 * - 小文字化: "ABC" → "abc"
 *
 * 戻り値の構造:
 * ```javascript
 * [{
 *   customer: "山田商店",
 *   confidence: 5,           // 使用回数
 *   lastUsed: Date,          // 最終使用日
 *   source: 'mapping',       // データソース
 *   matchType: 'exact'       // 'exact' or 'partial'
 * }]
 * ```
 *
 * @param {string} shippingToName - 発送先名
 * @returns {Array} - 候補リスト（信頼度降順、同率なら最近使用順）
 * @see normalizeShippingName() - 発送先名正規化処理
 * @see estimateCustomerFromShippingTo() - 統合推定処理
 */
function getCustomerFromMapping(shippingToName) {
  if (!shippingToName) return [];

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('発送先マッピング');

  if (!sheet) {
    Logger.log('発送先マッピングシートが存在しません');
    return [];
  }

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) {
    Logger.log('発送先マッピングデータが存在しません');
    return [];
  }

  const headers = data[0];
  const aiNameCol = 0;
  const masterNameCol = 1;
  const customerNameCol = 2;
  const confidenceCol = 3;
  const lastUsedCol = 4;

  const candidates = [];
  const normalizedInput = normalizeShippingName(shippingToName);

  // 完全一致または部分一致で検索
  for (let i = 1; i < data.length; i++) {
    const aiName = data[i][aiNameCol] || '';
    const masterName = data[i][masterNameCol] || '';
    const customerName = data[i][customerNameCol] || '';
    const confidence = Number(data[i][confidenceCol]) || 0;
    const lastUsed = data[i][lastUsedCol];

    // 正規化して比較
    const normalizedAiName = normalizeShippingName(aiName);
    const normalizedMasterName = normalizeShippingName(masterName);

    // 完全一致または部分一致
    if (normalizedAiName === normalizedInput ||
        normalizedMasterName === normalizedInput ||
        normalizedInput.includes(normalizedAiName) ||
        normalizedAiName.includes(normalizedInput)) {

      candidates.push({
        customer: customerName,
        confidence: confidence,
        lastUsed: lastUsed,
        source: 'mapping',
        matchType: normalizedAiName === normalizedInput ? 'exact' : 'partial'
      });
    }
  }

  // 信頼度と最終使用日でソート（信頼度優先、同じなら最近使用）
  candidates.sort((a, b) => {
    if (b.confidence !== a.confidence) {
      return b.confidence - a.confidence;
    }
    return new Date(b.lastUsed) - new Date(a.lastUsed);
  });

  Logger.log(`マッピングから${candidates.length}件の候補を発見: ${shippingToName}`);
  return candidates;
}

/**
 * 発送先名を正規化（空白除去、全角→半角変換）
 * @param {string} name - 発送先名
 * @returns {string} - 正規化された名前
 */
function normalizeShippingName(name) {
  if (!name) return '';

  return String(name)
    .replace(/\s+/g, '')  // 全ての空白を除去
    .replace(/[Ａ-Ｚａ-ｚ０-９]/g, function(s) {
      // 全角英数字を半角に変換
      return String.fromCharCode(s.charCodeAt(0) - 0xFEE0);
    })
    .replace(/[‐－―−]/g, '-')  // 各種ハイフンを統一
    .toLowerCase();  // 小文字に統一
}

/**
 * 受注履歴から顧客を推定
 * @param {string} shippingToName - 発送先名
 * @param {number} maxResults - 最大候補数（デフォルト: 5）
 * @returns {Array} - 候補リスト [{customer, confidence, frequency, lastOrderDate, source: 'order_history'}]
 */
function getCustomerFromOrderHistory(shippingToName, maxResults = 5) {
  if (!shippingToName) return [];

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('受注');

  if (!sheet) {
    Logger.log('受注シートが存在しません');
    return [];
  }

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) {
    Logger.log('受注データが存在しません');
    return [];
  }

  const headers = data[0];
  const shippingToCol = headers.indexOf('発送先名');
  const customerCol = headers.indexOf('顧客名');
  const orderDateCol = headers.indexOf('受注日');

  if (shippingToCol === -1 || customerCol === -1) {
    Logger.log('必要な列が見つかりません');
    return [];
  }

  const normalizedInput = normalizeShippingName(shippingToName);
  const customerMap = new Map(); // 顧客名 → {count, lastOrderDate}

  // 受注履歴から発送先名が一致するレコードを検索
  for (let i = 1; i < data.length; i++) {
    const shippingTo = data[i][shippingToCol] || '';
    const customer = data[i][customerCol] || '';
    const orderDate = data[i][orderDateCol];

    if (!customer) continue;

    const normalizedShippingTo = normalizeShippingName(shippingTo);

    // 完全一致または部分一致
    if (normalizedShippingTo === normalizedInput ||
        normalizedInput.includes(normalizedShippingTo) ||
        normalizedShippingTo.includes(normalizedInput)) {

      if (!customerMap.has(customer)) {
        customerMap.set(customer, {
          count: 0,
          lastOrderDate: orderDate
        });
      }

      const entry = customerMap.get(customer);
      entry.count++;

      // より新しい受注日で更新
      if (orderDate && (!entry.lastOrderDate || new Date(orderDate) > new Date(entry.lastOrderDate))) {
        entry.lastOrderDate = orderDate;
      }
    }
  }

  // Map を配列に変換
  const candidates = Array.from(customerMap.entries()).map(([customer, data]) => {
    // 時系列重み付け: 最近の受注ほど高いスコア
    let timeWeight = 1.0;
    if (data.lastOrderDate) {
      const daysSinceOrder = (new Date() - new Date(data.lastOrderDate)) / (1000 * 60 * 60 * 24);
      // 30日以内: 1.5倍、60日以内: 1.2倍、それ以降: 1.0倍
      if (daysSinceOrder <= 30) {
        timeWeight = 1.5;
      } else if (daysSinceOrder <= 60) {
        timeWeight = 1.2;
      }
    }

    // 信頼度 = 出現回数 × 時系列重み
    const confidence = data.count * timeWeight;

    return {
      customer: customer,
      confidence: Math.round(confidence * 10) / 10, // 小数点1桁に丸める
      frequency: data.count,
      lastOrderDate: data.lastOrderDate,
      source: 'order_history'
    };
  });

  // 信頼度でソート（高い順）
  candidates.sort((a, b) => b.confidence - a.confidence);

  // 上位N件のみ返す
  const topCandidates = candidates.slice(0, maxResults);

  Logger.log(`受注履歴から${topCandidates.length}件の候補を発見: ${shippingToName}`);
  return topCandidates;
}

/**
 * マルチソース統合: 複数の情報源から顧客を推定（Phase 3-4: AI自動学習の核心機能）
 *
 * 発送先名から顧客を推定する最も重要な関数です。
 * マッピングデータと受注履歴の両方を統合し、重み付けスコアリングで最適な顧客を推定します。
 *
 * 処理フロー:
 * 1. マッピングデータから候補を取得（getCustomerFromMapping）
 * 2. 受注履歴から候補を取得（getCustomerFromOrderHistory）
 * 3. 全候補を顧客名でグループ化
 * 4. ソースごとの重み付けスコアリング:
 *    - マッピングデータ: ×2.0（明示的な学習なので最も信頼性が高い）
 *    - 受注履歴: ×1.5（確定データなので信頼性高い）
 * 5. スコア順にソート、上位候補を選出
 * 6. 信頼度を0-100%に正規化
 *
 * 重み付けロジック:
 * ```javascript
 * スコア = Σ(各候補の信頼度 × ソース重み)
 *
 * ソース重み:
 * - mapping: 2.0  // 受注完了時の明示的な学習
 * - order_history: 1.5  // 過去の受注実績
 * ```
 *
 * 信頼度正規化:
 * ```javascript
 * 正規化信頼度 = min(100, round((スコア / (スコア + 5)) × 100))
 * // スコアが高いほど100%に近づく
 * // スコア=5で50%, スコア=15で75%, スコア=45で90%
 * ```
 *
 * 戻り値の構造:
 * ```javascript
 * {
 *   customer: "山田商店",          // 推定顧客名
 *   confidence: 85,                // 信頼度（0-100%）
 *   alternatives: [                // 代替候補（上位3件）
 *     { customer: "佐藤商店", confidence: 60, sources: ['mapping'] },
 *     { customer: "田中商店", confidence: 45, sources: ['order_history'] }
 *   ],
 *   sources: ['mapping', 'order_history'],  // データソース
 *   rawDetails: [...]              // デバッグ用の詳細情報
 * }
 * ```
 *
 * 使用場面:
 * - AI解析時の顧客自動推定（shippingAi.html: displayCustomerSuggestion）
 * - 1クリック適用機能（shippingAi.html: applyCustomerSuggestion）
 *
 * @param {string} shippingToName - 発送先名
 * @returns {Object} - 推定結果 {customer, confidence, alternatives, sources, rawDetails}
 * @see getCustomerFromMapping() - マッピングデータ検索
 * @see getCustomerFromOrderHistory() - 受注履歴検索
 * @see recordShippingMapping() - 学習データ記録（受注完了時）
 */
function estimateCustomerFromShippingTo(shippingToName) {
  if (!shippingToName) {
    return {
      customer: null,
      confidence: 0,
      alternatives: [],
      sources: []
    };
  }

  // 1. マッピングデータから検索
  const mappingCandidates = getCustomerFromMapping(shippingToName);

  // 2. 受注履歴から検索
  const historyCandidates = getCustomerFromOrderHistory(shippingToName);

  // 3. 全候補を統合
  const allCandidates = [...mappingCandidates, ...historyCandidates];

  if (allCandidates.length === 0) {
    Logger.log(`候補が見つかりませんでした: ${shippingToName}`);
    return {
      customer: null,
      confidence: 0,
      alternatives: [],
      sources: []
    };
  }

  // 顧客名でグループ化してスコアを統合
  const customerScores = new Map();

  allCandidates.forEach(candidate => {
    const customer = candidate.customer;

    if (!customerScores.has(customer)) {
      customerScores.set(customer, {
        customer: customer,
        totalConfidence: 0,
        sources: [],
        details: []
      });
    }

    const entry = customerScores.get(customer);

    // ソースごとの重み付け
    let weight = 1.0;
    if (candidate.source === 'mapping') {
      weight = 2.0; // マッピングデータは明示的な学習なので重視
    } else if (candidate.source === 'order_history') {
      weight = 1.5; // 受注履歴も信頼性高い
    }

    entry.totalConfidence += candidate.confidence * weight;
    if (!entry.sources.includes(candidate.source)) {
      entry.sources.push(candidate.source);
    }
    entry.details.push(candidate);
  });

  // スコア順にソート
  const rankedCandidates = Array.from(customerScores.values())
    .sort((a, b) => b.totalConfidence - a.totalConfidence);

  const topCandidate = rankedCandidates[0];
  const alternatives = rankedCandidates.slice(1, 4); // 上位3件を代替候補として

  // 信頼度を0-100のパーセンテージに正規化
  const maxConfidence = topCandidate.totalConfidence;
  const normalizedConfidence = Math.min(100, Math.round((maxConfidence / (maxConfidence + 5)) * 100));

  Logger.log(`推定結果: ${shippingToName} → ${topCandidate.customer} (信頼度: ${normalizedConfidence}%)`);

  return {
    customer: topCandidate.customer,
    confidence: normalizedConfidence,
    alternatives: alternatives.map(c => ({
      customer: c.customer,
      confidence: Math.min(100, Math.round((c.totalConfidence / (maxConfidence + 5)) * 100)),
      sources: c.sources
    })),
    sources: topCandidate.sources,
    rawDetails: topCandidate.details // デバッグ用
  };
}

/**
 * 初期化: 発送先マッピングシートを作成
 * 手動実行用の関数
 */
function initializeShippingMappingSystem() {
  const sheet = createShippingMappingSheet();
  Logger.log('発送先マッピングシステムの初期化が完了しました');
  return { success: true, sheetName: sheet.getName() };
}

/**
 * 初回移行: 既存の受注データから発送先マッピングを一括生成
 *
 * 【使用方法】
 * 1. GASエディタでこの関数を選択
 * 2. 実行ボタンをクリック
 * 3. 完了メッセージを確認
 *
 * 【処理内容】
 * - 受注シートの全データをスキャン
 * - 発送先名と顧客名の組み合わせをマッピングシートに記録
 * - 同じ組み合わせが複数回出現する場合は信頼度をカウント
 * - 既存のマッピングデータは保持され、新規データのみ追加
 *
 * @returns {Object} - 処理結果 {success, recordsProcessed, mappingsCreated, message}
 */
function migrateExistingOrdersToMapping() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const orderSheet = ss.getSheetByName('受注');

    if (!orderSheet) {
      return {
        success: false,
        error: '受注シートが見つかりません'
      };
    }

    // マッピングシートを取得（なければ作成）
    let mappingSheet = ss.getSheetByName('発送先マッピング');
    if (!mappingSheet) {
      mappingSheet = createShippingMappingSheet();
      Logger.log('発送先マッピングシートを新規作成しました');
    }

    const orderData = orderSheet.getDataRange().getValues();
    if (orderData.length <= 1) {
      return {
        success: false,
        error: '受注データが存在しません'
      };
    }

    const headers = orderData[0];
    const shippingToCol = headers.indexOf('発送先名');
    const customerCol = headers.indexOf('顧客名');

    if (shippingToCol === -1 || customerCol === -1) {
      return {
        success: false,
        error: '必要な列（発送先名、顧客名）が見つかりません'
      };
    }

    // 発送先→顧客の組み合わせを集計
    // Key: "発送先名|顧客名", Value: {shippingTo, customer, count}
    const mappingMap = new Map();

    let recordsProcessed = 0;
    for (let i = 1; i < orderData.length; i++) {
      const shippingTo = orderData[i][shippingToCol];
      const customer = orderData[i][customerCol];

      if (!shippingTo || !customer) continue;

      const shippingToName = String(shippingTo).trim();
      const customerName = String(customer).trim();

      if (!shippingToName || !customerName) continue;

      const key = `${shippingToName}|${customerName}`;

      if (!mappingMap.has(key)) {
        mappingMap.set(key, {
          shippingTo: shippingToName,
          customer: customerName,
          count: 0
        });
      }

      mappingMap.get(key).count++;
      recordsProcessed++;
    }

    // 既存のマッピングデータを取得（重複チェック用）
    const existingData = mappingSheet.getDataRange().getValues();
    const existingKeys = new Set();

    for (let i = 1; i < existingData.length; i++) {
      const aiName = existingData[i][0] || '';
      const masterName = existingData[i][1] || '';
      const customer = existingData[i][2] || '';

      if (aiName && customer) {
        existingKeys.add(`${aiName}|${customer}`);
      }
      if (masterName && customer) {
        existingKeys.add(`${masterName}|${customer}`);
      }
    }

    // マッピングシートに一括追加（既存データを除く）
    const newMappings = [];
    mappingMap.forEach((value, key) => {
      // 既存データにない組み合わせのみ追加
      if (!existingKeys.has(key)) {
        newMappings.push([
          value.shippingTo,  // AI認識名（受注データの発送先名）
          value.shippingTo,  // マスタ名（受注データの発送先名）
          value.customer,    // 顧客名
          value.count,       // 信頼度（出現回数）
          new Date()         // 最終使用日
        ]);
      }
    });

    if (newMappings.length > 0) {
      // データを追加（ヘッダー行の次の行から）
      const lastRow = mappingSheet.getLastRow();
      mappingSheet.getRange(lastRow + 1, 1, newMappings.length, 5).setValues(newMappings);

      Logger.log(`初回移行完了: ${newMappings.length}件のマッピングを追加しました`);
    } else {
      Logger.log('初回移行: 新規追加するマッピングはありませんでした（すべて既存データ）');
    }

    return {
      success: true,
      recordsProcessed: recordsProcessed,
      uniqueCombinations: mappingMap.size,
      mappingsCreated: newMappings.length,
      message: `${recordsProcessed}件の受注レコードを処理し、${newMappings.length}件の新規マッピングを作成しました`
    };

  } catch (error) {
    Logger.log('初回移行エラー: ' + error.toString());
    return {
      success: false,
      error: error.toString()
    };
  }
}
