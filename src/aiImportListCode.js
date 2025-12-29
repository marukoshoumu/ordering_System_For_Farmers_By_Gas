/**
 * AI取込一覧 - バックエンド関数
 * 
 * LINE Botプロジェクトの仮受注シートからデータを取得し、
 * 受注システムで表示・処理するための関数群
 */

// ========== 設定 ==========

// LINE Botプロジェクトのスプレッドシートの情報は config.js から取得
// getLineBotSpreadsheetId() を使用
  
// ========== AI取込一覧 ==========

/**
 * AI取込一覧データを取得
 * @returns {string} JSON文字列
 */
function getAIImportList() {
  try {
    const ss = SpreadsheetApp.openById(getLineBotSpreadsheetId());
    const sheet = ss.getSheetByName('仮受注');
    
    if (!sheet) {
      return JSON.stringify([]);
    }
    
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      return JSON.stringify([]);
    }
    
    const headers = data[0];
    const result = [];
    
    // ヘッダーのインデックスを取得
    const idxTempOrderId = headers.indexOf('仮受注ID');
    const idxCreatedAt = headers.indexOf('登録日時');
    const idxStatus = headers.indexOf('ステータス');
    const idxCustomerName = headers.indexOf('顧客名');
    const idxShippingTo = headers.indexOf('発送先');
    const idxItemsData = headers.indexOf('商品データ');
    const idxOriginalText = headers.indexOf('原文');
    const idxAnalysisResult = headers.indexOf('解析結果');
    
    // データ行を処理（新しい順にソート）
    for (let i = data.length - 1; i >= 1; i--) {
      const row = data[i];
      
      // 商品データをパース
      let items = [];
      try {
        items = JSON.parse(row[idxItemsData] || '[]');
      } catch (e) {
        items = [];
      }
      
      // 解析結果から信頼度を取得
      let confidence = '中';
      try {
        const analysis = JSON.parse(row[idxAnalysisResult] || '{}');
        confidence = analysis.overallConfidence || '中';
        // high/medium/low を 高/中/低 に変換
        if (confidence === 'high') confidence = '高';
        else if (confidence === 'medium') confidence = '中';
        else if (confidence === 'low') confidence = '低';
      } catch (e) {
        confidence = '中';
      }
      
      result.push({
        tempOrderId: row[idxTempOrderId] || '',
        createdAt: row[idxCreatedAt] ? new Date(row[idxCreatedAt]).toISOString() : '',
        status: row[idxStatus] || '未処理',
        customerName: row[idxCustomerName] || '',
        shippingTo: row[idxShippingTo] || '',
        items: items.map(item => ({
          productName: item.productName || '',
          originalText: item.originalText || '',
          quantity: item.quantity || '',
          unit: item.unit || 'ケース'
        })),
        originalText: row[idxOriginalText] || '',
        confidence: confidence
      });
    }
    
    return JSON.stringify(result);
    
  } catch (error) {
    console.error('getAIImportList エラー:', error);
    return JSON.stringify([]);
  }
}

/**
 * 仮受注を却下
 * @param {string} tempOrderId - 仮受注ID
 * @returns {Object} 結果
 */
function rejectTempOrder(tempOrderId) {
  try {
    const ss = SpreadsheetApp.openById(getLineBotSpreadsheetId());
    const sheet = ss.getSheetByName('仮受注');
    
    if (!sheet) {
      return { success: false, error: '仮受注シートがありません' };
    }
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const idxTempOrderId = headers.indexOf('仮受注ID');
    const idxStatus = headers.indexOf('ステータス');
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][idxTempOrderId] === tempOrderId) {
        // ステータスを「却下」に更新
        sheet.getRange(i + 1, idxStatus + 1).setValue('却下');
        return { success: true };
      }
    }
    
    return { success: false, error: '該当する仮受注が見つかりません' };
    
  } catch (error) {
    console.error('rejectTempOrder エラー:', error);
    return { success: false, error: error.message };
  }
}

/**
 * 仮受注データを取得（受注画面でのプリセット用）
 * @param {string} tempOrderId - 仮受注ID
 * @returns {Object|null} 仮受注データ
 */
function getTempOrderById(tempOrderId) {
  try {
    const ss = SpreadsheetApp.openById(getLineBotSpreadsheetId());
    const sheet = ss.getSheetByName('仮受注');
    
    if (!sheet) {
      return null;
    }
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    const idxTempOrderId = headers.indexOf('仮受注ID');
    const idxOriginalText = headers.indexOf('原文');
    const idxAnalysisResult = headers.indexOf('解析結果');
    const idxStatus = headers.indexOf('ステータス');
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][idxTempOrderId] === tempOrderId) {
        let analysisResult = null;
        try {
          analysisResult = JSON.parse(data[i][idxAnalysisResult] || '{}');
        } catch (e) {
          analysisResult = {};
        }
        
        return {
          tempOrderId: tempOrderId,
          originalText: data[i][idxOriginalText] || '',
          analysisResult: analysisResult,
          status: data[i][idxStatus] || '未処理'
        };
      }
    }
    
    return null;
    
  } catch (error) {
    console.error('getTempOrderById エラー:', error);
    return null;
  }
}

/**
 * 仮受注のステータスを更新
 * @param {string} tempOrderId - 仮受注ID
 * @param {string} status - 新しいステータス
 * @returns {Object} 結果
 */
function updateTempOrderStatus(tempOrderId, status) {
  try {
    const ss = SpreadsheetApp.openById(getLineBotSpreadsheetId());
    const sheet = ss.getSheetByName('仮受注');

    if (!sheet) {
      return { success: false, error: '仮受注シートがありません' };
    }

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const idxTempOrderId = headers.indexOf('仮受注ID');
    const idxStatus = headers.indexOf('ステータス');

    for (let i = 1; i < data.length; i++) {
      if (data[i][idxTempOrderId] === tempOrderId) {
        sheet.getRange(i + 1, idxStatus + 1).setValue(status);
        return { success: true };
      }
    }

    return { success: false, error: '該当する仮受注が見つかりません' };

  } catch (error) {
    console.error('updateTempOrderStatus エラー:', error);
    return { success: false, error: error.message };
  }
}

/**
 * 仮受注データを完全に削除
 * @param {string} tempOrderId - 仮受注ID
 * @returns {Object} 結果
 */
function deleteTempOrderData(tempOrderId) {
  try {
    const ss = SpreadsheetApp.openById(getLineBotSpreadsheetId());
    const sheet = ss.getSheetByName('仮受注');

    if (!sheet) {
      return { success: false, error: '仮受注シートがありません' };
    }

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const idxTempOrderId = headers.indexOf('仮受注ID');

    for (let i = 1; i < data.length; i++) {
      if (data[i][idxTempOrderId] === tempOrderId) {
        sheet.deleteRow(i + 1);
        return { success: true };
      }
    }

    return { success: false, error: '該当する仮受注が見つかりません' };

  } catch (error) {
    console.error('deleteTempOrderData エラー:', error);
    return { success: false, error: error.message };
  }
}

/**
 * 却下した仮受注を再登録（ステータスを「未処理」に戻す）
 * @param {string} tempOrderId - 仮受注ID
 * @returns {Object} 結果
 */
function reactivateTempOrder(tempOrderId) {
  try {
    const ss = SpreadsheetApp.openById(getLineBotSpreadsheetId());
    const sheet = ss.getSheetByName('仮受注');

    if (!sheet) {
      return { success: false, error: '仮受注シートがありません' };
    }

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const idxTempOrderId = headers.indexOf('仮受注ID');
    const idxStatus = headers.indexOf('ステータス');

    for (let i = 1; i < data.length; i++) {
      if (data[i][idxTempOrderId] === tempOrderId) {
        const currentStatus = data[i][idxStatus];

        // 却下状態のみ再登録可能
        if (currentStatus !== '却下') {
          return { success: false, error: '却下状態のデータのみ再登録できます' };
        }

        sheet.getRange(i + 1, idxStatus + 1).setValue('未処理');
        return { success: true };
      }
    }

    return { success: false, error: '該当する仮受注が見つかりません' };

  } catch (error) {
    console.error('reactivateTempOrder エラー:', error);
    return { success: false, error: error.message };
  }
}

/**
 * 未処理の仮受注件数を取得（バッジ表示用）
 * @returns {number} 未処理件数
 */
function getPendingImportCount() {
  try {
    const ss = SpreadsheetApp.openById(getLineBotSpreadsheetId());
    const sheet = ss.getSheetByName('仮受注');
    
    if (!sheet) {
      return 0;
    }
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const idxStatus = headers.indexOf('ステータス');
    
    let count = 0;
    for (let i = 1; i < data.length; i++) {
      const status = data[i][idxStatus];
      // 「未処理」「未確定」または空白を未処理としてカウント（後方互換性のため）
      if (status === '未処理' || status === '未確定' || !status) {
        count++;
      }
    }

    return count;
    
  } catch (error) {
    console.error('getPendingImportCount エラー:', error);
    return 0;
  }
}
