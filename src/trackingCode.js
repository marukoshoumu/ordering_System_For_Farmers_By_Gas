/**
 * 受注シートのヘッダーをセットアップする
 * 「追跡番号」列がない場合は追加し、既存のデータの整合性を保つ
 */
function setupOrderSheetHeaders() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('受注');
    if (!sheet) return;

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const trackingColIndex = headers.indexOf('追跡番号');
    const shippedColIndex = headers.indexOf('出荷済');
    const subtotalColIndex = headers.indexOf('小計');

    // 追跡番号列がない場合
    if (trackingColIndex === -1) {
        if (shippedColIndex !== -1) {
            // 「出荷済」の右隣に挿入
            sheet.insertColumnAfter(shippedColIndex + 1);
            sheet.getRange(1, shippedColIndex + 2).setValue('追跡番号');
            Logger.log('「追跡番号」列を追加しました（出荷済の隣）');
        } else if (subtotalColIndex !== -1) {
            // 「小計」の左隣に挿入
            sheet.insertColumnBefore(subtotalColIndex + 1);
            sheet.getRange(1, subtotalColIndex + 1).setValue('追跡番号');
            Logger.log('「追跡番号」列を追加しました（小計の隣）');
        } else {
            // 末尾に追加
            sheet.getRange(1, headers.length + 1).setValue('追跡番号');
            Logger.log('「追跡番号」列を末尾に追加しました');
        }
    }
}

/**
 * 追跡番号を更新し、ステータスを「出荷済」にする
 * @param {string} orderId - 受注ID
 * @param {string} trackingNumber - 追跡番号（送り状番号）
 * @returns {Object} 実行結果
 */
function updateOrderTrackingStatus(orderId, trackingNumber) {
    try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const sheet = ss.getSheetByName('受注');
        const data = sheet.getDataRange().getValues();
        const headers = data[0];

        const orderIdCol = headers.indexOf('受注ID');
        const trackingCol = headers.indexOf('追跡番号');
        const shippedCol = headers.indexOf('出荷済');

        if (orderIdCol === -1 || trackingCol === -1 || shippedCol === -1) {
            return { success: false, message: '必要な列（受注ID, 追跡番号, 出荷済）が見つかりません。' };
        }

        let updatedCount = 0;
        for (let i = 1; i < data.length; i++) {
            if (data[i][orderIdCol] === orderId) {
                // 追跡番号をセット
                sheet.getRange(i + 1, trackingCol + 1).setValue(trackingNumber);
                // 出荷済を「○」にセット
                sheet.getRange(i + 1, shippedCol + 1).setValue('○');
                updatedCount++;
            }
        }

        if (updatedCount === 0) {
            return { success: false, message: '該当する受注IDが見つかりませんでした。' };
        }

        return {
            success: true,
            message: `受注ID: ${orderId} に追跡番号 ${trackingNumber} を紐付け、出荷済に更新しました。`
        };
    } catch (error) {
        return { success: false, message: '更新エラー: ' + error.toString() };
    }
}

/**
 * 定期トリガー用: 配送状況の自動同期
 * 追跡番号があり、かつステータスが「完了」でないものを対象に巡回
 */
function syncDeliveryStatus() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('受注');
    const data = sheet.getDataRange().getValues();
    const headers = data[0];

    const orderIdCol = headers.indexOf('受注ID');
    const trackingCol = headers.indexOf('追跡番号');
    const statusCol = headers.indexOf('ステータス');
    const deliveryMethodCol = headers.indexOf('納品方法');

    if (trackingCol === -1) return;

    // メモ：ステータス列がない場合は、最後の列に追加するなどの処理が必要だが、
    // 現状は既存のステータス管理に準ずる

    // ユニークな受注ID（1受注複数行を想定）ごとに処理
    const processedOrderIds = new Set();

    for (let i = 1; i < data.length; i++) {
        const orderId = data[i][orderIdCol];
        const trackingNo = String(data[i][trackingCol]).trim();
        const deliveryMethod = data[i][deliveryMethodCol] || '';
        const currentStatus = statusCol !== -1 ? data[i][statusCol] : '';

        if (!trackingNo || trackingNo === '' || processedOrderIds.has(orderId)) continue;
        if (currentStatus === '完了' || currentStatus === '配達完了') continue;

        processedOrderIds.add(orderId);

        // 配送状況の取得
        const newStatus = fetchTrackingStatus(trackingNo, deliveryMethod);

        if (newStatus && newStatus !== currentStatus) {
            // 全行のステータスを更新
            updateOrderStatusInSheet(orderId, newStatus);
            Logger.log(`受注ID: ${orderId} のステータスを ${newStatus} に更新しました。`);
        }
    }
}

/**
 * 各運送業者の追跡サイトからステータスを取得（HTMLスクレイピング）
 * 
 * @param {string} trackingNo - 伝票番号（追跡番号）
 * @param {string} deliveryMethod - 納品方法（'ヤマト','佐川','西濃運輸'など）
 * @returns {string|null} 判定されたステータス（「配達完了」「発送済」など）または null
 * 
 * @see normalizeTrackingStatus - 取得した生文言をシステム標準に変換
 */
function fetchTrackingStatus(trackingNo, deliveryMethod) {
    // 簡易バリデーション
    if (!trackingNo || trackingNo.length < 8) return null;

    let url = '';
    let options = { muteHttpExceptions: true };

    // キャリアごとのリクエスト設定を構築
    // ヤマト運輸系（部分一致で判定）
    if (deliveryMethod && deliveryMethod.includes('ヤマト')) {
        // 2026/01時点の最新仕様: POSTメソッドでnumber01パラメータを送信
        url = 'https://toi.kuronekoyamato.co.jp/cgi-bin/tneko';
        options.method = 'post';
        options.payload = {
            'number01': trackingNo,
            'category': '0'
        };
    }
    // 佐川急便系（部分一致で判定）
    else if (deliveryMethod && deliveryMethod.includes('佐川')) {
        // GETメソッドで追跡番号をパラメータとして送信
        // 参考: https://blog.djuggernaut.com/spreadsheet-package-tracking/
        url = 'https://k2k.sagawa-exp.co.jp/p/web/okurijosearch.do?okurijoNo=' + trackingNo;
    }
    // 西濃運輸系（部分一致で判定）
    else if (deliveryMethod && deliveryMethod.includes('西濃')) {
        // vfs.seino.co.jp ではなく、標準的な track.seino.co.jp を使用
        // POSTメソッドでGNPNO1パラメータを送信
        url = 'https://track.seino.co.jp/cgi-bin/gnpquery.pgm';
        options.method = 'post';
        options.payload = {
            'GNPNO1': trackingNo
        };
    }
    // 追跡対象外（店舗受取、配達、レターパック、その他など）
    else {
        return null;
    }

    try {
        const response = UrlFetchApp.fetch(url, options);
        let html = response.getContentText('UTF-8');

        // 西濃運輸のサイトは Shift_JIS でエンコードされているため、再取得して文字化けを防ぐ
        if (url.includes('seino.co.jp')) {
            html = response.getContentText('Shift_JIS');
        }

        // 解析しやすくするため、改行や連続する空白を1つのスペースに統合
        const cleanHtml = html.replace(/\s+/g, ' ');

        let statusText = '';

        // ヤマト運輸
        if (url.includes('kuronekoyamato')) {
            // 例: <h4 class="tracking-invoice-block-state-title">伝票番号未登録</h4>
            // タグ名(h4/div)に関わらずクラス名でターゲット
            const match = cleanHtml.match(/class="tracking-invoice-block-state-title"[^>]*>([^<]+)</);
            if (match) {
                statusText = match[1].trim();
            } else {
                // HTML構造取得失敗等のフォールバック
                // リンク文字などの誤検知を防ぐため、単純なincludesは使用せず、構造が見つからない場合はnullとする
                return null;
            }
        }
        // 佐川急便
        else if (url.includes('sagawa-exp')) {
            // <span class="state">該当なし</span> や <span class="state">配達完了</span> を抽出
            const match = cleanHtml.match(/<span\s+class="state"[^>]*>([^<]+)<\/span>/);
            if (match) {
                statusText = match[1].trim();
                if (statusText === '該当なし' || statusText === '該当する送り状が見つかりませんでした') {
                    return null; // 未登録または誤り
                }
            } else {
                // フォールバック: 主要キーワードで判定
                if (html.includes('配達完了') || html.includes('配送完了')) return '配達完了';
                if (html.includes('配達中') || html.includes('配送中')) return '配達中';
                if (html.includes('輸送中') || html.includes('集荷') || html.includes('発送')) return '発送済';
                if (html.includes('不在')) return '不在/持戻';
                return null;
            }
        }
        // 西濃運輸
        else if (url.includes('seino.co.jp')) {
            // テーブル構造を想定した抽出: <th>現在の状況</th><td>配達完了</td> 形式
            const match = cleanHtml.match(/現在の状況<\/th>\s*<td[^>]*>([^<]+)<\/td>/i);
            if (match) {
                statusText = match[1].trim();
            } else {
                // フォールバック: 主要キーワードで判定（西濃固有の表現を優先）
                if (html.includes('お届け完了')) return '配達完了';
                if (html.includes('持出中')) return '配達中';
                if (html.includes('発送')) return '発送済';
                if (html.includes('不在') || html.includes('持戻')) return '不在/持戻';
                return null;
            }
        }

        return normalizeTrackingStatus(statusText);
    } catch (e) {
        Logger.log('Fetchエラー (' + trackingNo + '): ' + e.toString());
        return null;
    }
}

/**
 * 送り状のステータス文言をシステム標準のステータスに変換する
 * @param {string} statusText キャリアサイトから抽出した生文言
 * @return {string|null} 変換後のステータス（該当なしはnull）
 * 
 * 注意: 判定順序が重要！
 * - 「発送済み」を「配達完了」と誤判定しないよう、発送系を先にチェック
 * - 具体的な文言（「配達完了」「お届け済」等）を優先判定
 */
function normalizeTrackingStatus(statusText) {
    if (!statusText) return null;

    // 配達完了系（最優先でチェック）
    // ※「完了」「お届け済」「配達済」等を判定
    if (statusText.includes('完了') ||
        statusText.includes('お届け済') ||
        statusText.includes('配達済') ||
        statusText.includes('受取済')) {
        return '配達完了';
    }

    // 配達中系
    if (statusText.includes('配達中') ||
        statusText.includes('持出') ||
        statusText.includes('出発')) {
        return '配達中';
    }

    // 不在・持戻系（「保管」は除外 - 保管中は配達予定のための一時保管なので発送済として扱う）
    if (statusText.includes('不在') ||
        statusText.includes('持戻') ||
        statusText.includes('調査中') ||
        statusText.includes('入館不可')) {
        return '不在/持戻';
    }

    // 発送済・輸送中・受付・保管系
    // ※「保管中」は配達営業所での一時保管なので「発送済」として扱う
    // ※「発送済」「発送済み」は配達完了ではなく発送済として扱う
    if (statusText.includes('発送') ||
        statusText.includes('受付') ||
        statusText.includes('集荷') ||
        statusText.includes('通過') ||
        statusText.includes('輸送') ||
        statusText.includes('到着') ||
        statusText.includes('準備') ||
        statusText.includes('保管') ||
        statusText.includes('中継')) {
        return '発送済';
    }

    return null;
}

/**
 * ステータス正規化ロジックのテスト用関数
 * GASのエディタでこの関数を選択して実行し、ログを確認してください
 */
function testTrackingNormalization() {
    const testCases = [
        { input: "配達完了", expected: "配達完了" },
        { input: "お届け完了です", expected: "配達完了" },
        { input: "お届け済", expected: "配達完了" },
        { input: "配達済", expected: "配達完了" },
        { input: "受取済", expected: "配達完了" },
        { input: "荷物受付", expected: "発送済" },
        { input: "発送済み", expected: "発送済" },  // ★重要: 「発送済み」は発送済
        { input: "発送済", expected: "発送済" },    // ★重要: 「発送済」は発送済
        { input: "作業店通過", expected: "発送済" },
        { input: "配達店到着", expected: "発送済" },
        { input: "配達準備中", expected: "発送済" },
        { input: "輸送中", expected: "発送済" },
        { input: "中継センター通過", expected: "発送済" },
        { input: "配達中", expected: "配達中" },
        { input: "持出中", expected: "配達中" },
        { input: "配達担当店を出発しました", expected: "配達中" },
        { input: "ご不在（持戻）", expected: "不在/持戻" },
        { input: "営業所保管中", expected: "不在/持戻" },
        { input: "住所不明（調査中）", expected: "不在/持戻" },
        { input: "伝票番号未登録", expected: null },
        { input: "該当なし", expected: null }
    ];

    Logger.log("--- 追跡ステータス判定テスト開始 ---");
    let successCount = 0;

    testCases.forEach((t, i) => {
        const result = normalizeTrackingStatus(t.input);
        const isSuccess = (result === t.expected);
        const statusLabel = isSuccess ? "成功" : "失敗";
        if (isSuccess) successCount++;

        Logger.log(`[${statusLabel}] Case ${i + 1}: 入力="${t.input}" -> 判定="${result}" (期待値="${t.expected}")`);
    });

    Logger.log(`--- テスト終了: ${successCount}/${testCases.length} 成功 ---`);
}

/**
 * Cloud Vision API を使用して画像から追跡番号を抽出する
 * @param {string} base64Data - 画像のBase64文字列
 * @returns {Object} 抽出結果 {success: boolean, code: string, message: string}
 */
function recognizeTrackingNumber(base64Data) {
    const apiKey = getVisionApiKey();
    if (!apiKey) {
        return { success: false, message: 'Google Cloud Vision API キーが設定されていません。' };
    }

    const url = `https://vision.googleapis.com/v1/images:annotate?key=${apiKey}`;
    const requestBody = {
        requests: [
            {
                image: { content: base64Data },
                features: [{ type: 'TEXT_DETECTION' }]
            }
        ]
    };

    try {
        const response = UrlFetchApp.fetch(url, {
            method: 'POST',
            contentType: 'application/json',
            payload: JSON.stringify(requestBody),
            muteHttpExceptions: true
        });

        const result = JSON.parse(response.getContentText());
        const annotations = result.responses[0].textAnnotations;

        if (!annotations || annotations.length === 0) {
            return { success: false, message: '文字が検出されませんでした。' };
        }

        // 全体の文字列を取得
        const fullText = annotations[0].description;

        // 配送伝票番号特有のパターンを探す (10〜12桁の数字、またはハイフン区切りの数字)
        // 例: 1234-5678-9012, 123456789012
        // 配送伝票番号特有のパターンを探す
        // 優先順位: 1. ハイフンあり(12桁) 2. ハイフンあり(10桁) 3. 12桁 4. 10桁
        const patterns = [
            /\d{4}-\d{4}-\d{4}/, // ヤマト/佐川等 (12桁)
            /\d{3}-\d{3}-\d{4}/, // 西濃/郵便等 (10桁/11桁)
            /\d{4}-\d{4}-\d{2}/, // 西濃2
            /\d{12}/,            // 12桁
            /\d{11}/,            // 11桁
            /\d{10}/             // 10桁
        ];

        let candidates = [];
        for (const pattern of patterns) {
            const matches = fullText.match(new RegExp(pattern, 'g'));
            if (matches) {
                for (const m of matches) {
                    const cleaned = m.replace(/-/g, '');
                    // 日本の電話番号（0から始まる10/11桁）は、他に候補があればスキップ
                    if ((cleaned.length === 10 || cleaned.length === 11) && cleaned.startsWith('0')) {
                        continue;
                    }
                    // 最初に見つかった有力な候補を返す
                    return { success: true, code: cleaned };
                }
            }
        }

        return { success: false, message: '認識されたテキストに伝票番号が見つかりませんでした。', rawText: fullText };

    } catch (e) {
        return { success: false, message: 'API呼び出しエラー: ' + e.toString() };
    }
}

/**
 * シート内の特定受注IDに紐づく全行のステータスを一括更新
 * 
 * @param {string} orderId - 対象の受注ID
 * @param {string} newStatus - 設定する新しいステータス名
 */
function updateOrderStatusInSheet(orderId, newStatus) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('受注');
    const data = sheet.getDataRange().getValues();
    const headers = data[0];

    const orderIdCol = headers.indexOf('受注ID');
    let statusCol = headers.indexOf('ステータス');

    if (statusCol === -1) {
        // ステータス列がなければ追加
        statusCol = headers.length;
        sheet.getRange(1, statusCol + 1).setValue('ステータス');
    }

    for (let i = 1; i < data.length; i++) {
        if (data[i][orderIdCol] === orderId) {
            sheet.getRange(i + 1, statusCol + 1).setValue(newStatus);
        }
    }
}

/**
 * 定期同期トリガーのセットアップ（初回のみ実行）
 * 1時間ごとに syncDeliveryStatus を実行するトリガーを作成
 */
function setupSyncTrigger() {
    const triggerName = 'syncDeliveryStatus';

    // 既存の同名トリガーがあれば削除（重複防止）
    const triggers = ScriptApp.getProjectTriggers();
    for (const trigger of triggers) {
        if (trigger.getHandlerFunction() === triggerName) {
            ScriptApp.deleteTrigger(trigger);
        }
    }

    // 新しいトリガーを作成（1時間ごと）
    ScriptApp.newTrigger(triggerName)
        .timeBased()
        .everyHours(1)
        .create();

    Logger.log('配送状況同期トリガーを設定しました（1時間毎）');
}
