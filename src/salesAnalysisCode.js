/**
 * 売上分析システム
 * 月次集計・顧客リテンション分析・季節パターン学習・顧客傾向分析
 */

/**
 * 受注シートから月次集計データを生成
 * トリガー: 受注シート更新時に自動実行
 */
function aggregateMonthlySales() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const orders = getAllRecords('受注');

    if (!orders || orders.length === 0) {
      return { success: false, message: '受注データがありません' };
    }

    // 月次×商品でグルーピング
    const monthlyData = {};

    orders.forEach(order => {
      try {
        const orderDate = new Date(order['受注日']);
        if (isNaN(orderDate.getTime())) return; // 無効な日付はスキップ

        const yearMonth = Utilities.formatDate(orderDate, 'JST', 'yyyy-MM');
        const productName = order['商品名'] || '';
        const quantity = Number(order['受注数']) || 0;
        const price = Number(order['販売価格']) || 0;
        const category = order['商品分類'] || '';

        if (!productName) return; // 商品名なしはスキップ

        // 送料・代引手数料を除外
        if (productName === '送料' || productName === '代引手数料') {
          return;
        }

        const key = `${yearMonth}|${productName}`;

        if (!monthlyData[key]) {
          monthlyData[key] = {
            yearMonth: yearMonth,
            productName: productName,
            category: category,
            totalQuantity: 0,
            totalSales: 0,
            orderCount: 0
          };
        }

        monthlyData[key].totalQuantity += quantity;
        monthlyData[key].totalSales += (quantity * price);
        monthlyData[key].orderCount += 1;
      } catch (e) {
        // エラーをスキップ
      }
    });

    // シート「月次集計」取得または作成
    let aggregationSheet = ss.getSheetByName('月次集計');
    if (!aggregationSheet) {
      aggregationSheet = ss.insertSheet('月次集計');
    }

    // ヘッダーとデータ行を作成
    const headers = ['年月', '商品名', '商品分類', '受注数合計',
                     '売上高', '受注回数', '最終更新'];
    const rows = [headers];

    Object.values(monthlyData).forEach(data => {
      rows.push([
        data.yearMonth,
        data.productName,
        data.category,
        data.totalQuantity,
        data.totalSales,
        data.orderCount,
        new Date()
      ]);
    });

    // シートクリア＆書き込み
    aggregationSheet.clear();
    if (rows.length > 1) {
      // 年月列を先にテキスト形式に設定（日付型への自動変換を防ぐ）
      aggregationSheet.getRange(1, 1, rows.length, 1).setNumberFormat('@');

      aggregationSheet.getRange(1, 1, rows.length, headers.length).setValues(rows);

      // ヘッダー装飾
      aggregationSheet.getRange(1, 1, 1, headers.length)
        .setBackground('#4a90d9')
        .setFontColor('#ffffff')
        .setFontWeight('bold');

      // 数値列のフォーマット
      if (rows.length > 1) {
        aggregationSheet.getRange(2, 4, rows.length - 1, 1).setNumberFormat('#,##0'); // 受注数
        aggregationSheet.getRange(2, 5, rows.length - 1, 1).setNumberFormat('¥#,##0'); // 売上高
        aggregationSheet.getRange(2, 6, rows.length - 1, 1).setNumberFormat('#,##0'); // 受注回数
      }
    }

    return {
      success: true,
      message: `月次集計完了: ${Object.keys(monthlyData).length}件`,
      count: Object.keys(monthlyData).length
    };
  } catch (error) {
    return { success: false, message: 'エラー: ' + error.toString() };
  }
}

/**
 * 顧客リテンション分析
 * 最終注文日からの経過日数で顧客をセグメント化
 */
function analyzeCustomerRetention() {
  try {
    const orders = getAllRecords('受注');
    if (!orders || orders.length === 0) {
      return { success: false, message: '受注データがありません' };
    }

    // 顧客ごとの最終注文日を集計
    const customerLastOrder = {};
    const customerTotalSales = {};
    const customerOrderCount = {};

    orders.forEach(order => {
      const customerName = order['顧客名'] || '';
      const orderDate = new Date(order['受注日']);
      const productName = order['商品名'] || '';
      const quantity = Number(order['受注数']) || 0;
      const price = Number(order['販売価格']) || 0;
      const sales = quantity * price;

      if (!customerName || isNaN(orderDate.getTime())) return;

      // 送料・代引手数料を除外
      const isShippingFee = productName === '送料' || productName === '代引手数料';

      // 最終注文日を更新（送料・代引手数料も含めて判定）
      if (!customerLastOrder[customerName] || orderDate > customerLastOrder[customerName]) {
        customerLastOrder[customerName] = orderDate;
      }

      // 売上累計（送料・代引手数料は除外）
      if (!isShippingFee) {
        customerTotalSales[customerName] = (customerTotalSales[customerName] || 0) + sales;
      }

      // 注文回数（送料・代引手数料も含めてカウント）
      customerOrderCount[customerName] = (customerOrderCount[customerName] || 0) + 1;
    });

    // 現在日時
    const now = new Date();

    // 顧客セグメント分類
    const segments = {
      active: [],      // 30日以内に注文
      warning: [],     // 31-60日
      risk: [],        // 61-90日
      churned: []      // 91日以上注文なし
    };

    Object.keys(customerLastOrder).forEach(customerName => {
      const lastOrderDate = customerLastOrder[customerName];
      const daysSinceLastOrder = Math.floor((now - lastOrderDate) / (1000 * 60 * 60 * 24));

      const customerData = {
        customerName: customerName,
        lastOrderDate: Utilities.formatDate(lastOrderDate, 'JST', 'yyyy-MM-dd'),
        daysSinceLastOrder: daysSinceLastOrder,
        totalSales: customerTotalSales[customerName],
        orderCount: customerOrderCount[customerName]
      };

      if (daysSinceLastOrder <= 30) {
        segments.active.push(customerData);
      } else if (daysSinceLastOrder <= 60) {
        segments.warning.push(customerData);
      } else if (daysSinceLastOrder <= 90) {
        segments.risk.push(customerData);
      } else {
        segments.churned.push(customerData);
      }
    });

    // 各セグメントを売上順にソート
    ['active', 'warning', 'risk', 'churned'].forEach(segment => {
      segments[segment].sort((a, b) => b.totalSales - a.totalSales);
    });

    // シート「顧客リテンション分析」に保存
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let retentionSheet = ss.getSheetByName('顧客リテンション分析');
    if (!retentionSheet) {
      retentionSheet = ss.insertSheet('顧客リテンション分析');
    }

    retentionSheet.clear();

    const headers = ['セグメント', '顧客名', '最終注文日', '経過日数', '累計売上', '注文回数'];
    const rows = [headers];

    const segmentLabels = {
      active: 'アクティブ(30日以内)',
      warning: '要注意(31-60日)',
      risk: 'リスク(61-90日)',
      churned: '離脱(91日以上)'
    };

    ['active', 'warning', 'risk', 'churned'].forEach(segment => {
      segments[segment].forEach(customer => {
        rows.push([
          segmentLabels[segment],
          customer.customerName,
          customer.lastOrderDate,
          customer.daysSinceLastOrder,
          customer.totalSales,
          customer.orderCount
        ]);
      });
    });

    if (rows.length > 1) {
      retentionSheet.getRange(1, 1, rows.length, headers.length).setValues(rows);

      // ヘッダー装飾
      retentionSheet.getRange(1, 1, 1, headers.length)
        .setBackground('#4a90d9')
        .setFontColor('#ffffff')
        .setFontWeight('bold');

      // 数値フォーマット
      if (rows.length > 1) {
        retentionSheet.getRange(2, 4, rows.length - 1, 1).setNumberFormat('#,##0'); // 経過日数
        retentionSheet.getRange(2, 5, rows.length - 1, 1).setNumberFormat('¥#,##0'); // 累計売上
        retentionSheet.getRange(2, 6, rows.length - 1, 1).setNumberFormat('#,##0'); // 注文回数
      }
    }


    return {
      success: true,
      segments: {
        active: segments.active.length,
        warning: segments.warning.length,
        risk: segments.risk.length,
        churned: segments.churned.length
      },
      topChurnRisk: segments.churned.slice(0, 10) // 離脱リスク上位10顧客
    };
  } catch (error) {
    return { success: false, message: 'エラー: ' + error.toString() };
  }
}

/**
 * 季節パターン学習
 * 過去データから商品ごとの季節トレンドを分析
 */
function analyzeSeasonalPatterns() {
  try {
    const orders = getAllRecords('受注');
    if (!orders || orders.length === 0) {
      return { success: false, message: '受注データがありません' };
    }

    // 商品×月のマトリクス（過去2年分）
    const productMonthData = {};
    const now = new Date();
    const twoYearsAgo = new Date(now.getFullYear() - 2, now.getMonth(), 1);

    orders.forEach(order => {
      const orderDate = new Date(order['受注日']);
      const productName = order['商品名'] || '';
      const quantity = Number(order['受注数']) || 0;
      const price = Number(order['販売価格']) || 0;
      const sales = quantity * price;

      if (!productName || isNaN(orderDate.getTime()) || orderDate < twoYearsAgo) return;

      // 送料・代引手数料を除外
      if (productName === '送料' || productName === '代引手数料') return;

      const month = orderDate.getMonth() + 1; // 1-12
      const year = orderDate.getFullYear();

      if (!productMonthData[productName]) {
        productMonthData[productName] = {};
      }

      if (!productMonthData[productName][month]) {
        productMonthData[productName][month] = {
          totalSales: 0,
          totalQuantity: 0,
          years: {}
        };
      }

      productMonthData[productName][month].totalSales += sales;
      productMonthData[productName][month].totalQuantity += quantity;

      if (!productMonthData[productName][month].years[year]) {
        productMonthData[productName][month].years[year] = {
          sales: 0,
          quantity: 0
        };
      }

      productMonthData[productName][month].years[year].sales += sales;
      productMonthData[productName][month].years[year].quantity += quantity;
    });

    // 各商品の季節パターンを分析
    const seasonalPatterns = [];

    Object.keys(productMonthData).forEach(productName => {
      const monthlyData = productMonthData[productName];

      // 月別平均売上を計算
      const monthlyAvg = {};
      let totalYearlySales = 0;

      for (let month = 1; month <= 12; month++) {
        if (monthlyData[month]) {
          const years = Object.keys(monthlyData[month].years).length;
          monthlyAvg[month] = monthlyData[month].totalSales / years;
          totalYearlySales += monthlyAvg[month];
        } else {
          monthlyAvg[month] = 0;
        }
      }

      if (totalYearlySales === 0) return;

      // ピーク月とオフシーズン月を特定
      let peakMonth = 1;
      let peakSales = 0;
      let offMonth = 1;
      let offSales = Infinity;

      for (let month = 1; month <= 12; month++) {
        if (monthlyAvg[month] > peakSales) {
          peakSales = monthlyAvg[month];
          peakMonth = month;
        }
        if (monthlyAvg[month] < offSales && monthlyAvg[month] > 0) {
          offSales = monthlyAvg[month];
          offMonth = month;
        }
      }

      // 季節性の強さ（変動係数）
      const avgSales = totalYearlySales / 12;
      const variance = Object.values(monthlyAvg).reduce((sum, val) => {
        return sum + Math.pow(val - avgSales, 2);
      }, 0) / 12;
      const stdDev = Math.sqrt(variance);
      const seasonalityStrength = avgSales > 0 ? (stdDev / avgSales) : 0;

      seasonalPatterns.push({
        productName: productName,
        peakMonth: peakMonth,
        peakSales: peakSales,
        offMonth: offMonth,
        offSales: offSales,
        seasonalityStrength: seasonalityStrength,
        avgMonthlySales: avgSales
      });
    });

    // 季節性の強い順にソート
    seasonalPatterns.sort((a, b) => b.seasonalityStrength - a.seasonalityStrength);

    // シート「季節パターン分析」に保存
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let seasonalSheet = ss.getSheetByName('季節パターン分析');
    if (!seasonalSheet) {
      seasonalSheet = ss.insertSheet('季節パターン分析');
    }

    seasonalSheet.clear();

    const headers = ['商品名', 'ピーク月', 'ピーク売上', 'オフ月', 'オフ売上',
                     '季節性強度', '月平均売上', '推奨アクション'];
    const rows = [headers];

    seasonalPatterns.forEach(pattern => {
      const monthNames = ['', '1月', '2月', '3月', '4月', '5月', '6月',
                          '7月', '8月', '9月', '10月', '11月', '12月'];

      let recommendation = '';
      if (pattern.seasonalityStrength > 0.5) {
        recommendation = `${monthNames[pattern.peakMonth]}に重点的に仕入れ・販促`;
      } else if (pattern.seasonalityStrength > 0.3) {
        recommendation = `やや季節性あり（${monthNames[pattern.peakMonth]}が売れ筋）`;
      } else {
        recommendation = '通年安定需要';
      }

      rows.push([
        pattern.productName,
        monthNames[pattern.peakMonth],
        pattern.peakSales,
        monthNames[pattern.offMonth],
        pattern.offSales,
        pattern.seasonalityStrength,
        pattern.avgMonthlySales,
        recommendation
      ]);
    });

    if (rows.length > 1) {
      seasonalSheet.getRange(1, 1, rows.length, headers.length).setValues(rows);

      // ヘッダー装飾
      seasonalSheet.getRange(1, 1, 1, headers.length)
        .setBackground('#4a90d9')
        .setFontColor('#ffffff')
        .setFontWeight('bold');

      // 数値フォーマット
      if (rows.length > 1) {
        seasonalSheet.getRange(2, 3, rows.length - 1, 1).setNumberFormat('¥#,##0'); // ピーク売上
        seasonalSheet.getRange(2, 5, rows.length - 1, 1).setNumberFormat('¥#,##0'); // オフ売上
        seasonalSheet.getRange(2, 6, rows.length - 1, 1).setNumberFormat('0.00'); // 季節性強度
        seasonalSheet.getRange(2, 7, rows.length - 1, 1).setNumberFormat('¥#,##0'); // 月平均売上
      }
    }


    return {
      success: true,
      count: seasonalPatterns.length,
      topSeasonal: seasonalPatterns.slice(0, 10) // 季節性の強い上位10商品
    };
  } catch (error) {
    return { success: false, message: 'エラー: ' + error.toString() };
  }
}

/**
 * 顧客傾向データパターン学習
 * 顧客ごとの購入パターンを分析（RFM分析 + 商品親和性）
 */
function analyzeCustomerTrends() {
  try {
    const orders = getAllRecords('受注');
    if (!orders || orders.length === 0) {
      return { success: false, message: '受注データがありません' };
    }

    // 顧客ごとのデータを集計
    const customerData = {};
    const now = new Date();

    orders.forEach(order => {
      const customerName = order['顧客名'] || '';
      const orderDate = new Date(order['受注日']);
      const productName = order['商品名'] || '';
      const quantity = Number(order['受注数']) || 0;
      const price = Number(order['販売価格']) || 0;
      const sales = quantity * price;

      if (!customerName || isNaN(orderDate.getTime())) return;

      // 送料・代引手数料を除外（お気に入り商品リストから除外するため）
      if (productName === '送料' || productName === '代引手数料') return;

      if (!customerData[customerName]) {
        customerData[customerName] = {
          lastOrderDate: orderDate,
          firstOrderDate: orderDate,
          totalSales: 0,
          orderCount: 0,
          products: {},
          orderDates: []
        };
      }

      const customer = customerData[customerName];

      // Recency: 最終注文日
      if (orderDate > customer.lastOrderDate) {
        customer.lastOrderDate = orderDate;
      }

      // 初回注文日
      if (orderDate < customer.firstOrderDate) {
        customer.firstOrderDate = orderDate;
      }

      // Monetary: 累計売上
      customer.totalSales += sales;

      // Frequency: 注文回数
      customer.orderCount += 1;

      // 商品購入履歴
      if (!customer.products[productName]) {
        customer.products[productName] = {
          count: 0,
          totalQuantity: 0,
          totalSales: 0
        };
      }
      customer.products[productName].count += 1;
      customer.products[productName].totalQuantity += quantity;
      customer.products[productName].totalSales += sales;

      // 注文日リスト
      customer.orderDates.push(orderDate);
    });

    // RFMスコア計算
    const customerTrends = [];

    Object.keys(customerData).forEach(customerName => {
      const customer = customerData[customerName];

      // Recency（日数）
      const recencyDays = Math.floor((now - customer.lastOrderDate) / (1000 * 60 * 60 * 24));

      // Frequency（注文回数）
      const frequency = customer.orderCount;

      // Monetary（累計売上）
      const monetary = customer.totalSales;

      // 顧客ライフスパン（日数）
      const lifespan = Math.floor((customer.lastOrderDate - customer.firstOrderDate) / (1000 * 60 * 60 * 24));

      // 平均注文間隔
      const avgOrderInterval = lifespan > 0 && frequency > 1
        ? lifespan / (frequency - 1)
        : 0;

      // 次回注文予測日
      const nextOrderPrediction = avgOrderInterval > 0
        ? new Date(customer.lastOrderDate.getTime() + avgOrderInterval * 24 * 60 * 60 * 1000)
        : null;

      // お気に入り商品（購入回数TOP3）
      const favoriteProducts = Object.entries(customer.products)
        .sort((a, b) => b[1].count - a[1].count)
        .slice(0, 3)
        .map(([name, data]) => ({ name: name, count: data.count, sales: data.totalSales }));

      // 顧客セグメント分類
      let segment = '';
      if (recencyDays <= 30 && monetary >= 100000 && frequency >= 10) {
        segment = 'VIP';
      } else if (recencyDays <= 30 && frequency >= 5) {
        segment = 'ロイヤル';
      } else if (recencyDays <= 60 && monetary >= 50000) {
        segment = '優良';
      } else if (recencyDays > 90) {
        segment = '休眠';
      } else {
        segment = '一般';
      }

      customerTrends.push({
        customerName: customerName,
        recencyDays: recencyDays,
        frequency: frequency,
        monetary: monetary,
        avgOrderInterval: avgOrderInterval,
        nextOrderPrediction: nextOrderPrediction,
        favoriteProducts: favoriteProducts,
        segment: segment,
        lifespan: lifespan
      });
    });

    // Monetaryの高い順にソート
    customerTrends.sort((a, b) => b.monetary - a.monetary);

    // シート「顧客傾向分析」に保存
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let trendSheet = ss.getSheetByName('顧客傾向分析');
    if (!trendSheet) {
      trendSheet = ss.insertSheet('顧客傾向分析');
    }

    trendSheet.clear();

    const headers = ['顧客名', 'セグメント', '最終注文からの日数', '注文回数', '累計売上',
                     '平均注文間隔(日)', '次回注文予測日', 'お気に入り商品TOP3'];
    const rows = [headers];

    customerTrends.forEach(trend => {
      const favoriteProductsStr = trend.favoriteProducts
        .map(p => `${p.name}(${p.count}回)`)
        .join(', ');

      const nextOrderStr = trend.nextOrderPrediction
        ? Utilities.formatDate(trend.nextOrderPrediction, 'JST', 'yyyy-MM-dd')
        : '-';

      rows.push([
        trend.customerName,
        trend.segment,
        trend.recencyDays,
        trend.frequency,
        trend.monetary,
        Math.round(trend.avgOrderInterval),
        nextOrderStr,
        favoriteProductsStr
      ]);
    });

    if (rows.length > 1) {
      trendSheet.getRange(1, 1, rows.length, headers.length).setValues(rows);

      // ヘッダー装飾
      trendSheet.getRange(1, 1, 1, headers.length)
        .setBackground('#4a90d9')
        .setFontColor('#ffffff')
        .setFontWeight('bold');

      // 数値フォーマット
      if (rows.length > 1) {
        trendSheet.getRange(2, 3, rows.length - 1, 1).setNumberFormat('#,##0'); // Recency
        trendSheet.getRange(2, 4, rows.length - 1, 1).setNumberFormat('#,##0'); // Frequency
        trendSheet.getRange(2, 5, rows.length - 1, 1).setNumberFormat('¥#,##0'); // Monetary
        trendSheet.getRange(2, 6, rows.length - 1, 1).setNumberFormat('#,##0'); // 平均間隔
      }
    }


    // セグメント別集計
    const segmentSummary = {};
    customerTrends.forEach(trend => {
      if (!segmentSummary[trend.segment]) {
        segmentSummary[trend.segment] = { count: 0, totalSales: 0 };
      }
      segmentSummary[trend.segment].count += 1;
      segmentSummary[trend.segment].totalSales += trend.monetary;
    });

    return {
      success: true,
      count: customerTrends.length,
      segmentSummary: segmentSummary,
      vipCustomers: customerTrends.filter(t => t.segment === 'VIP')
    };
  } catch (error) {
    return { success: false, message: 'エラー: ' + error.toString() };
  }
}

/**
 * 全分析を一括実行
 */
function runAllAnalyses() {

  const results = {
    monthly: aggregateMonthlySales(),
    retention: analyzeCustomerRetention(),
    seasonal: analyzeSeasonalPatterns(),
    trends: analyzeCustomerTrends()
  };


  return results;
}

/**
 * 商品略称マッピングを取得
 * @return {Object} 商品名をキー、略称を値とするオブジェクト
 */
function getProductAbbreviationMap() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const productSheet = ss.getSheetByName('商品');

    if (!productSheet) {
      return {};
    }

    const values = productSheet.getDataRange().getValues();
    const headers = values[0];

    const nameIndex = headers.indexOf('商品名');
    const abbrevIndex = headers.indexOf('略称');

    if (nameIndex === -1) {
      return {};
    }

    const abbreviationMap = {};

    for (let i = 1; i < values.length; i++) {
      const productName = values[i][nameIndex];
      const abbreviation = abbrevIndex !== -1 ? values[i][abbrevIndex] : '';

      if (productName) {
        // 略称がなければ商品名をそのまま使用
        abbreviationMap[productName] = abbreviation || productName;
      }
    }

    return abbreviationMap;

  } catch (error) {
    return {};
  }
}

/**
 * ダッシュボード用データ取得
 * @param {string} targetMonth - 対象月（yyyy-MM形式）
 */
function getSalesDashboardData(targetMonth) {
  try {
    if (!targetMonth) {
      targetMonth = Utilities.formatDate(new Date(), 'JST', 'yyyy-MM');
    }

    // 商品略称マッピング取得
    const abbreviationMap = getProductAbbreviationMap();

    // 月次集計データ取得
    let aggregationData = getAllRecords('月次集計');

    if (!aggregationData || aggregationData.length === 0) {
      // データがなければ集計実行
      aggregateMonthlySales();
      Utilities.sleep(2000);
      aggregationData = getAllRecords('月次集計');

      if (!aggregationData || aggregationData.length === 0) {
        return {
          success: false,
          message: '月次集計データが取得できませんでした。runAllAnalyses()を実行してください。'
        };
      }
    }

    // 対象月のデータフィルタ（数値型に変換）+ 送料・代引手数料を除外
    const monthData = aggregationData
      .filter(d => {
        // 年月が Date オブジェクトの場合は文字列に変換して比較
        let yearMonth = d['年月'];
        if (yearMonth instanceof Date) {
          yearMonth = Utilities.formatDate(yearMonth, 'JST', 'yyyy-MM');
        }
        return yearMonth === targetMonth;
      })
      .filter(d => {
        // 送料・代引手数料を除外
        const productName = d['商品名'] || '';
        return productName !== '送料' && productName !== '代引手数料';
      })
      .map(d => ({
        ...d,
        '売上高': Number(d['売上高']) || 0,
        '受注数合計': Number(d['受注数合計']) || 0,
        '受注回数': Number(d['受注回数']) || 0
      }));

    // 売上順にソート
    monthData.sort((a, b) => b['売上高'] - a['売上高']);

    // TOP20取得
    const top20 = monthData.slice(0, 20);

    // サマリー計算
    const totalSales = monthData.reduce((sum, d) => sum + d['売上高'], 0);
    const totalQuantity = monthData.reduce((sum, d) => sum + d['受注数合計'], 0);
    const totalOrders = monthData.reduce((sum, d) => sum + d['受注回数'], 0);
    const productCount = monthData.length;

    // 前月データ取得（数値型に変換）
    const prevMonth = getPreviousMonth(targetMonth);
    const prevMonthData = aggregationData
      .filter(d => {
        // 年月が Date オブジェクトの場合は文字列に変換して比較
        let yearMonth = d['年月'];
        if (yearMonth instanceof Date) {
          yearMonth = Utilities.formatDate(yearMonth, 'JST', 'yyyy-MM');
        }
        return yearMonth === prevMonth;
      })
      .map(d => ({
        ...d,
        '売上高': Number(d['売上高']) || 0,
        '受注回数': Number(d['受注回数']) || 0
      }));
    const prevTotalSales = prevMonthData.reduce((sum, d) => sum + d['売上高'], 0);
    const prevTotalOrders = prevMonthData.reduce((sum, d) => sum + d['受注回数'], 0);

    // 前月比計算
    const salesGrowth = prevTotalSales > 0
      ? ((totalSales - prevTotalSales) / prevTotalSales * 100)
      : 0;
    const ordersGrowth = prevTotalOrders > 0
      ? (totalOrders - prevTotalOrders)
      : 0;

    // 商品分類別集計
    const categoryData = {};
    monthData.forEach(d => {
      const category = d['商品分類'] || '未分類';
      if (!categoryData[category]) {
        categoryData[category] = { sales: 0, quantity: 0 };
      }
      categoryData[category].sales += (d['売上高'] || 0);
      categoryData[category].quantity += (d['受注数合計'] || 0);
    });

    // 顧客別売上TOP10集計（受注シートから直接集計）
    const orderRecords = getAllRecords('受注');
    const customerSales = {};

    orderRecords.forEach(order => {
      // 受注日が対象月に該当するかチェック
      const orderDate = order['受注日'];
      if (!orderDate) return;

      let orderYearMonth = '';
      if (orderDate instanceof Date) {
        orderYearMonth = Utilities.formatDate(orderDate, 'JST', 'yyyy-MM');
      } else if (typeof orderDate === 'string') {
        // 文字列の場合、yyyy/MM/dd または yyyy-MM-dd 形式をパース
        const dateStr = orderDate.toString();

        // yyyy/MM/dd形式の場合
        if (dateStr.includes('/')) {
          const parts = dateStr.split('/');
          if (parts.length >= 2) {
            orderYearMonth = parts[0] + '-' + parts[1].padStart(2, '0');
          }
        }
        // yyyy-MM-dd形式の場合
        else if (dateStr.includes('-') && dateStr.length >= 7) {
          orderYearMonth = dateStr.substring(0, 7);
        }
      }

      if (orderYearMonth !== targetMonth) return;

      // 顧客名と商品名を取得（前後の空白を削除）
      const customerName = (order['顧客名'] || '不明').trim();
      const productName = (order['商品名'] || '').trim();

      // 受注数と販売価格から売上計算
      const quantity = Number(order['受注数']) || 0;
      const price = Number(order['販売価格']) || 0;
      const sales = quantity * price;

      // 送料・代引手数料除外
      if (productName === '送料' || productName === '代引手数料') {
        return;
      }

      if (!customerSales[customerName]) {
        customerSales[customerName] = {
          '顧客名': customerName,
          '売上高': 0,
          '受注回数': 0,
          '購入数量': 0
        };
      }

      customerSales[customerName]['売上高'] += sales;
      customerSales[customerName]['受注回数'] += 1;
      customerSales[customerName]['購入数量'] += quantity;
    });

    // 顧客別売上を配列化してソート
    const customerTop10 = Object.values(customerSales)
      .sort((a, b) => b['売上高'] - a['売上高'])
      .slice(0, 10);

    // 顧客リテンション分析データ取得
    const retentionData = getAllRecords('顧客リテンション分析');

    const churnRiskCustomers = retentionData
      ? retentionData.filter(d => d['セグメント'] && d['セグメント'].includes('離脱'))
        .slice(0, 10)
        .map(d => {
          // Date オブジェクトを文字列に変換（クライアント転送のため）
          const customer = { ...d };
          if (customer['最終注文日'] instanceof Date) {
            customer['最終注文日'] = Utilities.formatDate(customer['最終注文日'], 'JST', 'yyyy-MM-dd');
          }
          return customer;
        })
      : [];

    // TOP20のデータからもDateオブジェクトを除外し、略称を追加
    const cleanTop20 = top20.map(item => {
      const cleanItem = {};
      for (const key in item) {
        if (item[key] instanceof Date) {
          cleanItem[key] = Utilities.formatDate(item[key], 'JST', 'yyyy-MM-dd');
        } else {
          cleanItem[key] = item[key];
        }
      }
      // 商品略称を追加
      const productName = item['商品名'];
      cleanItem['商品略称'] = abbreviationMap[productName] || productName;
      return cleanItem;
    });

    return {
      success: true,
      targetMonth: targetMonth,
      summary: {
        totalSales: totalSales,
        totalQuantity: totalQuantity,
        totalOrders: totalOrders,
        productCount: productCount,
        salesGrowth: salesGrowth,
        ordersGrowth: ordersGrowth
      },
      top20: cleanTop20,
      customerTop10: customerTop10,
      categoryData: categoryData,
      churnRiskCustomers: churnRiskCustomers
    };

  } catch (error) {
    return {
      success: false,
      message: 'エラー: ' + error.toString()
    };
  }
}

/**
 * 前月を取得
 * @param {string} yearMonth - yyyy-MM形式
 * @return {string} 前月（yyyy-MM形式）
 */
function getPreviousMonth(yearMonth) {
  const parts = yearMonth.split('-');
  const year = parseInt(parts[0]);
  const month = parseInt(parts[1]);

  if (month === 1) {
    return `${year - 1}-12`;
  } else {
    const prevMonth = month - 1;
    return `${year}-${String(prevMonth).padStart(2, '0')}`;
  }
}

/**
 * 年間売上推移データ取得
 * @param {number} year - 対象年（例: 2026）
 * @return {Object} 年間データ
 */
function getYearlySalesData(year) {
  try {
    if (!year) {
      year = new Date().getFullYear();
    }

    // 月次集計データ取得
    const aggregationData = getAllRecords('月次集計');

    if (!aggregationData || aggregationData.length === 0) {
      return {
        success: false,
        message: '月次集計データがありません'
      };
    }

    // 12ヶ月分のデータを初期化
    const monthlyData = {};
    for (let month = 1; month <= 12; month++) {
      const yearMonth = `${year}-${String(month).padStart(2, '0')}`;
      monthlyData[yearMonth] = {
        totalSales: 0,
        totalQuantity: 0,
        totalOrders: 0,
        productCount: 0,
        products: {}
      };
    }

    // データを集計
    aggregationData.forEach(d => {
      let yearMonth = d['年月'];
      if (yearMonth instanceof Date) {
        yearMonth = Utilities.formatDate(yearMonth, 'JST', 'yyyy-MM');
      }

      if (!yearMonth || !yearMonth.startsWith(year.toString())) {
        return;
      }

      const productName = d['商品名'] || '';

      // 送料・代引手数料を除外
      if (productName === '送料' || productName === '代引手数料') {
        return;
      }

      const sales = Number(d['売上高']) || 0;
      const quantity = Number(d['受注数合計']) || 0;
      const orders = Number(d['受注回数']) || 0;

      if (monthlyData[yearMonth]) {
        monthlyData[yearMonth].totalSales += sales;
        monthlyData[yearMonth].totalQuantity += quantity;
        monthlyData[yearMonth].totalOrders += orders;
        monthlyData[yearMonth].productCount += 1;

        // 商品別データも保存
        if (!monthlyData[yearMonth].products[productName]) {
          monthlyData[yearMonth].products[productName] = {
            sales: 0,
            quantity: 0,
            orders: 0
          };
        }
        monthlyData[yearMonth].products[productName].sales += sales;
        monthlyData[yearMonth].products[productName].quantity += quantity;
        monthlyData[yearMonth].products[productName].orders += orders;
      }
    });

    // 全商品リストを取得
    const allProducts = new Set();
    Object.values(monthlyData).forEach(month => {
      Object.keys(month.products).forEach(product => {
        allProducts.add(product);
      });
    });

    return {
      success: true,
      year: year,
      monthlyData: monthlyData,
      products: Array.from(allProducts).sort()
    };

  } catch (error) {
    return {
      success: false,
      message: 'エラー: ' + error.toString()
    };
  }
}

/**
 * クライアント転送用にDateオブジェクトを文字列に変換
 * @param {Array} records - レコード配列
 * @return {Array} Date変換済みレコード配列
 */
function convertDatesForClient(records) {
  if (!records || !Array.isArray(records)) {
    return records;
  }

  return records.map(record => {
    const cleanRecord = {};
    for (const key in record) {
      if (record[key] instanceof Date) {
        cleanRecord[key] = Utilities.formatDate(record[key], 'JST', 'yyyy-MM-dd');
      } else {
        cleanRecord[key] = record[key];
      }
    }
    return cleanRecord;
  });
}

/**
 * 分析データ取得（クライアント転送用にDate変換済み）
 * @param {string} sheetName - シート名
 * @return {Array} Date変換済みレコード配列
 */
function getAnalysisRecords(sheetName) {
  try {
    const records = getAllRecords(sheetName);

    if (!records || records.length === 0) {
      return [];
    }

    return convertDatesForClient(records);

  } catch (error) {
    return [];
  }
}

