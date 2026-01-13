/**
 * CSVファイルインポートのメイン関数
 *
 * 各ECプラットフォームからのCSVデータを受注シートに取り込む機能を提供します。
 * ファイル種別に応じて適切な処理関数を呼び出します。
 *
 * 処理フロー:
 * 1. JSON文字列をパースしてデータオブジェクトに変換
 * 2. データ内容をログに記録
 * 3. ファイル種別（fileClass）により分岐:
 *    - '1': 食べチョクCSVデータをtabeAddRecord()で処理
 *    - '2': ふるさと納税CSVデータをfuruAddRecord()で処理
 *    - '3': ポケマルCSVデータをmaruSheAddRecord()で処理
 *
 * @param {string} fileClass - ファイル種別を示す文字列 ('1':食べチョク, '2':ふるさと納税, '3':ポケマル)
 * @param {string} datas - JSON形式の文字列（CSVから変換されたデータ配列）
 * @returns {void} 返却値なし
 *
 * @see tabeAddRecord - 食べチョクデータの処理
 * @see furuAddRecord - ふるさと納税データの処理
 * @see maruSheAddRecord - ポケマルデータの処理
 *
 * @example
 * // 食べチョクのCSVデータをインポート
 * const jsonData = JSON.stringify([{注文番号: '12345', ...}]);
 * csvFileImport('1', jsonData);
 */
function csvFileImport(fileClass, datas) {
  var data = JSON.parse(datas);
  Logger.log(data);
  // 食べチョク
  if (fileClass == '1') {
    tabeAddRecord(data);
  }
  // ふるさと納税
  if (fileClass == '2') {
    furuAddRecord(data);
  }
  // ポケマル
  if (fileClass == '3') {
    maruSheAddRecord(data);
  }
}
/**
 * 食べチョクの日付フォーマット変換関数
 *
 * 食べチョクCSVの8桁日付形式（YYYYMMDD）を標準的な日付形式（YYYY/MM/DD）に変換します。
 *
 * 処理フロー:
 * 1. 入力値がnullまたは空文字の場合、空文字を返却
 * 2. 文字列長が8桁の場合:
 *    - 年（0-3文字目）を抽出
 *    - 月（4-5文字目）を抽出
 *    - 日（6-7文字目）を抽出
 *    - スラッシュ区切りで結合
 * 3. 8桁以外の場合は空文字を返却
 *
 * @param {string} str - 8桁の日付文字列（例: '20231225'）
 * @returns {string} YYYY/MM/DD形式の日付文字列（例: '2023/12/25'）、または空文字
 *
 * @see tabeAddRecord - この関数を使用して注文日と発送日を変換
 *
 * @example
 * tabeDateFormat('20231225'); // '2023/12/25'
 * tabeDateFormat('');         // ''
 * tabeDateFormat(null);       // ''
 * tabeDateFormat('2023');     // ''（8桁でないため）
 */
function tabeDateFormat(str) {
  if (str == null || str == '') {
    return '';
  }
  if (str.length == 8) {
    return str.substring(0, 4) + '/' + str.substring(4, 6) + '/' + str.substring(6, 8);
  }
  else {
    return '';
  }
}
/**
 * 食べチョクCSVデータを受注シートに追加する関数
 *
 * 食べチョクからのCSVデータを解析し、内部の受注フォーマットに変換して受注シートに一括登録します。
 * 商品マスタを参照して商品情報を変換し、会社情報やお届け先情報を含む完全なレコードを作成します。
 *
 * 処理フロー:
 * 1. 空の追加用配列（adds）を初期化
 * 2. 「食べチョク商品」シートから全商品マスタを取得
 * 3. 各CSVレコードに対して以下を実行:
 *    a. 注文番号が空またはnullの場合はスキップ
 *    b. 商品名で商品マスタを検索し、変換値・変換数量を取得
 *    c. 受注レコード配列を作成（注文情報、お届け先、発送元、配送情報など40項目）
 *    d. 日付フィールドはtabeDateFormat()で変換後、Utilities.formatDate()でフォーマット
 *    e. 金額は商品代金+送料+クール代の合計を計算
 *    f. 作成したレコードを配列に追加
 * 4. 作成したレコード配列をログに記録
 * 5. addRecords()で「受注」シートに一括追加
 *
 * @param {Object[]} data - 食べチョクCSVから変換されたオブジェクト配列
 * @param {string} data[].注文番号 - 注文を一意に識別する番号
 * @param {string} data[].注文日 - YYYYMMDD形式の注文日
 * @param {string} data[].商品名 - 商品マスタと紐付けるための商品名
 * @param {number} data[].商品代金(お客様に提示した額) - 商品本体価格
 * @param {number} data[].生産者に支払われる送料 - 送料
 * @param {number} data[].生産者に支払われるクール代 - クール便料金
 * @param {string} data[].お届け先名 - 配送先の宛名
 * @param {string} data[].お届け先郵便番号 - 配送先郵便番号
 * @param {string} data[].お届け先住所 - 配送先住所
 * @param {string} data[].お届け先電話番号 - 配送先電話番号
 * @param {string} data[].発送日 - YYYYMMDD形式の発送日
 * @param {string} data[].お届け日 - YYYYMMDD形式のお届け日
 * @param {string} data[].配送業者 - 配送業者名
 * @param {string} data[].お届け時間帯 - 希望配送時間帯
 * @param {string} data[].発払い - 配送料金の支払い区分
 * @param {string} data[].クール種別 - クール便の種別（冷蔵/冷凍）
 * @param {string} data[].特記事項 - その他の特記事項
 * @returns {void} 返却値なし（受注シートへの追加を実行）
 *
 * @see tabeDateFormat - 日付フォーマット変換に使用
 * @see getAllRecords - 商品マスタの取得に使用
 * @see addRecords - 受注シートへのレコード追加に使用
 * @see CONFIG.TABECHOKU - 食べチョク固有の発送元情報
 * @see CONFIG.COMPANY - 自社の発送元情報
 *
 * @example
 * const csvData = [{
 *   注文番号: 'T12345',
 *   注文日: '20231225',
 *   商品名: 'いちご',
 *   商品代金: 3000,
 *   ...
 * }];
 * tabeAddRecord(csvData);
 *
 * @note 金額は商品代金+送料+クール代の合計が計算されます
 * @note 商品マスタに存在しない商品名の場合はエラーが発生します
 */
function tabeAddRecord(data) {
  const adds = [];
  const productItems = getAllRecords('食べチョク商品');
  data.forEach(function (wVal) {
    if (wVal['注文番号'] == '' || wVal['注文番号'] == null) {
      return;
    }
    const productData = productItems.find((ele) => ele['商品名'] == wVal['商品名']);

    const record = [
      wVal['注文番号'],
      Utilities.formatDate(new Date(tabeDateFormat(wVal['注文日'])), 'JST', 'yyyy/MM/dd'),
      '',
      productData['変換値'],
      productData['変換数量'],
      Number(wVal['商品代金(お客様に提示した額)']) + Number(wVal['生産者に支払われる送料']) + Number(wVal['生産者に支払われるクール代']),
      CONFIG.TABECHOKU.COMPANY,
      CONFIG.TABECHOKU.ZIPCODE,
      CONFIG.TABECHOKU.ADDRESS,
      '',
      wVal['お届け先名'],
      wVal['お届け先郵便番号'],
      wVal['お届け先住所'],
      wVal['お届け先電話番号'],
      CONFIG.COMPANY.NAME,
      CONFIG.COMPANY.ZIPCODE,
      CONFIG.COMPANY.ADDRESS,
      CONFIG.COMPANY.TEL,
      Utilities.formatDate(new Date(tabeDateFormat(wVal['発送日'])), 'JST', 'yyyy/MM/dd'),
      Utilities.formatDate(new Date(tabeDateFormat(wVal['お届け日'])), 'JST', 'yyyy/MM/dd'),
      '',
      '',
      wVal['配送業者'],
      wVal['お届け時間帯'],
      '',
      '',
      '',
      '',
      '',
      '',
      '',
      wVal['発払い'],
      wVal['クール種別'],
      '',
      '',
      '',
      '',
      '',
      '',
      '',
      wVal['特記事項'],
      Number(wVal['商品代金(お客様に提示した額)']) + Number(wVal['生産者に支払われる送料']) + Number(wVal['生産者に支払われるクール代'])
    ];
    adds.push(record);
  });
  Logger.log(adds);
  addRecords('受注', adds);
}
/**
 * ふるさと納税CSVデータを受注シートに追加する関数
 *
 * ふるさと納税プラットフォームからのCSVデータを解析し、内部の受注フォーマットに変換して受注シートに一括登録します。
 * 商品マスタを参照して商品情報を変換し、個数に応じた数量・金額計算を行います。
 *
 * 処理フロー:
 * 1. 空の追加用配列（adds）を初期化
 * 2. 「ふるさと納税商品」シートから全商品マスタを取得
 * 3. 各CSVレコードに対して以下を実行:
 *    a. 配送管理IDが空またはnullの場合はスキップ
 *    b. 返礼品名で商品マスタを検索し、変換値・変換数量を取得
 *    c. 受注レコード配列を作成（配送管理ID、日付、商品情報、配送先など40項目）
 *    d. 注文日は現在日時を設定
 *    e. 変換数量は商品マスタの変換数量×個数で計算
 *    f. お届け先住所は住所１と住所２を結合
 *    g. 発送元・依頼主はともにCONFIG.FURUSATOの情報を使用
 *    h. 合計金額は変換数量×個数×提供価格で計算
 *    i. 作成したレコードを配列に追加
 * 4. 作成したレコード配列をログに記録
 * 5. addRecords()で「受注」シートに一括追加
 *
 * @param {Object[]} data - ふるさと納税CSVから変換されたオブジェクト配列
 * @param {string} data[].配送管理ID - 配送を一意に識別するID
 * @param {string} data[].返礼品 - 商品マスタと紐付けるための返礼品名
 * @param {number} data[].個数 - 注文個数（変換数量の乗数として使用）
 * @param {number} data[].提供価格 - 返礼品の単価
 * @param {string} data[].お届け先名 - 配送先の宛名
 * @param {string} data[].届け先郵便番号 - 配送先郵便番号
 * @param {string} data[].届け先住所１ - 配送先住所（市区町村まで）
 * @param {string} data[].届け先住所２ - 配送先住所（番地以降）
 * @param {string} data[].届け先電話番号 - 配送先電話番号
 * @param {Date|string} data[].発送日 - 発送日（Dateオブジェクトまたは日付文字列）
 * @param {string} data[].ヤマト産直伝票 - ヤマト産直サービスの伝票情報
 * @returns {void} 返却値なし（受注シートへの追加を実行）
 *
 * @see getAllRecords - 商品マスタの取得に使用
 * @see addRecords - 受注シートへのレコード追加に使用
 * @see CONFIG.FURUSATO - ふるさと納税固有の発送元・依頼主情報
 *
 * @example
 * const csvData = [{
 *   配送管理ID: 'F12345',
 *   返礼品: 'いちご詰め合わせ',
 *   個数: 2,
 *   提供価格: 10000,
 *   発送日: '2023-12-25',
 *   ...
 * }];
 * furuAddRecord(csvData);
 *
 * @note 注文日は常に現在日時が設定されます
 * @note 数量計算: 変換数量 = 商品マスタの変換数量 × 個数
 * @note 金額計算: 合計金額 = 変換数量 × 個数 × 提供価格
 * @note 住所は届け先住所１と届け先住所２が結合されます
 * @note 商品マスタに存在しない返礼品名の場合はエラーが発生します
 */
function furuAddRecord(data) {
  const adds = [];
  const productItems = getAllRecords('ふるさと納税商品');
  data.forEach(function (wVal) {
    if (wVal['配送管理ID'] == '' || wVal['配送管理ID'] == null) {
      return;
    }
    const productData = productItems.find((ele) => ele['商品名'] == wVal['返礼品']);
    const record = [
      wVal['配送管理ID'],
      Utilities.formatDate(new Date(), 'JST', 'yyyy/MM/dd'),
      '',
      productData['変換値'],
      Number(productData['変換数量'])*Number(wVal['個数']),
      wVal['提供価格'],
      CONFIG.FURUSATO.COMPANY,
      CONFIG.FURUSATO.ZIPCODE,
      CONFIG.FURUSATO.ADDRESS,
      CONFIG.FURUSATO.TEL,
      wVal['お届け先名'],
      wVal['届け先郵便番号'],
      wVal['届け先住所１'] + wVal['届け先住所２'],
      wVal['届け先電話番号'],
      CONFIG.FURUSATO.COMPANY,
      CONFIG.FURUSATO.ZIPCODE,
      CONFIG.FURUSATO.ADDRESS,
      CONFIG.FURUSATO.TEL,
      Utilities.formatDate(new Date(wVal['発送日']), 'JST', 'yyyy/MM/dd'),
      '',
      '',
      '',
      wVal['ヤマト産直伝票'],
      '',
      '',
      '',
      '',
      '',
      '',
      '',
      '',
      '',
      '',
      '',
      '',
      '',
      '',
      '',
      '',
      '',
      '',
      Number(productData['変換数量'])*Number(wVal['個数']) * Number(wVal['提供価格'])
    ];
    adds.push(record);
  });
  Logger.log(adds);
  addRecords('受注', adds);
}
/**
 * ポケマルの日付フォーマット変換関数
 *
 * ポケマルCSVの日付形式（MM/DD形式など）を現在年のDateオブジェクトに変換します。
 * 年情報がない月日形式の日付を、現在の年と組み合わせて完全な日付に変換する特殊処理を行います。
 *
 * 処理フロー:
 * 1. 入力値がnullまたは空文字の場合、現在日時を返却
 * 2. 現在日時のDateオブジェクトを取得
 * 3. 文字列長が5文字より大きい場合（通常のMM/DD形式）:
 *    a. 先頭5文字（MM/DD部分）を抽出
 *    b. 仮の年（2020）と組み合わせて日付文字列を作成
 *    c. Dateオブジェクトに変換
 *    d. setFullYear()で現在の年に更新
 *    e. 変換されたDateオブジェクトを返却
 * 4. 5文字以下の場合は現在日時を返却
 *
 * @param {string} str - ポケマルCSVの日付文字列（例: '12/25'、'01/15'）
 * @returns {Date} 現在年と組み合わせたDateオブジェクト、または現在日時
 *
 * @see maruSheAddRecord - この関数を使用して注文確定日と発送日を変換
 *
 * @example
 * // 2023年12月に実行した場合
 * maruDateFormat('12/25'); // Date: 2023/12/25
 * maruDateFormat('01/15'); // Date: 2023/01/15
 * maruDateFormat('');      // Date: 現在日時
 * maruDateFormat(null);    // Date: 現在日時
 * maruDateFormat('1/1');   // Date: 現在日時（5文字以下のため）
 *
 * @note ポケマルのCSVには年情報がないため、常に現在の年が適用されます
 * @note 年末年始をまたぐ場合の年の扱いには注意が必要です
 * @note 仮の年として2020を使用するのは、Date変換の便宜上の処理です
 */
function maruDateFormat(str) {
  if (str == null || str == '') {
    return new Date();
  }
  var maruDate = new Date();
  if (str.length > 5) {
    const dateStr = '2020/'+ str.substring(0, 5);
    let dateObject = new Date(dateStr);
    dateObject.setFullYear(maruDate.getFullYear());
    return dateObject;
  }
  else {
    return maruDate;
  }
}
/**
 * ポケマルCSVデータを受注シートに追加する関数
 *
 * ポケットマルシェからのCSVデータを解析し、内部の受注フォーマットに変換して受注シートに一括登録します。
 * 商品マスタを商品名と分量の両方で検索し、配送先住所を複数フィールドから結合して作成します。
 *
 * 処理フロー:
 * 1. 空の追加用配列（adds）を初期化
 * 2. 「ポケマル商品」シートから全商品マスタを取得
 * 3. 各CSVレコードに対して以下を実行:
 *    a. 注文IDが空またはnullの場合はスキップ
 *    b. 商品名と分量の両方の条件で商品マスタを検索し、変換値・変換数量を取得
 *    c. 受注レコード配列を作成（注文ID、日付、商品情報、配送先など40項目）
 *    d. 注文確定日と発送日はmaruDateFormat()で現在年の日付に変換
 *    e. 配送先住所は都道府県+市区町村+番地+建物名を結合
 *    f. 発送元はCONFIG.POCKEMARU、依頼主はCONFIG.COMPANYの情報を使用
 *    g. 作成したレコードを配列に追加
 * 4. 作成したレコード配列をログに記録
 * 5. addRecords()で「受注」シートに一括追加
 *
 * @param {Object[]} data - ポケマルCSVから変換されたオブジェクト配列
 * @param {string} data[].注文ID - 注文を一意に識別するID
 * @param {string} data[].注文確定日 - MM/DD形式の注文確定日
 * @param {string} data[].商品名 - 商品マスタと紐付けるための商品名
 * @param {string} data[].分量 - 商品マスタと紐付けるための分量情報（商品名と組み合わせて検索）
 * @param {string} data[].配送先宛名 - 配送先の宛名
 * @param {string} data[].配送先〒 - 配送先郵便番号
 * @param {string} data[].配送先都道府県 - 配送先都道府県
 * @param {string} data[].配送先市区町村 - 配送先市区町村
 * @param {string} data[].配送先番地 - 配送先番地
 * @param {string} data[].配送先建物名・部屋番号 - 配送先建物名・部屋番号
 * @param {string} data[].配送先電話番号 - 配送先電話番号
 * @param {string} data[].発送日 - MM/DD形式の発送日
 * @param {string} data[].着日指定 - 配送希望日
 * @param {string} data[].ヤマト産直伝票 - ヤマト産直サービスの伝票情報
 * @param {string} data[].申し送り事項 - その他の特記事項
 * @returns {void} 返却値なし（受注シートへの追加を実行）
 *
 * @see maruDateFormat - 日付フォーマット変換に使用（年なしMM/DD形式を現在年のDateに変換）
 * @see getAllRecords - 商品マスタの取得に使用
 * @see addRecords - 受注シートへのレコード追加に使用
 * @see CONFIG.POCKEMARU - ポケマル固有の発送元情報
 * @see CONFIG.COMPANY - 自社の依頼主情報
 *
 * @example
 * const csvData = [{
 *   注文ID: 'P12345',
 *   注文確定日: '12/25',
 *   商品名: 'いちご',
 *   分量: '1kg',
 *   配送先都道府県: '東京都',
 *   配送先市区町村: '渋谷区',
 *   ...
 * }];
 * maruSheAddRecord(csvData);
 *
 * @note 商品マスタの検索は商品名と分量の両方が一致する必要があります
 * @note 日付は年情報がないため、常に現在年が適用されます
 * @note 配送先住所は4つのフィールド（都道府県+市区町村+番地+建物名）が結合されます
 * @note 商品マスタに存在しない商品名・分量の組み合わせの場合はエラーが発生します
 */
function maruSheAddRecord(data) {
  const adds = [];
  const productItems = getAllRecords('ポケマル商品');
  data.forEach(function (wVal) {
    if (wVal['注文ID'] == '' || wVal['注文ID'] == null) {
      return;
    }
    const productData = productItems.find((ele) => ele['商品名'] == wVal['商品名'] && ele['分量'] == wVal['分量']);
    const record = [
      wVal['注文ID'],
      Utilities.formatDate(maruDateFormat(wVal['注文確定日']), 'JST', 'yyyy/MM/dd'),
      '',
      productData['変換値'],
      productData['変換数量'],
      '',
      CONFIG.POCKEMARU.COMPANY,
      CONFIG.POCKEMARU.ZIPCODE,
      CONFIG.POCKEMARU.ADDRESS,
      '',
      wVal['配送先宛名'],
      wVal['配送先〒'],
      wVal['配送先都道府県'] + wVal['配送先市区町村'] + wVal['配送先番地'] + wVal['配送先建物名・部屋番号'],
      wVal['配送先電話番号'],
      CONFIG.COMPANY.NAME,
      CONFIG.COMPANY.ZIPCODE,
      CONFIG.COMPANY.ADDRESS,
      CONFIG.COMPANY.TEL,
      Utilities.formatDate(maruDateFormat(wVal['発送日']), 'JST', 'yyyy/MM/dd'),
      wVal['着日指定'],
      '',
      '',
      wVal['ヤマト産直伝票'],
      '',
      '',
      '',
      '',
      '',
      '',
      '',
      '',
      '',
      '',
      '',
      '',
      '',
      '',
      '',
      '',
      '',
      wVal['申し送り事項'],
      ''
    ];
    adds.push(record);
  });
  Logger.log(adds);
  addRecords('受注', adds);
}