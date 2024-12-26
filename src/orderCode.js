// 納品書のテンプレートファイル
const DELIVERED_TEMPLATE = DriveApp.getFileById("####ドライブIDを指定####");
// 納品書PDF出力先
const DELIVERED_PDF_OUTDIR = DriveApp.getFolderById("####ドライブIDを指定####");
// 領収書のテンプレートファイル
const RECEIPT_TEMPLATE = DriveApp.getFileById("####ドライブIDを指定####");
// 領収書PDF出力先
const RECEIPT_PDF_OUTDIR = DriveApp.getFolderById("####ドライブIDを指定####");

// 受注管理画面
function getshippingHTML(e, alert = "") {
  const items = getAllRecords("商品");
  const recipients = getAllRecords("担当者");
  const deliveryMethods = getAllRecords("納品方法");
  const receipts = getAllRecords("受付方法");
  const deliveryTimes = getAllRecords("配送時間帯");
  const invoiceTypes = getAllRecords("送り状種別");
  const coolClss = getAllRecords("クール区分");
  const cargos = getAllRecords("荷扱い");
  const customers = getAllRecords("顧客情報");
  var nDate = new Date();
  var strDate = Utilities.formatDate(nDate, "JST", "yyyy-MM-dd");
  const orderDate = e.parameter.orderDate ? e.parameter.orderDate : strDate;
  var n = 2;
  nDate.setDate(nDate.getDate() + n);
  strDate = Utilities.formatDate(nDate, "JST", "yyyy-MM-dd");
  const shippingDate = e.parameter.shippingDate
    ? e.parameter.shippingDate
    : strDate;
  nDate.setDate(nDate.getDate() + 1);
  strDate = Utilities.formatDate(nDate, "JST", "yyyy-MM-dd");
  const deliveryDate = e.parameter.deliveryDate
    ? e.parameter.deliveryDate
    : strDate;

  let html = ``;
  html += `<p class="text-danger">${alert}</p>`;
  html += `<div style="background-color: magenta; color: white;">【顧客情報】　`;
  html += `      <button type='button' class="customerInsertBtn_open">新規登録</button>`;
  html += `      <button type='button' class="customerSearchBtn_open">顧客検索</button>`;
  html += `      <button type='button' class="customerSameBtn" onclick="customerSame()">発送先同上</button>`;
  html += `      <button type='button' id="productSearch" onclick="setProductSearch()">前回商品反映</button>`;
  html += `      <button type='button' id="quotationSearch" onclick="setQuotationSearch()">見積書反映</button>`;
  html += `</div>`;
  html += `<div>
                <label for="customerName" class="text-left form-label">顧客名</label>`;
  html += `<input type="text" class="form-control" id="customerName" name="customerName" required value="${
    e.parameter.customerName ? e.parameter.customerName : ""
  }" >`;
  html += `</div>`;
  html += `<div class="mt-0 mb-2 row g-3 align-items-center">`;
  html += `    <div class="col-auto">`;
  html += `        <label for="customerZipcode" class="col-form-label">郵便番号</label>`;
  html += `    </div>`;
  html += `    <div class="col-auto">`;
  html += `        <input type="text" class="form-control" id="customerZipcode" name="customerZipcode" required value="${
    e.parameter.customerZipcode ? e.parameter.customerZipcode : ""
  }" maxlength=7 pattern="[0-9]{7}" title="7桁の数字のみを入力してください。">`;
  html += `    </div>`;
  html += `</div>`;
  html += `<div>
                <label for="customerAddress" class="text-left form-label">住所</label>`;
  html += `<input type="text" class="form-control" id="customerAddress" name="customerAddress" required value="${
    e.parameter.customerAddress ? e.parameter.customerAddress : ""
  }" >`;
  html += `</div>`;
  html += `<div class="mt-0 mb-2 row g-3 align-items-center">`;
  html += `    <div class="col-auto">`;
  html += `        <label for="customerTel" class="col-form-label">電話番号</label>`;
  html += `    </div>`;
  html += `    <div class="col-auto">`;
  html += `        <input type="text" class="form-control" id="customerTel" name="customerTel" required value="${
    e.parameter.customerTel ? e.parameter.customerTel : ""
  }" maxlength=11 pattern="^0[0-9]{9,10}$" title="10~11桁の数字のみを入力してください。">`;
  html += `    </div>`;
  html += `</div>`;
  html += `<div style="background-color: darkgoldenrod; color: white;">【発送先情報】　`;
  html += `      <button type='button' class="shippingToInsertBtn_open">新規登録</button>`;
  html += `      <button type='button' class="shippingToSearchBtn_open">発送先検索</button>`;
  html += `</div>`;
  html += `<div>
                <label for="shippingToName" class="text-left form-label">発送先名</label>`;
  html += `<input type="text" class="form-control" id="shippingToName" name="shippingToName" required value="${
    e.parameter.shippingToName ? e.parameter.shippingToName : ""
  }" >`;
  html += `</div>`;
  html += `<div class="mt-0 mb-2 row g-3 align-items-center">`;
  html += `    <div class="col-auto">`;
  html += `        <label for="shippingToZipcode" class="col-form-label">郵便番号</label>`;
  html += `    </div>`;
  html += `    <div class="col-auto">`;
  html += `        <input type="text" class="form-control" id="shippingToZipcode" name="shippingToZipcode" required value="${
    e.parameter.shippingToZipcode ? e.parameter.shippingToZipcode : ""
  }" maxlength=7 pattern="[0-9]{7}" title="7桁の数字のみを入力してください。">`;
  html += `    </div>`;
  html += `</div>`;
  html += `<div>
                <label for="shippingToAddress" class="text-left form-label">住所</label>`;
  html += `<input type="text" class="form-control" id="shippingToAddress" name="shippingToAddress" required value="${
    e.parameter.shippingToAddress ? e.parameter.shippingToAddress : ""
  }" >`;
  html += `</div>`;
  html += `<div class="mt-0 mb-2 row g-3 align-items-center">`;
  html += `    <div class="col-auto">`;
  html += `        <label for="shippingToTel" class="col-form-label">電話番号</label>`;
  html += `    </div>`;
  html += `    <div class="col-auto">`;
  html += `        <input type="text" class="form-control" id="shippingToTel" name="shippingToTel" required value="${
    e.parameter.shippingToTel ? e.parameter.shippingToTel : ""
  }" maxlength=11 pattern="^0[0-9]{9,10}$" title="10~11桁の数字のみを入力してください。">`;
  html += `    </div>`;
  html += `</div>`;
  html += `<div style="background-color: brown; color: white;">【発送元情報】　`;
  html += ` <button type="button" class="" id="farmBtn" name="farmBtn" value="true" onclick="farmChange()">会社名</button>`;
  html += ` <button type="button" class="" id="custCopyBtn" name="custCopyBtn" value="true" onclick="custCopy()">顧客コピー</button>`;
  html += ` <button type="button" class="" id="sendCopyBtn" name="sendCopyBtn" value="true" onclick="sendCopy()">発送先コピー</button>`;
  html += `</div>`;
  html += `<div>
                <label for="shippingFromName" class="text-left form-label">発送元名</label>`;
  html += `<input type="text" class="form-control" id="shippingFromName" name="shippingFromName" required value="${
    e.parameter.shippingFromName ? e.parameter.shippingFromName : ""
  }" >`;
  html += `</div>`;
  html += `<div class="mt-0 mb-2 row g-3 align-items-center">`;
  html += `    <div class="col-auto">`;
  html += `        <label for="shippingFromZipcode" class="col-form-label">郵便番号</label>`;
  html += `    </div>`;
  html += `    <div class="col-auto">`;
  html += `        <input type="text" class="form-control" id="shippingFromZipcode" name="shippingFromZipcode" required value="${
    e.parameter.shippingFromZipcode ? e.parameter.shippingFromZipcode : ""
  }" maxlength=7 pattern="[0-9]{7}" title="7桁の数字のみを入力してください。">`;
  html += `    </div>`;
  html += `</div>`;
  html += `<div>
                <label for="shippingFromAddress" class="text-left form-label">住所</label>`;
  html += `<input type="text" class="form-control" id="shippingFromAddress" name="shippingFromAddress" required value="${
    e.parameter.shippingFromAddress ? e.parameter.shippingFromAddress : ""
  }" >`;
  html += `</div>`;
  html += `<div class="mt-0 mb-2 row g-3 align-items-center">`;
  html += `    <div class="col-auto">`;
  html += `        <label for="shippingFromTel" class="col-form-label">電話番号</label>`;
  html += `    </div>`;
  html += `    <div class="col-auto">`;
  html += `        <input type="text" class="form-control" id="shippingFromTel" name="shippingFromTel" required value="${
    e.parameter.shippingFromTel ? e.parameter.shippingFromTel : ""
  }" maxlength=11 pattern="^0[0-9]{9,10}$" title="10~11桁の数字のみを入力してください。">`;
  html += `    </div>`;
  html += `</div>`;
  html += `<div style="background-color: darkcyan; color: white;">【受注基本情報】</div>`;
  html += `<div class="mt-0 mb-2 row g-3 align-items-center">`;
  html += `    <div class="col-auto">`;
  html += `        <label for="shippingDate" class="col-form-label">発送日</label>`;
  html += `    </div>`;
  html += `    <div class="col-auto">`;
  html += `       <input type="date" class="form-control" id="shippingDate" name="shippingDate" required value="${shippingDate}" >`;
  html += `    </div>`;
  html += `    <div class="col-auto">`;
  html += `        <label for="deliveryDate" class="col-form-label">納品日</label>`;
  html += `    </div>`;
  html += `    <div class="col-auto">`;
  html += `       <input type="date" class="form-control" id="deliveryDate" name="deliveryDate" required value="${deliveryDate}" >`;
  html += `    </div>`;
  html += `</div>`;
  html += `<div>`;
  html += `<label for="receiptWay" class="text-left form-label">受付方法</label>`;
  html += `<select class="form-select" id="receiptWay" name="receiptWay" required>`;
  for (const receipt of receipts) {
    const receiptWay = receipt["受付方法"];
    if (receiptWay == e.parameter.receiptWay) {
      html += `<option value="${receiptWay}" selected>${receiptWay}</option>`;
    } else {
      html += `<option value="${receiptWay}">${receiptWay}</option>`;
    }
  }
  html += `</select>`;
  html += `</div>`;
  html += `<div>`;
  html += `<label for="recipient" class="text-left form-label">受付者</label>`;
  html += `<select class="form-select" id="recipient" name="recipient" required>`;
  for (const recipient of recipients) {
    const recip = recipient["名前"];
    if (recip == e.parameter.recipient) {
      html += `<option value="${recip}" selected>${recip}</option>`;
    } else {
      html += `<option value="${recip}">${recip}</option>`;
    }
  }
  html += `</select>`;
  html += `</div>`;
  html += `<div class="mt-0 mb-2 row g-3 align-items-center">`;
  html += `    <div class="col-auto">`;
  html += `        <label for="deliveryMethod" class="col-form-label">納品方法</label>`;
  html += `    </div>`;
  html += `    <div class="col-auto">`;
  html += `       <select class="form-select" id="deliveryMethod" name="deliveryMethod" required onchange=deliveryMethodChange() >`;
  html += `         <option value=""></option>`;
  for (const deliveryMethod of deliveryMethods) {
    const deliMethod = deliveryMethod["納品方法"];
    if (deliMethod == e.parameter.deliveryMethod) {
      html += `<option value="${deliMethod}" selected>${deliMethod}</option>`;
    } else {
      html += `<option value="${deliMethod}">${deliMethod}</option>`;
    }
  }
  html += `    </select>`;
  html += `    </div>`;
  html += `    <div class="col-auto">`;
  html += `        <label for="deliveryTime" class="col-form-label">配達時間帯</label>`;
  html += `    </div>`;
  html += `    <div class="col-auto">`;
  html += `       <select class="form-select" id="deliveryTime" name="deliveryTime" >`;
  html += `         <option value=""></option>`;
  for (const deliveryTime of deliveryTimes) {
    const deliTime = deliveryTime["時間指定"];
    const deliTimeVal = deliveryTime["時間指定値"];
    const deliveryMethod = deliveryTime["納品方法"];
    if (deliTimeVal == e.parameter.deliveryTime) {
      html += `<option value="${deliTimeVal}" data-val="${deliveryMethod}" selected>${deliTime}</option>`;
    } else {
      html += `<option value="${deliTimeVal}" data-val="${deliveryMethod}">${deliTime}</option>`;
    }
  }
  html += `    </select>`;
  html += `    </div>`;
  html += `</div>`;
  html += `<div>`;
  html += `<div class="mt-0 mb-2 row g-3 align-items-center">`;
  html += `   <div class="col-auto">`;
  html += `     <div class="form-check">`;
  if (e.parameters.checklist && e.parameters.checklist.includes("納品書")) {
    html += `       <input class="form-check-input" type="checkbox" value="納品書" id="deliveryChk" name="checklist" checked>納品書`;
  } else {
    html += `       <input class="form-check-input" type="checkbox" value="納品書" id="deliveryChk" name="checklist" >納品書`;
  }
  html += `     </div>`;
  html += `   </div>`;
  html += `   <div class="col-auto">`;
  html += `     <div class="form-check">`;
  if (e.parameters.checklist && e.parameters.checklist.includes("請求書")) {
    html += `       <input class="form-check-input" type="checkbox" value="請求書" id="billChk" name="checklist" checked>請求書`;
  } else {
    html += `       <input class="form-check-input" type="checkbox" value="請求書" id="billChk" name="checklist" >請求書`;
  }
  html += `     </div>`;
  html += `   </div>`;
  html += `   <div class="col-auto">`;
  html += `     <div class="form-check">`;
  if (e.parameters.checklist && e.parameters.checklist.includes("領収書")) {
    html += `       <input class="form-check-input" type="checkbox" value="領収書" id="receiptChk"  name="checklist" checked>領収書`;
  } else {
    html += `       <input class="form-check-input" type="checkbox" value="領収書" id="receiptChk" name="checklist" >領収書`;
  }
  html += `     </div>`;
  html += `   </div>`;
  html += `   <div class="col-auto">`;
  html += `     <div class="form-check">`;
  if (e.parameters.checklist && e.parameters.checklist.includes("パンフ")) {
    html += `       <input class="form-check-input" type="checkbox" value="パンフ" id="pamphletChk" name="checklist" checked>パンフ`;
  } else {
    html += `       <input class="form-check-input" type="checkbox" value="パンフ" id="pamphletChk" name="checklist" >パンフ`;
  }
  html += `     </div>`;
  html += `   </div>`;
  html += `   <div class="col-auto">`;
  html += `     <div class="form-check">`;
  if (e.parameters.checklist && e.parameters.checklist.includes("レシピ")) {
    html += `       <input class="form-check-input" type="checkbox" value="レシピ" id="recipeChk" name="checklist" checked>レシピ`;
  } else {
    html += `       <input class="form-check-input" type="checkbox" value="レシピ" id="recipeChk" name="checklist" >レシピ`;
  }
  html += `     </div>`;
  html += `   </div>`;
  html += `</div>`;
  html += `<div>`;
  html += `   <label for="otherAttach" class="col-form-label">その他添付</label>`;
  html += `   <input type="text" class="form-control" id="otherAttach" name="otherAttach" value="${
    e.parameter.otherAttach ? e.parameter.otherAttach : ""
  }" >`;
  html += `</div>`;
  html += `<div style="background-color: forestgreen; color: white;" class="mt-3">【商品情報】</div>`;
  html += `
    <table class="table text-center">
      <thead>
        <tr>
          <th scope="col">商品分類</th>
          <th scope="col">商品名</th>
          <th scope="col">価格</th>
          <th scope="col">個数</th>
        </tr>
      </thead>
      <tbody>
  `;

  const categorys = getAllRecords("商品分類");
  var rowNum = 0;
  for (let i = 0; i < 10; i++) {
    rowNum++;
    html += `<tr>`;
    html += `<td>`;
    var bunrui = "bunrui" + rowNum;
    html += `<select class="form-select" id="${bunrui}" name="${bunrui}" onchange=bunruiChange(${rowNum}) >`;
    html += `<option value=""></option>`;
    for (const category of categorys) {
      if (e.parameter[bunrui] == category["商品分類"]) {
        html += `<option value="${category["商品分類"]}" selected>${category["商品分類"]}</option>`;
      } else {
        html += `<option value="${category["商品分類"]}" >${category["商品分類"]}</option>`;
      }
    }
    html += `</select>`;
    html += `</td>`;
    var product = "product" + rowNum;
    html += `<td>`;
    html += `<select class="form-select" id="${product}" name="${product}" onchange=productChange(${rowNum}) >`;
    html += `<option value="" data-val="" data-zaiko=""></option>`;
    for (const item of items) {
      if (item["在庫数"] > 0) {
        if (e.parameter[product] == item["商品名"]) {
          html += `<option value="${item["商品名"]}" data-val="${item["商品分類"]}" data-name="${item["商品名"]}" data-zaiko="${item["在庫数"]}" 
          data-price="${item["価格（P)"]}" selected>${item["商品名"]}</option>`;
        } else {
          html += `<option value="${item["商品名"]}" data-val="${item["商品分類"]}" data-name="${item["商品名"]}" data-zaiko="${item["在庫数"]}"
          data-price="${item["価格（P)"]}" >${item["商品名"]}</option>`;
        }
      }
    }
    html += `</select>`;
    html += `</td>`;
    var price = "price" + rowNum;
    html += `<td>`;
    html += `<input type="number" class="form-control no-spin" id="${price}" name="${price}" min='0'  value="${
      e.parameter[price] ? e.parameter[price] : ""
    }" >`;
    html += `</td>`;
    var quantity = "quantity" + rowNum;
    html += `<td>`;
    html += `<input type="number" class="form-control no-spin" id="${quantity}" name="${quantity}" min='0' max='999' step="0.1" title="整数部3桁小数部1桁の数字のみを入力してください。" value="${
      e.parameter[quantity] ? e.parameter[quantity] : ""
    }" >`;
    // html += `<select class="form-select" id="${quantity}" name="${quantity}" style="min-width: 65px;">`;
    // for (let i = 0; i <= 100; i++) {
    //   if (i == Number(e.parameter[quantity])) {
    //     html += `<option value="${i}" selected>${i}</option>`;
    //   } else {
    //     html += `<option value="${i}">${i}</option>`;
    //   }
    // }
    // html += `</select>`;
    html += `</td>`;
    html += `</tr>`;
  }
  html += `</tbody>`;
  html += `</table>`;
  html += `<div style="background-color: blue; color: white;">【発送情報】</div>`;
  html += `<div>
                <label for="sendProduct" class="text-left form-label">品名</label>`;
  html += `       <input type="text" class="form-control" id="sendProduct" name="sendProduct" value="${
    e.parameter.sendProduct ? e.parameter.sendProduct : ""
  }">`;
  html += `</div>`;
  html += `<div class="mt-0 mb-2 row g-3 align-items-center">`;
  html += `    <div class="col-auto">`;
  html += `        <label for="invoiceType" class="col-form-label">送り状種別</label>`;
  html += `    </div>`;
  html += `    <div class="col-auto">`;
  html += `<select class="form-select" id="invoiceType" name="invoiceType" >`;
  html += `<option value=""></option>`;
  for (const invoiceType of invoiceTypes) {
    if (e.parameter.invoiceType == invoiceType["種別値"]) {
      html += `<option value="${invoiceType["種別値"]}" data-val="${invoiceType["納品方法"]}" selected>${invoiceType["種別"]}</option>`;
    } else {
      html += `<option value="${invoiceType["種別値"]}" data-val="${invoiceType["納品方法"]}" >${invoiceType["種別"]}</option>`;
    }
  }
  html += `</select>`;
  html += `    </div>`;
  html += `    <div class="col-auto">`;
  html += `        <label for="coolCls" class="col-form-label">クール区分</label>`;
  html += `    </div>`;
  html += `    <div class="col-auto">`;
  html += `<select class="form-select" id="coolCls" name="coolCls" >`;
  html += `<option value=""></option>`;
  for (const coolCls of coolClss) {
    if (e.parameter.coolCls == coolCls["種別値"]) {
      html += `<option value="${coolCls["種別値"]}" data-val="${coolCls["納品方法"]}" selected>${coolCls["種別"]}</option>`;
    } else {
      html += `<option value="${coolCls["種別値"]}" data-val="${coolCls["納品方法"]}" >${coolCls["種別"]}</option>`;
    }
  }
  html += `</select>`;
  html += `    </div>`;
  html += `</div>`;
  html += `<div class="mt-0 mb-2 row g-3 align-items-center">`;
  html += `    <div class="col-auto">`;
  html += `        <label for="cargo1" class="col-form-label">荷扱い１</label>`;
  html += `    </div>`;
  html += `    <div class="col-auto">`;
  html += `<select class="form-select" id="cargo1" name="cargo1" >`;
  html += `<option value=""></option>`;
  for (const cargo of cargos) {
    if (e.parameter.cargo1 == cargo["種別値"]) {
      html += `<option value="${cargo["種別値"]}" data-val="${cargo["納品方法"]}" selected>${cargo["種別"]}</option>`;
    } else {
      html += `<option value="${cargo["種別値"]}" data-val="${cargo["納品方法"]}" >${cargo["種別"]}</option>`;
    }
  }
  html += `</select>`;
  html += `    </div>`;
  html += `    <div class="col-auto">`;
  html += `        <label for="cargo2" class="col-form-label">荷扱い２</label>`;
  html += `    </div>`;
  html += `    <div class="col-auto">`;
  html += `<select class="form-select" id="cargo2" name="cargo2" >`;
  html += `<option value=""></option>`;
  for (const cargo of cargos) {
    if (e.parameter.cargo2 == cargo["種別値"]) {
      html += `<option value="${cargo["種別値"]}" data-val="${cargo["納品方法"]}" selected>${cargo["種別"]}</option>`;
    } else {
      html += `<option value="${cargo["種別値"]}" data-val="${cargo["納品方法"]}" >${cargo["種別"]}</option>`;
    }
  }
  html += `</select>`;
  html += `    </div>`;
  html += `</div>`;
  html += `<div class="mt-0 mb-2 row g-3 align-items-center">`;
  html += `    <div class="col-auto">`;
  html += `        <label for="cashOnDelivery" class="col-form-label">代引総額</label>`;
  html += `    </div>`;
  html += `    <div class="col-auto">`;
  html += `       <input type="number" class="form-control" id="cashOnDelivery" name="cashOnDelivery" min='0'  value="${
    e.parameter.cashOnDelivery ? e.parameter.cashOnDelivery : ""
  }">`;
  html += `    </div>`;
  html += `    <div class="col-auto">`;
  html += `        <label for="cashOnDeliTax" class="col-form-label">代引内税</label>`;
  html += `    </div>`;
  html += `    <div class="col-auto">`;
  html += `       <input type="number" class="form-control" id="cashOnDeliTax" name="cashOnDeliTax" min='0'  value="${
    e.parameter.cashOnDeliTax ? e.parameter.cashOnDeliTax : ""
  }">`;
  html += `    </div>`;
  html += `</div>`;
  html += `<div class="mt-0 mb-2 row g-3 align-items-center">`;
  html += `    <div class="col-auto">`;
  html += `        <label for="copiePrint" class="col-form-label">発行枚数</label>`;
  html += `    </div>`;
  html += `    <div class="col-auto">`;
  html += `       <input type="number" class="form-control" id="copiePrint" name="copiePrint" min='0'  value="${
    e.parameter.copiePrint ? e.parameter.copiePrint : ""
  }">`;
  html += `    </div>`;
  html += `</div>`;
  html += `<div class="mb-3">`;
  html += `<label for="csvmemo" class="text-left form-label">送り状　備考欄</label>`;
  html += `<textarea class="form-control" id="csvmemo" name="csvmemo" rows="2" cols="30" maxlength="22">${
    e.parameter.csvmemo ? e.parameter.csvmemo : ""
  }</textarea>`;
  html += `</div>`;
  html += `<div class="mb-3">`;
  html += `<label for="deliveryMemo" class="text-left form-label">納品書　備考欄</label>`;
  html += `<textarea class="form-control" id="deliveryMemo" name="deliveryMemo" rows="3" cols="30" maxlength="90">${
    e.parameter.deliveryMemo ? e.parameter.deliveryMemo : ""
  }</textarea>`;
  html += `<div class="mb-3">`;
  html += `<label for="memo" class="text-left form-label">メモ</label>`;
  html += `<textarea class="form-control" id="memo" name="memo" rows="3" cols="30">${
    e.parameter.memo ? e.parameter.memo : ""
  }</textarea>`;
  html += `</div>`;
  html += `</div>`;
  return html;
}
// 受注確認画面
function getShippingComfirmHTML(e) {
  const items = getAllRecords("商品");
  const recipients = getAllRecords("担当者");
  const deliveryMethods = getAllRecords("納品方法");
  const receipts = getAllRecords("受付方法");
  const deliveryTimes = getAllRecords("配送時間帯");
  const invoiceTypes = getAllRecords("送り状種別");
  const coolClss = getAllRecords("クール区分");
  const cargos = getAllRecords("荷扱い");
  const customers = getAllRecords("顧客情報");
  var nDate = new Date();
  var strDate = Utilities.formatDate(nDate, "JST", "yyyy-MM-dd");
  const orderDate = e.parameter.orderDate ? e.parameter.orderDate : strDate;
  var n = 2;
  nDate.setDate(nDate.getDate() + n);
  strDate = Utilities.formatDate(nDate, "JST", "yyyy-MM-dd");
  const shippingDate = e.parameter.shippingDate
    ? e.parameter.shippingDate
    : strDate;
  nDate.setDate(nDate.getDate() + 1);
  strDate = Utilities.formatDate(nDate, "JST", "yyyy-MM-dd");
  const deliveryDate = e.parameter.deliveryDate
    ? e.parameter.deliveryDate
    : strDate;
  let html = ``;
  html += `<div style="background-color: magenta; color: white;">【顧客情報】</div>`;
  html += `<div>
                <label for="customerName" class="text-left form-label">顧客名</label>`;
  html += `<input type="text" class="form-control" id="customerName" name="customerName" required value="${
    e.parameter.customerName ? e.parameter.customerName : ""
  }" readonly>`;
  html += `</div>`;
  html += `<div class="mt-0 mb-2 row g-3 align-items-center">`;
  html += `    <div class="col-auto">`;
  html += `        <label for="customerZipcode" class="col-form-label">郵便番号</label>`;
  html += `    </div>`;
  html += `    <div class="col-auto">`;
  html += `        <input type="text" class="form-control" id="customerZipcode" name="customerZipcode" required value="${
    e.parameter.customerZipcode ? e.parameter.customerZipcode : ""
  }" maxlength=7 pattern="[0-9]{7}" title="7桁の数字のみを入力してください。" readonly>`;
  html += `    </div>`;
  html += `</div>`;
  html += `<div>
                <label for="customerAddress" class="text-left form-label">住所</label>`;
  html += `<input type="text" class="form-control" id="customerAddress" name="customerAddress" required value="${
    e.parameter.customerAddress ? e.parameter.customerAddress : ""
  }" readonly>`;
  html += `</div>`;
  html += `<div class="mt-0 mb-2 row g-3 align-items-center">`;
  html += `    <div class="col-auto">`;
  html += `        <label for="customerTel" class="col-form-label">電話番号</label>`;
  html += `    </div>`;
  html += `    <div class="col-auto">`;
  html += `        <input type="text" class="form-control" id="customerTel" name="customerTel" required value="${
    e.parameter.customerTel ? e.parameter.customerTel : ""
  }" maxlength=11 pattern="^0[0-9]{9,10}$" title="10~11桁の数字のみを入力してください。"readonly>`;
  html += `    </div>`;
  html += `</div>`;
  html += `<div style="background-color: darkgoldenrod; color: white;">【発送先情報】</div>`;
  html += `<div>
                <label for="shippingToName" class="text-left form-label">発送先名</label>`;
  html += `<input type="text" class="form-control" id="shippingToName" name="shippingToName" required value="${
    e.parameter.shippingToName ? e.parameter.shippingToName : ""
  }" readonly>`;
  html += `</div>`;
  html += `<div class="mt-0 mb-2 row g-3 align-items-center">`;
  html += `    <div class="col-auto">`;
  html += `        <label for="shippingToZipcode" class="col-form-label">郵便番号</label>`;
  html += `    </div>`;
  html += `    <div class="col-auto">`;
  html += `        <input type="text" class="form-control" id="shippingToZipcode" name="shippingToZipcode" required value="${
    e.parameter.shippingToZipcode ? e.parameter.shippingToZipcode : ""
  }" maxlength=7 pattern="[0-9]{7}" title="7桁の数字のみを入力してください。"readonly>`;
  html += `    </div>`;
  html += `</div>`;
  html += `<div>
                <label for="shippingToAddress" class="text-left form-label">住所</label>`;
  html += `<input type="text" class="form-control" id="shippingToAddress" name="shippingToAddress" required value="${
    e.parameter.shippingToAddress ? e.parameter.shippingToAddress : ""
  }" readonly>`;
  html += `</div>`;
  html += `<div class="mt-0 mb-2 row g-3 align-items-center">`;
  html += `    <div class="col-auto">`;
  html += `        <label for="shippingToTel" class="col-form-label">電話番号</label>`;
  html += `    </div>`;
  html += `    <div class="col-auto">`;
  html += `        <input type="text" class="form-control" id="shippingToTel" name="shippingToTel" required value="${
    e.parameter.shippingToTel ? e.parameter.shippingToTel : ""
  }" maxlength=11 pattern="^0[0-9]{9,10}$" title="10~11桁の数字のみを入力してください。"readonly>`;
  html += `    </div>`;
  html += `</div>`;
  html += `<div style="background-color: brown; color: white;">【発送元情報】</div>`;
  html += `<div>
                <label for="shippingFromName" class="text-left form-label">発送元名</label>`;
  html += `<input type="text" class="form-control" id="shippingFromName" name="shippingFromName" required value="${
    e.parameter.shippingFromName ? e.parameter.shippingFromName : ""
  }" readonly>`;
  html += `</div>`;
  html += `<div class="mt-0 mb-2 row g-3 align-items-center">`;
  html += `    <div class="col-auto">`;
  html += `        <label for="shippingFromZipcode" class="col-form-label">郵便番号</label>`;
  html += `    </div>`;
  html += `    <div class="col-auto">`;
  html += `        <input type="text" class="form-control" id="shippingFromZipcode" name="shippingFromZipcode" required value="${
    e.parameter.shippingFromZipcode ? e.parameter.shippingFromZipcode : ""
  }" maxlength=7 pattern="[0-9]{7}" title="7桁の数字のみを入力してください。"readonly>`;
  html += `    </div>`;
  html += `</div>`;
  html += `<div>
                <label for="shippingFromAddress" class="text-left form-label">住所</label>`;
  html += `<input type="text" class="form-control" id="shippingFromAddress" name="shippingFromAddress" required value="${
    e.parameter.shippingFromAddress ? e.parameter.shippingFromAddress : ""
  }" readonly>`;
  html += `</div>`;
  html += `<div class="mt-0 mb-2 row g-3 align-items-center">`;
  html += `    <div class="col-auto">`;
  html += `        <label for="shippingFromTel" class="col-form-label">電話番号</label>`;
  html += `    </div>`;
  html += `    <div class="col-auto">`;
  html += `        <input type="text" class="form-control" id="shippingFromTel" name="shippingFromTel" required value="${
    e.parameter.shippingFromTel ? e.parameter.shippingFromTel : ""
  }" maxlength=11 pattern="^0[0-9]{9,10}$" title="10~11桁の数字のみを入力してください。"readonly>`;
  html += `    </div>`;
  html += `</div>`;
  html += `<div style="background-color: darkcyan; color: white;">【受注基本情報】</div>`;
  html += `<div class="mt-0 mb-2 row g-3 align-items-center">`;
  html += `    <div class="col-auto">`;
  html += `        <label for="shippingDate" class="col-form-label">発送日</label>`;
  html += `    </div>`;
  html += `    <div class="col-auto">`;
  html += `       <input type="date" class="form-control" id="shippingDate" name="shippingDate" required value="${shippingDate}" readonly>`;
  html += `    </div>`;
  html += `    <div class="col-auto">`;
  html += `        <label for="deliveryDate" class="col-form-label">納品日</label>`;
  html += `    </div>`;
  html += `    <div class="col-auto">`;
  html += `       <input type="date" class="form-control" id="deliveryDate" name="deliveryDate" required value="${deliveryDate}" readonly>`;
  html += `    </div>`;
  html += `</div>`;
  html += `<div>`;
  html += `<label for="receiptWay" class="text-left form-label">受付方法</label>`;
  html += `<select class="form-select" id="receiptWay" name="receiptWay" required >`;
  for (const receipt of receipts) {
    const receiptWay = receipt["受付方法"];
    if (receiptWay == e.parameter.receiptWay) {
      html += `<option value="${receiptWay}" selected>${receiptWay}</option>`;
    } else {
      html += `<option value="${receiptWay}" disabled>${receiptWay}</option>`;
    }
  }
  html += `</select>`;
  html += `</div>`;
  html += `<div>`;
  html += `<label for="recipient" class="text-left form-label">受付者</label>`;
  html += `<select class="form-select" id="recipient" name="recipient" required >`;
  for (const recipient of recipients) {
    const recip = recipient["名前"];
    if (recip == e.parameter.recipient) {
      html += `<option value="${recip}" selected>${recip}</option>`;
    } else {
      html += `<option value="${recip}" disabled>${recip}</option>`;
    }
  }
  html += `</select>`;
  html += `</div>`;
  html += `<div class="mt-0 mb-2 row g-3 align-items-center">`;
  html += `    <div class="col-auto">`;
  html += `        <label for="deliveryMethod" class="col-form-label">納品方法</label>`;
  html += `    </div>`;
  html += `    <div class="col-auto">`;
  html += `       <select class="form-select" id="deliveryMethod" name="deliveryMethod" required >`;
  for (const deliveryMethod of deliveryMethods) {
    const deliMethod = deliveryMethod["納品方法"];
    if (deliMethod == e.parameter.deliveryMethod) {
      html += `<option value="${deliMethod}" selected>${deliMethod}</option>`;
    } else {
      html += `<option value="${deliMethod}" disabled>${deliMethod}</option>`;
    }
  }
  html += `    </select>`;
  html += `    </div>`;
  html += `    <div class="col-auto">`;
  html += `        <label for="deliveryTime" class="col-form-label">配達時間帯</label>`;
  html += `    </div>`;
  html += `    <div class="col-auto">`;
  html += `       <select class="form-select" id="deliveryTime" name="deliveryTime" >`;
  html += `         <option value=""></option>`;
  for (const deliveryTime of deliveryTimes) {
    const deliTime = deliveryTime["時間指定"];
    const deliTimeVal = deliveryTime["時間指定値"];
    const deliveryMethod = deliveryTime["納品方法"];
    if (deliTimeVal == e.parameter.deliveryTime) {
      html += `<option value="${deliTimeVal}" data-val="${deliveryMethod}" selected>${deliTime}</option>`;
    } else {
      html += `<option value="${deliTimeVal}" data-val="${deliveryMethod}" disabled>${deliTime}</option>`;
    }
  }
  html += `    </select>`;
  html += `    </div>`;
  html += `</div>`;
  html += `<div class="mt-0 mb-2 row g-3 align-items-center">`;
  html += `   <div class="col-auto">`;
  html += `     <div class="form-check">`;
  if (e.parameters.checklist && e.parameters.checklist.includes("納品書")) {
    html += `       <input class="form-check-input" type="checkbox" value="納品書" name="checklist" checked onclick="return false;">納品書`;
  } else {
    html += `       <input class="form-check-input" type="checkbox" value="納品書" name="checklist" onclick="return false;">納品書`;
  }
  html += `     </div>`;
  html += `   </div>`;
  html += `   <div class="col-auto">`;
  html += `     <div class="form-check">`;
  if (e.parameters.checklist && e.parameters.checklist.includes("請求書")) {
    html += `       <input class="form-check-input" type="checkbox" value="請求書" name="checklist" checked onclick="return false;">請求書`;
  } else {
    html += `       <input class="form-check-input" type="checkbox" value="請求書" name="checklist" onclick="return false;">請求書`;
  }
  html += `     </div>`;
  html += `   </div>`;
  html += `   <div class="col-auto">`;
  html += `     <div class="form-check">`;
  if (e.parameters.checklist && e.parameters.checklist.includes("領収書")) {
    html += `       <input class="form-check-input" type="checkbox" value="領収書" name="checklist" checked onclick="return false;">領収書`;
  } else {
    html += `       <input class="form-check-input" type="checkbox" value="領収書" name="checklist" onclick="return false;">領収書`;
  }
  html += `     </div>`;
  html += `   </div>`;
  html += `   <div class="col-auto">`;
  html += `     <div class="form-check">`;
  if (e.parameters.checklist && e.parameters.checklist.includes("パンフ")) {
    html += `       <input class="form-check-input" type="checkbox" value="パンフ" name="checklist" checked onclick="return false;">パンフ`;
  } else {
    html += `       <input class="form-check-input" type="checkbox" value="パンフ" name="checklist" onclick="return false;">パンフ`;
  }
  html += `     </div>`;
  html += `   </div>`;
  html += `   <div class="col-auto">`;
  html += `     <div class="form-check">`;
  if (e.parameters.checklist && e.parameters.checklist.includes("レシピ")) {
    html += `       <input class="form-check-input" type="checkbox" value="レシピ" name="checklist" checked onclick="return false;">レシピ`;
  } else {
    html += `       <input class="form-check-input" type="checkbox" value="レシピ" name="checklist" onclick="return false;">レシピ`;
  }
  html += `     </div>`;
  html += `   </div>`;
  html += `</div>`;
  html += `<div>`;
  html += `   <label for="otherAttach" class="col-form-label">その他添付</label>`;
  html += `   <input type="text" class="form-control" id="otherAttach" name="otherAttach" value="${
    e.parameter.otherAttach ? e.parameter.otherAttach : ""
  }" readonly >`;
  html += `</div>`;
  html += `
    <p class="m-3 fw-bold">以下の内容で受注していいですか？</p>

    <table class="table">
      <thead>
        <tr>
          <th scope="col" class="text-start">商品分類</th>
          <th scope="col" class="text-end">商品名</th>
          <th scope="col" class="text-end">価格</th>
          <th scope="col" class="text-end">個数</th>
          <th scope="col" class="text-end">金額</th>
        </tr>
      </thead>
      <tbody>
  `;
  let total = 0;
  var rowNum = 0;
  for (let i = 0; i < 10; i++) {
    rowNum++;
    var bunrui = "bunrui" + rowNum;
    var product = "product" + rowNum;
    var price = "price" + rowNum;
    var quantity = "quantity" + rowNum;
    const bunruiVal = e.parameter[bunrui];
    const productVal = e.parameter[product];
    const count = Number(e.parameter[quantity]);
    const unitPrice = Number(e.parameter[price]);
    if (count > 0) {
      const addPrice = unitPrice * count;
      total += addPrice;
      html += `<tr>`;
      html += `<td class="text-start">`;
      html += `<div class="d-flex justify-content-start">`;
      html += `<input type="text" style="max-width: 100px; min-width: 60px;" class="form-control text-start" id="${bunrui}" name="${bunrui}"  value="${bunruiVal}" readonly>`;
      html += `</div>`;
      html += `</td>`;
      html += `<td class="text-end">`;
      html += `<div class="d-flex justify-content-end">`;
      html += `<input type="text" style="max-width: 100px; min-width: 60px;" class="form-control text-end" id="${product}" name="${product}"  value="${productVal}" readonly>`;
      html += `</div>`;
      html += `</td>`;
      html += `<td class="text-end">`;
      html += `<div class="d-flex justify-content-end">`;
      html += `<input type="number" style="max-width: 100px; min-width: 60px;" class="form-control text-end" id="${price}" name="${price}"  value="${unitPrice}" readonly>`;
      html += `</div>`;
      html += `</td>`;
      html += `<td class="text-end">`;
      html += `<div class="d-flex justify-content-end">`;
      html += `<input type="number" style="max-width: 100px; min-width: 60px;" class="form-control text-end" id="${quantity}" name="${quantity}"  value="${count}" readonly>`;
      html += `</div>`;
      html += `</td>`;
      html += `<td class="text-end">¥${addPrice.toLocaleString()}</td>`;
      html += `</tr>`;
    }
  }

  html += `<tr>`;
  html += `<td class="text-end fs-2" colspan="3">合計:</td>`;
  html += `<td class="text-end fs-2" colspan="2">¥${total.toLocaleString()}</td>`;
  html += `</tr>`;
  html += `</tbody>`;
  html += `</table>`;
  html += `<div style="background-color: blue; color: white;">【発送情報】</div>`;
  html += `<div>
                <label for="sendProduct" class="text-left form-label">品名</label>`;
  html += `       <input type="text" class="form-control" id="sendProduct" name="sendProduct" value="${
    e.parameter.sendProduct ? e.parameter.sendProduct : ""
  }" readonly >`;
  html += `</div>`;
  html += `<div class="mt-0 mb-2 row g-3 align-items-center">`;
  html += `    <div class="col-auto">`;
  html += `        <label for="invoiceType" class="col-form-label">送り状種別</label>`;
  html += `    </div>`;
  html += `    <div class="col-auto">`;
  html += `<select class="form-select" id="invoiceType" name="invoiceType" >`;
  html += `<option value=""></option>`;
  for (const invoiceType of invoiceTypes) {
    if (e.parameter.invoiceType == invoiceType["種別値"]) {
      html += `<option value="${invoiceType["種別値"]}" data-val="${invoiceType["納品方法"]}" selected>${invoiceType["種別"]}</option>`;
    } else {
      html += `<option value="${invoiceType["種別値"]}" data-val="${invoiceType["納品方法"]}" disabled>${invoiceType["種別"]}</option>`;
    }
  }
  html += `</select>`;
  html += `    </div>`;
  html += `    <div class="col-auto">`;
  html += `        <label for="coolCls" class="col-form-label">クール区分</label>`;
  html += `    </div>`;
  html += `    <div class="col-auto">`;
  html += `<select class="form-select" id="coolCls" name="coolCls" >`;
  html += `<option value=""></option>`;
  for (const coolCls of coolClss) {
    if (e.parameter.coolCls == coolCls["種別値"]) {
      html += `<option value="${coolCls["種別値"]}" data-val="${coolCls["納品方法"]}" selected>${coolCls["種別"]}</option>`;
    } else {
      html += `<option value="${coolCls["種別値"]}" data-val="${coolCls["納品方法"]}" disabled>${coolCls["種別"]}</option>`;
    }
  }
  html += `</select>`;
  html += `    </div>`;
  html += `</div>`;
  html += `<div class="mt-0 mb-2 row g-3 align-items-center">`;
  html += `    <div class="col-auto">`;
  html += `        <label for="cargo1" class="col-form-label">荷扱い１</label>`;
  html += `    </div>`;
  html += `    <div class="col-auto">`;
  html += `<select class="form-select" id="cargo1" name="cargo1" >`;
  html += `<option value=""></option>`;
  for (const cargo of cargos) {
    if (e.parameter.cargo1 == cargo["種別値"]) {
      html += `<option value="${cargo["種別値"]}" data-val="${cargo["納品方法"]}" selected>${cargo["種別"]}</option>`;
    } else {
      html += `<option value="${cargo["種別値"]}" data-val="${cargo["納品方法"]}" disabled>${cargo["種別"]}</option>`;
    }
  }
  html += `</select>`;
  html += `    </div>`;
  html += `    <div class="col-auto">`;
  html += `        <label for="cargo2" class="col-form-label">荷扱い２</label>`;
  html += `    </div>`;
  html += `    <div class="col-auto">`;
  html += `<select class="form-select" id="cargo2" name="cargo2" >`;
  html += `<option value=""></option>`;
  for (const cargo of cargos) {
    if (e.parameter.cargo2 == cargo["種別値"]) {
      html += `<option value="${cargo["種別値"]}" data-val="${cargo["納品方法"]}" selected>${cargo["種別"]}</option>`;
    } else {
      html += `<option value="${cargo["種別値"]}" data-val="${cargo["納品方法"]}" disabled>${cargo["種別"]}</option>`;
    }
  }
  html += `</select>`;
  html += `    </div>`;
  html += `</div>`;
  html += `<div class="mt-0 mb-2 row g-3 align-items-center">`;
  html += `    <div class="col-auto">`;
  html += `        <label for="cashOnDelivery" class="col-form-label">代引総額</label>`;
  html += `    </div>`;
  html += `    <div class="col-auto">`;
  html += `       <input type="number" class="form-control" id="cashOnDelivery" name="cashOnDelivery" min='0'  value="${
    e.parameter.cashOnDelivery ? e.parameter.cashOnDelivery : ""
  }" readonly >`;
  html += `    </div>`;
  html += `    <div class="col-auto">`;
  html += `        <label for="cashOnDeliTax" class="col-form-label">代引内税</label>`;
  html += `    </div>`;
  html += `    <div class="col-auto">`;
  html += `       <input type="number" class="form-control" id="cashOnDeliTax" name="cashOnDeliTax" min='0'  value="${
    e.parameter.cashOnDeliTax ? e.parameter.cashOnDeliTax : ""
  }" readonly >`;
  html += `    </div>`;
  html += `</div>`;
  html += `<div class="mt-0 mb-2 row g-3 align-items-center">`;
  html += `    <div class="col-auto">`;
  html += `        <label for="copiePrint" class="col-form-label">発行枚数</label>`;
  html += `    </div>`;
  html += `    <div class="col-auto">`;
  html += `       <input type="number" class="form-control" id="copiePrint" name="copiePrint" min='0'  value="${
    e.parameter.copiePrint ? e.parameter.copiePrint : ""
  }" readonly >`;
  html += `    </div>`;
  html += `</div>`;
  html += `<div class="mb-3">`;
  html += `<label for="csvmemo" class="text-left form-label">送り状　備考欄</label>`;
  html += `<textarea class="form-control" id="csvmemo" name="csvmemo" rows="2" cols="30" maxlength="22" readonly>${
    e.parameter.csvmemo ? e.parameter.csvmemo : ""
  }</textarea>`;
  html += `</div>`;
  html += `<div class="mb-3">`;
  html += `<label for="deliveryMemo" class="text-left form-label">納品書　備考欄</label>`;
  html += `<textarea class="form-control" id="deliveryMemo" name="deliveryMemo" rows="3" cols="30" maxlength="90" readonly>${
    e.parameter.deliveryMemo ? e.parameter.deliveryMemo : ""
  }</textarea>`;
  html += `<div class="mb-3">`;
  html += `<label for="memo" class="text-left form-label">メモ</label>`;
  html += `<textarea class="form-control" id="memo" name="memo" rows="3" cols="30" readonly>${e.parameter.memo}</textarea>`;
  html += `</div>`;
  html += `</div>`;

  return html;
}
// 受注IDの生成
function generateId(length = 8) {
  const [alphabets, numbers] = ["abcdefghijklmnopqrstuvwxyz", "0123456789"];
  const string = alphabets + numbers;
  let id = alphabets.charAt(Math.floor(Math.random() * alphabets.length));
  for (let i = 0; i < length - 1; i++) {
    id += string.charAt(Math.floor(Math.random() * string.length));
  }
  return id;
}
// 受注登録
function createOrder(e) {
  // 納品ID
  const deliveryId = generateId();
  // 受注テーブルに複数レコードを追加する
  const records = [];
  const createRecords = [];
  var rowNum = 0;
  var dateNow = Utilities.formatDate(new Date(), "JST", "yyyy/MM/dd");

  for (let i = 0; i < 10; i++) {
    rowNum++;
    var bunrui = "bunrui" + rowNum;
    var product = "product" + rowNum;
    var price = "price" + rowNum;
    var quantity = "quantity" + rowNum;
    const bunruiVal = e.parameter[bunrui];
    const productVal = e.parameter[product];
    const count = Number(e.parameter[quantity]);
    const unitPrice = Number(e.parameter[price]);
    if (count > 0) {
      var record = [];
      record["受注ID"] = deliveryId;
      record["受注日"] = dateNow;
      record["顧客名"] = e.parameter.customerName;
      record["顧客郵便番号"] = e.parameter.customerZipcode;
      record["顧客住所"] = e.parameter.customerAddress;
      record["顧客電話番号"] = e.parameter.customerTel;
      record["発送先名"] = e.parameter.shippingToName;
      record["発送先郵便番号"] = e.parameter.shippingToZipcode;
      record["発送先住所"] = e.parameter.shippingToAddress;
      record["発送先電話番号"] = e.parameter.shippingToTel;
      record["発送元名"] = e.parameter.shippingFromName;
      record["発送元郵便番号"] = e.parameter.shippingFromZipcode;
      record["発送元住所"] = e.parameter.shippingFromAddress;
      record["発送元電話番号"] = e.parameter.shippingFromTel;
      record["発送日"] = e.parameter.shippingDate;
      record["納品日"] = e.parameter.deliveryDate;
      record["受付方法"] = e.parameter.receiptWay;
      record["受付者"] = e.parameter.recipient;
      record["納品方法"] = e.parameter.deliveryMethod;
      record["配達時間帯"] = e.parameter.deliveryTime
        ? e.parameter.deliveryTime.split(":")[1]
        : "";
      record["納品書"] = e.parameters.checklist
        ? e.parameters.checklist.includes("納品書")
          ? "○"
          : ""
        : "";
      record["請求書"] = e.parameters.checklist
        ? e.parameters.checklist.includes("請求書")
          ? "○"
          : ""
        : "";
      record["領収書"] = e.parameters.checklist
        ? e.parameters.checklist.includes("領収書")
          ? "○"
          : ""
        : "";
      record["パンフ"] = e.parameters.checklist
        ? e.parameters.checklist.includes("パンフ")
          ? "○"
          : ""
        : "";
      record["レシピ"] = e.parameters.checklist
        ? e.parameters.checklist.includes("レシピ")
          ? "○"
          : ""
        : "";
      record["その他添付"] = e.parameter.otherAttach;
      record["品名"] = e.parameter.sendProduct;
      record["送り状種別"] = e.parameter.invoiceType
        ? e.parameter.invoiceType.split(":")[1]
        : "";
      record["クール区分"] = e.parameter.coolCls
        ? e.parameter.coolCls.split(":")[1]
        : "";
      record["荷扱い１"] = e.parameter.cargo1
        ? e.parameter.cargo1.split(":")[1]
        : "";
      record["荷扱い２"] = e.parameter.cargo2
        ? e.parameter.cargo2.split(":")[1]
        : "";
      record["代引総額"] = e.parameter.cashOnDelivery;
      record["代引内税"] = e.parameter.cashOnDeliTax;
      record["発行枚数"] = e.parameter.copiePrint;
      record["送り状備考欄"] = e.parameter.csvmemo;
      record["納品書備考欄"] = e.parameter.deliveryMemo;
      record["メモ"] = e.parameter.memo;
      record["商品分類"] = bunruiVal;
      record["商品名"] = productVal;
      record["受注数"] = count;
      record["販売価格"] = unitPrice;
      record["小計"] = count * unitPrice;

      const addRecord = [
        record["受注ID"],
        record["受注日"],
        record["商品分類"],
        record["商品名"],
        record["受注数"],
        record["販売価格"],
        record["顧客名"],
        record["顧客郵便番号"],
        record["顧客住所"],
        record["顧客電話番号"],
        record["発送先名"],
        record["発送先郵便番号"],
        record["発送先住所"],
        record["発送先電話番号"],
        record["発送元名"],
        record["発送元郵便番号"],
        record["発送元住所"],
        record["発送元電話番号"],
        record["発送日"],
        record["納品日"],
        record["受付方法"],
        record["受付者"],
        record["納品方法"],
        record["配達時間帯"],
        record["納品書"],
        record["請求書"],
        record["領収書"],
        record["パンフ"],
        record["レシピ"],
        record["その他添付"],
        record["品名"],
        record["送り状種別"],
        record["クール区分"],
        record["荷扱い１"],
        record["荷扱い２"],
        record["代引総額"],
        record["代引内税"],
        record["発行枚数"],
        record["送り状備考欄"],
        record["納品書備考欄"],
        record["メモ"],
        record["小計"],
      ];
      records.push(addRecord);
      createRecords.push(record);
    }
  }
  addRecords("受注", records);
  if (e.parameter.deliveryMethod == "ヤマト") {
    addRecordYamato("ヤマトCSV", records, e);
  }
  if (e.parameter.deliveryMethod == "佐川") {
    addRecordSagawa("佐川CSV", records, e);
  }
  if (e.parameters.checklist && e.parameters.checklist.includes("納品書")) {
    createFile(createRecords);
  }
  if (e.parameters.checklist && e.parameters.checklist.includes("領収書")) {
    createReceiptFile(createRecords);
  }
}
// ヤマトCSV登録
function addRecordYamato(sheetName, records, e) {
  Logger.log(e);
  Logger.log(records);
  const adds = [];
  var record = [];
  record["発送日"] = Utilities.formatDate(
    new Date(records[0][18]),
    "JST",
    "yyyy/MM/dd"
  );
  record["お客様管理番号"] = "";
  record["送り状種別"] = e.parameter.invoiceType
    ? e.parameter.invoiceType.split(":")[0]
    : "";
  record["クール区分"] = e.parameter.coolCls
    ? e.parameter.coolCls.split(":")[0]
    : "";
  record["伝票番号"] = "";
  record["出荷予定日"] = Utilities.formatDate(
    new Date(records[0][18]),
    "JST",
    "yyyy/MM/dd"
  );
  record["お届け予定（指定）日"] = Utilities.formatDate(
    new Date(records[0][19]),
    "JST",
    "yyyy/MM/dd"
  );
  record["配達時間帯"] = e.parameter.deliveryTime
    ? e.parameter.deliveryTime.split(":")[0]
    : "";
  record["お届け先コード"] = "";
  record["お届け先電話番号"] = records[0][13];
  record["お届け先電話番号枝番"] = "";
  record["お届け先郵便番号"] = records[0][11];
  record["お届け先住所"] = records[0][12];
  record["お届け先住所（アパートマンション名）"] = "";
  record["お届け先会社・部門名１"] = "";
  record["お届け先会社・部門名２"] = "";
  record["お届け先名"] = records[0][10];
  record["お届け先名略称カナ"] = "";
  record["敬称"] = "";
  record["ご依頼主コード"] = "";
  record["ご依頼主電話番号"] = records[0][17];
  record["ご依頼主電話番号枝番"] = "";
  record["ご依頼主郵便番号"] = records[0][15];
  record["ご依頼主住所"] = records[0][16];
  record["ご依頼主住所（アパートマンション名）"] = "";
  record["ご依頼主名"] = records[0][14];
  record["ご依頼主略称カナ"] = "";
  record["品名コード１"] = "";
  record["品名１"] = records[0][30];
  record["品名コード２"] = "";
  record["品名２"] = "";
  record["荷扱い１"] = e.parameter.cargo1
    ? e.parameter.cargo1.split(":")[0]
    : "";
  record["荷扱い２"] = e.parameter.cargo2
    ? e.parameter.cargo2.split(":")[0]
    : "";
  record["記事"] = records[0][38];
  record["コレクト代金引換額（税込）"] = records[0][35];
  record["コレクト内消費税額等"] = records[0][36];
  record["営業所止置き"] = "";
  record["営業所コード"] = "";
  record["発行枚数"] = records[0][37];
  record["個数口枠の印字"] = "";
  record["ご請求先顧客コード"] = "9999999999";
  record["ご請求先分類コード"] = "";
  record["運賃管理番号"] = "01";
  record["クロネコwebコレクトデータ登録"] = "";
  record["クロネコwebコレクト加盟店番号"] = "";
  record["クロネコwebコレクト申込受付番号１"] = "";
  record["クロネコwebコレクト申込受付番号２"] = "";
  record["クロネコwebコレクト申込受付番号３"] = "";
  record["お届け予定ｅメール利用区分"] = "";
  record["お届け予定ｅメールe-mailアドレス"] = "";
  record["入力機種"] = "";
  record["お届け予定eメールメッセージ"] = "";
  const addList = [
    record["発送日"],
    record["お客様管理番号"],
    record["送り状種別"],
    record["クール区分"],
    record["伝票番号"],
    record["出荷予定日"],
    record["お届け予定（指定）日"],
    record["配達時間帯"],
    record["お届け先コード"],
    record["お届け先電話番号"],
    record["お届け先電話番号枝番"],
    record["お届け先郵便番号"],
    record["お届け先住所"],
    record["お届け先住所（アパートマンション名）"],
    record["お届け先会社・部門名１"],
    record["お届け先会社・部門名２"],
    record["お届け先名"],
    record["お届け先名略称カナ"],
    record["敬称"],
    record["ご依頼主コード"],
    record["ご依頼主電話番号"],
    record["ご依頼主電話番号枝番"],
    record["ご依頼主郵便番号"],
    record["ご依頼主住所"],
    record["ご依頼主住所（アパートマンション名）"],
    record["ご依頼主名"],
    record["ご依頼主略称カナ"],
    record["品名コード１"],
    record["品名１"],
    record["品名コード２"],
    record["品名２"],
    record["荷扱い１"],
    record["荷扱い２"],
    record["記事"],
    record["コレクト代金引換額（税込）"],
    record["コレクト内消費税額等"],
    record["営業所止置き"],
    record["営業所コード"],
    record["発行枚数"],
    record["個数口枠の印字"],
    record["ご請求先顧客コード"],
    record["ご請求先分類コード"],
    record["運賃管理番号"],
    record["クロネコwebコレクトデータ登録"],
    record["クロネコwebコレクト加盟店番号"],
    record["クロネコwebコレクト申込受付番号１"],
    record["クロネコwebコレクト申込受付番号２"],
    record["クロネコwebコレクト申込受付番号３"],
    record["お届け予定ｅメール利用区分"],
    record["お届け予定ｅメールe-mailアドレス"],
    record["入力機種"],
    record["お届け予定eメールメッセージ"],
  ];
  adds.push(addList);
  addRecords(sheetName, adds);
}
// 佐川CSV登録
function addRecordSagawa(sheetName, records, e) {
  Logger.log(e);
  Logger.log(records);
  const adds = [];
  var record = [];
  record["発送日"] = Utilities.formatDate(
    new Date(records[0][18]),
    "JST",
    "yyyy/MM/dd"
  );
  record["お届け先コード取得区分"] = "";
  record["お届け先コード"] = "";
  record["お届け先電話番号"] = records[0][13];
  record["お届け先郵便番号"] = records[0][11];
  record["お届け先住所１"] = records[0][12];
  record["お届け先住所２"] = "";
  record["お届け先住所３"] = "";
  record["お届け先名称１"] = records[0][10];
  record["お届け先名称２"] = "";
  record["お客様管理番号"] = "";
  record["お客様コード"] = "";
  record["部署ご担当者コード取得区分"] = "";
  record["部署ご担当者コード"] = "";
  record["部署ご担当者名称"] = "";
  record["荷送人電話番号"] = "";
  record["ご依頼主コード取得区分"] = "";
  record["ご依頼主コード"] = "";
  record["ご依頼主電話番号"] = records[0][17];
  record["ご依頼主郵便番号"] = records[0][15];
  record["ご依頼主住所１"] = records[0][16];
  record["ご依頼主住所２"] = "";
  record["ご依頼主名称１"] = records[0][14];
  record["ご依頼主名称２"] = "";
  record["荷姿"] = "";
  if (records[0][30].length > 16) {
    record["品名１"] = records[0][30].substring(0, 16);
  } else {
    record["品名１"] = records[0][30];
  }
  record["品名２"] = "";
  record["品名３"] = "";
  record["品名４"] = "";
  record["品名５"] = "";
  record["荷札荷姿"] = "";
  record["荷札品名１"] = "";
  record["荷札品名２"] = "";
  record["荷札品名３"] = "";
  record["荷札品名４"] = "";
  record["荷札品名５"] = "";
  record["荷札品名６"] = "";
  record["荷札品名７"] = "";
  record["荷札品名８"] = "";
  record["荷札品名９"] = "";
  record["荷札品名１０"] = "";
  record["荷札品名１１"] = "";
  record["出荷個数"] = records[0][37];
  record["スピード指定"] = "000";
  record["クール便指定"] = e.parameter.coolCls
    ? e.parameter.coolCls.split(":")[0]
    : "";
  record["配達日"] = Utilities.formatDate(
    new Date(records[0][19]),
    "JST",
    "yyyyMMdd"
  );
  record["配達指定時間帯"] = e.parameter.deliveryTime
    ? e.parameter.deliveryTime.split(":")[0]
    : "";
  record["配達指定時間（時分）"] = "";
  record["代引金額"] = records[0][35];
  record["消費税"] = records[0][36];
  record["決済種別"] = "";
  record["保険金額"] = "";
  record["指定シール１"] = e.parameter.cargo1
    ? e.parameter.cargo1.split(":")[0]
    : "";
  record["指定シール２"] = e.parameter.cargo2
    ? e.parameter.cargo2.split(":")[0]
    : "";
  record["指定シール３"] = "";
  record["営業所受取"] = "";
  record["SRC区分"] = "";
  record["営業所受取営業所コード"] = "";
  record["元着区分"] = e.parameter.invoiceType
    ? e.parameter.invoiceType.split(":")[0]
    : "";
  record["メールアドレス"] = "";
  record["ご不在時連絡先"] = records[0][13];
  record["出荷予定日"] = "";
  record["セット数"] = "";
  record["お問い合せ送り状No."] = "";
  record["出荷場印字区分"] = "";
  record["集約解除指定"] = "";
  record["編集０１"] = "";
  record["編集０２"] = "";
  record["編集０３"] = "";
  record["編集０４"] = "";
  record["編集０５"] = "";
  record["編集０６"] = "";
  record["編集０７"] = "";
  record["編集０８"] = "";
  record["編集０９"] = "";
  record["編集１０"] = "";
  const addList = [
    record["発送日"],
    record["お届け先コード取得区分"],
    record["お届け先コード"],
    record["お届け先電話番号"],
    record["お届け先郵便番号"],
    record["お届け先住所１"],
    record["お届け先住所２"],
    record["お届け先住所３"],
    record["お届け先名称１"],
    record["お届け先名称２"],
    record["お客様管理番号"],
    record["お客様コード"],
    record["部署ご担当者コード取得区分"],
    record["部署ご担当者コード"],
    record["部署ご担当者名称"],
    record["荷送人電話番号"],
    record["ご依頼主コード取得区分"],
    record["ご依頼主コード"],
    record["ご依頼主電話番号"],
    record["ご依頼主郵便番号"],
    record["ご依頼主住所１"],
    record["ご依頼主住所２"],
    record["ご依頼主名称１"],
    record["ご依頼主名称２"],
    record["荷姿"],
    record["品名１"],
    record["品名２"],
    record["品名３"],
    record["品名４"],
    record["品名５"],
    record["荷札荷姿"],
    record["荷札品名１"],
    record["荷札品名２"],
    record["荷札品名３"],
    record["荷札品名４"],
    record["荷札品名５"],
    record["荷札品名６"],
    record["荷札品名７"],
    record["荷札品名８"],
    record["荷札品名９"],
    record["荷札品名１０"],
    record["荷札品名１１"],
    record["出荷個数"],
    record["スピード指定"],
    record["クール便指定"],
    record["配達日"],
    record["配達指定時間帯"],
    record["配達指定時間（時分）"],
    record["代引金額"],
    record["消費税"],
    record["決済種別"],
    record["保険金額"],
    record["指定シール１"],
    record["指定シール２"],
    record["指定シール３"],
    record["営業所受取"],
    record["SRC区分"],
    record["営業所受取営業所コード"],
    record["元着区分"],
    record["メールアドレス"],
    record["ご不在時連絡先"],
    record["出荷予定日"],
    record["セット数"],
    record["お問い合せ送り状No."],
    record["出荷場印字区分"],
    record["集約解除指定"],
    record["編集０１"],
    record["編集０２"],
    record["編集０３"],
    record["編集０４"],
    record["編集０５"],
    record["編集０６"],
    record["編集０７"],
    record["編集０８"],
    record["編集０９"],
    record["編集１０"],
  ];
  adds.push(addList);
  addRecords(sheetName, adds);
}
// レコード登録
function addRecords(sheetName, records) {
  Logger.log(sheetName);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  const lastRow = sheet.getLastRow();
  sheet
    .getRange(lastRow + 1, 1, records.length, records[0].length)
    .setNumberFormat("@")
    .setValues(records)
    .setBorder(true, true, true, true, true, true);
}
// 在庫更新
function updateZaiko(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("商品");
  var rowNum = 0;
  for (let i = 0; i < 10; i++) {
    rowNum++;
    var product = "product" + rowNum;
    var quantity = "quantity" + rowNum;
    const productVal = e.parameter[product];
    const count = Number(e.parameter[quantity]);
    var zaiko = 0;
    if (count > 0) {
      const targetRow = sheet
        .getRange("B:B")
        .createTextFinder(productVal)
        .matchEntireCell(true)
        .findNext()
        .getRow();
      const targetCol = sheet
        .getRange("1:1")
        .createTextFinder("在庫数")
        .matchEntireCell(true)
        .findNext()
        .getColumn();
      zaiko = Number(sheet.getRange(targetRow, targetCol).getValue());
      sheet.getRange(targetRow, targetCol).setValue(zaiko - count);
    }
  }
}
// 納品書ファイル生成
function createFile(records) {
  // PDF変換する元ファイルを作成する
  let wFileRtn = createGDoc(records);
  // PDF変換
  createDeliveredPdf(wFileRtn[0], wFileRtn[1], wFileRtn[2]);
  // PDF変換したあとは元ファイルを削除する
  DriveApp.getFileById(wFileRtn[0]).setTrashed(true);
}

// 納品書のドキュメントの中身を置換
function createGDoc(rowVal) {
  // 顧客情報シートを取得する
  const customerItems = getAllRecords("顧客情報");
  const shippingToItems = getAllRecords("発送先情報");
  const productItems = getAllRecords("商品");
  var customerItem = [];
  Logger.log(rowVal);
  Logger.log(rowVal[0]);
  Logger.log(rowVal[0]["顧客名"]);
  const shippingName = rowVal[0]["顧客名"].split("　");
  Logger.log(shippingName);
  // データ走査
  customerItems.forEach(function (wVal) {
    if (shippingName.length > 1) {
      // 会社名と同じ
      if (
        shippingName[1] == wVal["氏名"] &&
        shippingName[0] == wVal["会社名"]
      ) {
        customerItem = wVal;
      }
    } else {
      // 会社名と同じ
      if (shippingName[0] == wVal["氏名"]) {
        customerItem = wVal;
      }
      if (shippingName[0] == wVal["会社名"]) {
        customerItem = wVal;
      }
    }
  });
  if (customerItem.length == 0) {
    customerItem["会社名"] = rowVal[0]["顧客名"].split("　")[0];
    customerItem["住所１"] = rowVal[0]["顧客住所"];
    customerItem["郵便番号"] = rowVal[0]["顧客郵便番号"];
    customerItem["住所２"] = "";
  }
  // テンプレートファイルをコピーする
  const wCopyFile = DELIVERED_TEMPLATE.makeCopy(),
    wCopyFileId = wCopyFile.getId(),
    wCopyDoc = DocumentApp.openById(wCopyFileId); // コピーしたファイルをGoogleドキュメントとして開く
  let wCopyDocBody = wCopyDoc.getBody(); // Googleドキュメント内の本文を取得する
  var post = String(customerItem["郵便番号"]);
  post = post.substring(0, 3).concat("-").concat(post.substring(3, 7));

  // 注文書ファイル内の可変文字部（として用意していた箇所）を変更する
  wCopyDocBody = wCopyDocBody.replaceText(
    "{{company_name}}",
    customerItem["会社名"] ? customerItem["会社名"] : customerItem["氏名"]
  );
  wCopyDocBody = wCopyDocBody.replaceText("{{post}}", post);
  wCopyDocBody = wCopyDocBody.replaceText(
    "{{address1}}",
    customerItem["住所１"]
  );
  wCopyDocBody = wCopyDocBody.replaceText(
    "{{address2}}",
    customerItem["住所２"]
  );
  wCopyDocBody = wCopyDocBody.replaceText(
    "{{delivery_num}}",
    rowVal[0]["受注ID"]
  );
  wCopyDocBody = wCopyDocBody.replaceText(
    "{{delivery_date}}",
    Utilities.formatDate(new Date(rowVal[0]["納品日"]), "JST", "yyyy年MM月dd日")
  );
  wCopyDocBody = wCopyDocBody.replaceText(
    "{{deliveryMemo}}",
    rowVal[0]["納品書備考欄"]
  );
  let totals = 0;
  let tentax = 0;
  let eigtax = 0;
  let tentax_t = 0;
  let eigtax_t = 0;
  let tax = 0;
  let amount = 0;

  for (let i = 0; i < 10; i++) {
    var total = 0;
    if (i < rowVal.length) {
      var productData;
      // 商品分類
      var changeText = "{{bunrui" + (i + 1) + "}}";
      wCopyDocBody = wCopyDocBody.replaceText(
        changeText,
        rowVal[i]["商品分類"]
      );
      // 商品名
      changeText = "{{product" + (i + 1) + "}}";
      wCopyDocBody = wCopyDocBody.replaceText(changeText, rowVal[i]["商品名"]);
      // データ走査
      productItems.forEach(function (wVal) {
        // 商品名と同じ
        if (wVal["商品名"] == rowVal[i]["商品名"]) {
          productData = wVal;
        }
      });
      // 価格（P)
      changeText = "{{price" + (i + 1) + "}}";
      wCopyDocBody = wCopyDocBody.replaceText(
        changeText,
        "￥ " + rowVal[i]["販売価格"]
      );
      // 数量
      changeText = "{{c" + (i + 1) + "}}";
      wCopyDocBody = wCopyDocBody.replaceText(changeText, rowVal[i]["受注数"]);
      // 金額
      changeText = "{{amount" + (i + 1) + "}}";
      total = rowVal[i]["受注数"] * rowVal[i]["販売価格"];
      if (Number(productData["税率"]) > 8) {
        var taxValTotal = Math.round(Number(total * 1.1));
        var taxVal = taxValTotal - Number(total);
        tentax += taxVal;
        tentax_t += total;
        tax += taxVal;
        amount += total;
        totals += taxValTotal;
      } else {
        var taxValTotal = Math.round(Number(total * 1.08));
        var taxVal = taxValTotal - Number(total);
        eigtax += taxVal;
        eigtax_t += total;
        tax += taxVal;
        amount += total;
        totals += taxValTotal;
      }
      wCopyDocBody = wCopyDocBody.replaceText(
        changeText,
        "￥ " + total.toLocaleString()
      );
    } else {
      var changeText = "{{bunrui" + (i + 1) + "}}";
      wCopyDocBody = wCopyDocBody.replaceText(changeText, "");
      changeText = "{{product" + (i + 1) + "}}";
      wCopyDocBody = wCopyDocBody.replaceText(changeText, "");
      changeText = "{{price" + (i + 1) + "}}";
      wCopyDocBody = wCopyDocBody.replaceText(changeText, "");
      changeText = "{{c" + (i + 1) + "}}";
      wCopyDocBody = wCopyDocBody.replaceText(changeText, "");
      changeText = "{{amount" + (i + 1) + "}}";
      wCopyDocBody = wCopyDocBody.replaceText(changeText, "");
    }
  }
  wCopyDocBody = wCopyDocBody.replaceText(
    "{{amount}}",
    amount.toLocaleString()
  );
  wCopyDocBody = wCopyDocBody.replaceText("{{tax}}", tax.toLocaleString());
  wCopyDocBody = wCopyDocBody.replaceText(
    "{{10tax_t}}",
    tentax_t.toLocaleString()
  );
  wCopyDocBody = wCopyDocBody.replaceText("{{10tax}}", tentax.toLocaleString());
  wCopyDocBody = wCopyDocBody.replaceText(
    "{{8tax_t}}",
    eigtax_t.toLocaleString()
  );
  wCopyDocBody = wCopyDocBody.replaceText("{{8tax}}", eigtax.toLocaleString());
  wCopyDocBody = wCopyDocBody.replaceText(
    "{{amount_delivered}}",
    totals.toLocaleString()
  );
  wCopyDocBody = wCopyDocBody.replaceText(
    "{{total}}",
    "￥ " + totals.toLocaleString()
  );
  wCopyDoc.saveAndClose();

  // ファイル名を変更する
  let fileName =
    Utilities.formatDate(new Date(), "JST", "yyyyMMdd") +
    "_" +
    customerItem["会社名"] +
    " 御中";
  wCopyFile.setName(fileName);
  // コピーしたファイルIDとファイル名を返却する（あとでこのIDをもとにPDFに変換するため）
  return [wCopyFileId, fileName, customerItem["会社名"]];
}
// PDF生成
function createDeliveredPdf(docId, fileName, targetFolderName) {
  // PDF変換するためのベースURLを作成する
  let wUrl = `https://docs.google.com/document/d/${docId}/export?exportFormat=pdf`;

  // headersにアクセストークンを格納する
  let wOtions = {
    headers: {
      Authorization: `Bearer ${ScriptApp.getOAuthToken()}`,
    },
  };
  // PDFを作成する
  let wBlob = UrlFetchApp.fetch(wUrl, wOtions)
    .getBlob()
    .setName(fileName + ".pdf");
  // 保存先のフォルダが存在するか確認
  var targetFolder = null;
  var folders = DELIVERED_PDF_OUTDIR.getFoldersByName(targetFolderName);

  if (folders.hasNext()) {
    // 既存のフォルダが見つかった場合
    targetFolder = folders.next();
  } else {
    // フォルダが存在しない場合は新規作成
    targetFolder = DELIVERED_PDF_OUTDIR.createFolder(targetFolderName);
  }
  //PDFを指定したフォルダに保存する
  return targetFolder.createFile(wBlob).getId();
}
// PDF生成
function createReceiptPdf(docId, fileName) {
  // PDF変換するためのベースURLを作成する
  let wUrl = `https://docs.google.com/document/d/${docId}/export?exportFormat=pdf`;

  // headersにアクセストークンを格納する
  let wOtions = {
    headers: {
      Authorization: `Bearer ${ScriptApp.getOAuthToken()}`,
    },
  };
  // PDFを作成する
  let wBlob = UrlFetchApp.fetch(wUrl, wOtions)
    .getBlob()
    .setName(fileName + ".pdf");

  //PDFを指定したフォルダに保存する
  return RECEIPT_PDF_OUTDIR.createFile(wBlob).getId();
}
// 領収書のドキュメントを置換
function createReceiptGDoc(rowVal) {
  // 顧客情報シートを取得する
  const customerItems = getAllRecords("顧客情報");
  const shippingToItems = getAllRecords("発送先情報");
  const productItems = getAllRecords("商品");
  var customerItem = [];
  Logger.log(rowVal);
  Logger.log(rowVal[0]);
  Logger.log(rowVal[0]["顧客名"]);
  const shippingName = rowVal[0]["顧客名"].split("　")[0];
  Logger.log(shippingName);
  // データ走査
  shippingToItems.forEach(function (wVal) {
    if (shippingName.length > 1) {
      // 会社名と同じ
      if (
        shippingName[1] == wVal["氏名"] &&
        shippingName[0] == wVal["会社名"]
      ) {
        customerItem = wVal;
      }
    } else {
      // 会社名と同じ
      if (shippingName == wVal["氏名"]) {
        customerItem = wVal;
      }
    }
  });
  if (customerItem.length == 0) {
    customerItem["会社名"] = rowVal[0]["顧客名"].split("　")[0];
    customerItem["住所１"] = rowVal[0]["顧客住所"];
    customerItem["郵便番号"] = rowVal[0]["顧客郵便番号"];
    customerItem["住所２"] = "";
  }
  // テンプレートファイルをコピーする
  const wCopyFile = RECEIPT_TEMPLATE.makeCopy(),
    wCopyFileId = wCopyFile.getId(),
    wCopyDoc = DocumentApp.openById(wCopyFileId); // コピーしたファイルをGoogleドキュメントとして開く
  let wCopyDocBody = wCopyDoc.getBody(); // Googleドキュメント内の本文を取得する
  var post = String(customerItem["郵便番号"]);
  post = post.substring(0, 3).concat("-").concat(post.substring(3, 7));

  // 注文書ファイル内の可変文字部（として用意していた箇所）を変更する
  wCopyDocBody = wCopyDocBody.replaceText(
    "{{company_name}}",
    customerItem["会社名"] ? customerItem["会社名"] : customerItem["氏名"]
  );
  wCopyDocBody = wCopyDocBody.replaceText("{{post}}", post);
  wCopyDocBody = wCopyDocBody.replaceText(
    "{{address1}}",
    customerItem["住所１"]
  );
  wCopyDocBody = wCopyDocBody.replaceText(
    "{{address2}}",
    customerItem["住所２"]
  );
  wCopyDocBody = wCopyDocBody.replaceText(
    "{{delivery_num}}",
    rowVal[0]["受注ID"]
  );
  wCopyDocBody = wCopyDocBody.replaceText(
    "{{delivery_date}}",
    Utilities.formatDate(new Date(rowVal[0]["納品日"]), "JST", "yyyy年MM月dd日")
  );
  let totals = 0;
  let tentax = 0;
  let eigtax = 0;
  let tentax_t = 0;
  let eigtax_t = 0;
  let tax = 0;
  let amount = 0;

  for (let i = 0; i < 10; i++) {
    var total = 0;
    if (i < rowVal.length) {
      var productData;
      // 商品分類
      var changeText = "{{bunrui" + (i + 1) + "}}";
      wCopyDocBody = wCopyDocBody.replaceText(
        changeText,
        rowVal[i]["商品分類"]
      );
      // 商品名
      changeText = "{{product" + (i + 1) + "}}";
      wCopyDocBody = wCopyDocBody.replaceText(changeText, rowVal[i]["商品名"]);
      // データ走査
      productItems.forEach(function (wVal) {
        // 商品名と同じ
        if (wVal["商品名"] == rowVal[i]["商品名"]) {
          productData = wVal;
        }
      });
      // 価格（P)
      changeText = "{{price" + (i + 1) + "}}";
      wCopyDocBody = wCopyDocBody.replaceText(
        changeText,
        "￥ " + rowVal[i]["販売価格"]
      );
      // 数量
      changeText = "{{c" + (i + 1) + "}}";
      wCopyDocBody = wCopyDocBody.replaceText(changeText, rowVal[i]["受注数"]);
      // 金額
      changeText = "{{amount" + (i + 1) + "}}";
      total = rowVal[i]["受注数"] * rowVal[i]["販売価格"];
      if (Number(productData["税率"]) > 8) {
        var taxValTotal = Math.round(Number(total * 1.1));
        var taxVal = taxValTotal - Number(total);
        tentax += taxVal;
        tentax_t += total;
        tax += taxVal;
        amount += total;
        totals += taxValTotal;
      } else {
        var taxValTotal = Math.round(Number(total * 1.08));
        var taxVal = taxValTotal - Number(total);
        eigtax += taxVal;
        eigtax_t += total;
        tax += taxVal;
        amount += total;
        totals += taxValTotal;
      }
      wCopyDocBody = wCopyDocBody.replaceText(
        changeText,
        "￥ " + total.toLocaleString()
      );
    } else {
      var changeText = "{{bunrui" + (i + 1) + "}}";
      wCopyDocBody = wCopyDocBody.replaceText(changeText, "");
      changeText = "{{product" + (i + 1) + "}}";
      wCopyDocBody = wCopyDocBody.replaceText(changeText, "");
      changeText = "{{price" + (i + 1) + "}}";
      wCopyDocBody = wCopyDocBody.replaceText(changeText, "");
      changeText = "{{c" + (i + 1) + "}}";
      wCopyDocBody = wCopyDocBody.replaceText(changeText, "");
      changeText = "{{amount" + (i + 1) + "}}";
      wCopyDocBody = wCopyDocBody.replaceText(changeText, "");
    }
  }
  wCopyDocBody = wCopyDocBody.replaceText(
    "{{amount}}",
    amount.toLocaleString()
  );
  wCopyDocBody = wCopyDocBody.replaceText("{{tax}}", tax.toLocaleString());
  wCopyDocBody = wCopyDocBody.replaceText(
    "{{10tax_t}}",
    tentax_t.toLocaleString()
  );
  wCopyDocBody = wCopyDocBody.replaceText("{{10tax}}", tentax.toLocaleString());
  wCopyDocBody = wCopyDocBody.replaceText(
    "{{8tax_t}}",
    eigtax_t.toLocaleString()
  );
  wCopyDocBody = wCopyDocBody.replaceText("{{8tax}}", eigtax.toLocaleString());
  wCopyDocBody = wCopyDocBody.replaceText(
    "{{amount_delivered}}",
    totals.toLocaleString()
  );
  wCopyDocBody = wCopyDocBody.replaceText(
    "{{total}}",
    "￥ " + totals.toLocaleString()
  );
  wCopyDoc.saveAndClose();

  // ファイル名を変更する
  let fileName =
    customerItem["会社名"] +
    "領収文書_" +
    Utilities.formatDate(new Date(), "JST", "yyyyMMdd");
  wCopyFile.setName(fileName);
  // コピーしたファイルIDとファイル名を返却する（あとでこのIDをもとにPDFに変換するため）
  return [wCopyFileId, fileName];
}
// 領収書のファイル生成
function createReceiptFile(records) {
  // PDF変換する元ファイルを作成する
  let wFileRtn = createReceiptGDoc(records);
  // PDF変換
  createReceiptPdf(wFileRtn[0], wFileRtn[1]);
  // PDF変換したあとは元ファイルを削除する
  DriveApp.getFileById(wFileRtn[0]).setTrashed(true);
}
