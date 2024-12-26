function selectShippingDate() {
  var nDate = new Date();
  var today = Utilities.formatDate(new Date(nDate), "JST", "yyyy/MM/dd");
  nDate.setDate(nDate.getDate() - 1);
  var yesterday = Utilities.formatDate(new Date(nDate), "JST", "yyyy/MM/dd");
  nDate.setDate(nDate.getDate() + 2);
  var tommorrow = Utilities.formatDate(new Date(nDate), "JST", "yyyy/MM/dd");
  nDate.setDate(nDate.getDate() + 1);
  var dayAfter2 = Utilities.formatDate(new Date(nDate), "JST", "yyyy/MM/dd");
  nDate.setDate(nDate.getDate() + 1);
  var dayAfter3 = Utilities.formatDate(new Date(nDate), "JST", "yyyy/MM/dd");
  const items = getAllRecords("受注");
  var yesterdayLists = [];
  var todayLists = [];
  var tommorrowLists = [];
  var dayAfter2Lists = [];
  var dayAfter3Lists = [];
  var yesterdayInvoiceTypes = [];
  var todayInvoiceTypes = [];
  var tommorrowInvoiceTypes = [];
  var dayAfter2InvoiceTypes = [];
  var dayAfter3InvoiceTypes = [];
  var yesterdayCnt = 0;
  var todayCnt = 0;
  var tommorrowCnt = 0;
  var dayAfter2Cnt = 0;
  var dayAfter3Cnt = 0;
  for (let i = 0; i < items.length; i++) {
    if (
      Utilities.formatDate(new Date(items[i]["発送日"]), "JST", "yyyy/MM/dd") ==
        yesterday &&
      items[i]["商品名"] != "送料"
    ) {
      const element = yesterdayLists.find(
        (value) =>
          value["発送先名"] === items[i]["発送先名"] &&
          value["クール区分"] === items[i]["クール区分"]
      );
      if (element) {
        const ele = element["商品"].find(
          (value) => value["商品名"] === items[i]["商品名"]
        );
        if (ele) {
          ele["受注数"] += Number(items[i]["受注数"]);
        } else {
          var product = {};
          var products = element["商品"].concat();
          product["商品名"] = items[i]["商品名"];
          product["受注数"] = Number(items[i]["受注数"]);
          product["メモ"] = items[i]["メモ"] ? items[i]["メモ"] : "";
          products.push(product);
          element["商品"] = products;
        }
      } else {
        if (
          yesterdayInvoiceTypes.some(
            (value) => value["納品方法"] == items[i]["納品方法"]
          )
        ) {
          yesterdayInvoiceTypes.forEach(function (wVal) {
            if (wVal["納品方法"] == items[i]["納品方法"]) {
              wVal["件数"]++;
              yesterdayCnt++;
            }
          });
        } else {
          var invoice = {};
          invoice["納品方法"] = items[i]["納品方法"];
          invoice["件数"] = 1;
          yesterdayCnt++;
          yesterdayInvoiceTypes.push(invoice);
        }
        var record = {};
        var product = {};
        var products = [];
        product["商品名"] = items[i]["商品名"];
        product["受注数"] = Number(items[i]["受注数"]);
        product["メモ"] = items[i]["メモ"] ? items[i]["メモ"] : "";
        products.push(product);
        record["顧客名"] = items[i]["顧客名"];
        record["発送先名"] = items[i]["発送先名"];
        record["クール区分"] = items[i]["クール区分"];
        record["商品"] = products;
        yesterdayLists.push(record);
      }
    }
    if (
      Utilities.formatDate(new Date(items[i]["発送日"]), "JST", "yyyy/MM/dd") ==
        today &&
      items[i]["商品名"] != "送料"
    ) {
      const element = todayLists.find(
        (value) =>
          value["発送先名"] === items[i]["発送先名"] &&
          value["クール区分"] === items[i]["クール区分"]
      );
      if (element) {
        const ele = element["商品"].find(
          (value) => value["商品名"] === items[i]["商品名"]
        );
        if (ele) {
          ele["受注数"] += Number(items[i]["受注数"]);
        } else {
          var product = {};
          var products = element["商品"].concat();
          product["商品名"] = items[i]["商品名"];
          product["受注数"] = Number(items[i]["受注数"]);
          product["メモ"] = items[i]["メモ"] ? items[i]["メモ"] : "";
          products.push(product);
          element["商品"] = products;
        }
      } else {
        if (
          todayInvoiceTypes.some(
            (value) => value["納品方法"] == items[i]["納品方法"]
          )
        ) {
          todayInvoiceTypes.forEach(function (wVal) {
            if (wVal["納品方法"] == items[i]["納品方法"]) {
              wVal["件数"]++;
              todayCnt++;
            }
          });
        } else {
          var invoice = {};
          invoice["納品方法"] = items[i]["納品方法"];
          invoice["件数"] = 1;
          todayCnt++;
          todayInvoiceTypes.push(invoice);
        }
        var record = {};
        var product = {};
        var products = [];
        product["商品名"] = items[i]["商品名"];
        product["受注数"] = Number(items[i]["受注数"]);
        product["メモ"] = items[i]["メモ"] ? items[i]["メモ"] : "";
        products.push(product);
        record["顧客名"] = items[i]["顧客名"];
        record["発送先名"] = items[i]["発送先名"];
        record["クール区分"] = items[i]["クール区分"];
        record["商品"] = products;
        todayLists.push(record);
      }
    }
    if (
      Utilities.formatDate(new Date(items[i]["発送日"]), "JST", "yyyy/MM/dd") ==
        tommorrow &&
      items[i]["商品名"] != "送料"
    ) {
      const element = tommorrowLists.find(
        (value) =>
          value["発送先名"] === items[i]["発送先名"] &&
          value["クール区分"] === items[i]["クール区分"]
      );
      if (element) {
        const ele = element["商品"].find(
          (value) => value["商品名"] === items[i]["商品名"]
        );
        if (ele) {
          ele["受注数"] += Number(items[i]["受注数"]);
        } else {
          var product = {};
          var products = element["商品"].concat();
          product["商品名"] = items[i]["商品名"];
          product["受注数"] = Number(items[i]["受注数"]);
          product["メモ"] = items[i]["メモ"] ? items[i]["メモ"] : "";
          products.push(product);
          element["商品"] = products;
        }
      } else {
        if (
          tommorrowInvoiceTypes.some(
            (value) => value["納品方法"] == items[i]["納品方法"]
          )
        ) {
          tommorrowInvoiceTypes.forEach(function (wVal) {
            if (wVal["納品方法"] == items[i]["納品方法"]) {
              wVal["件数"]++;
              tommorrowCnt++;
            }
          });
        } else {
          var invoice = {};
          invoice["納品方法"] = items[i]["納品方法"];
          invoice["件数"] = 1;
          tommorrowCnt++;
          tommorrowInvoiceTypes.push(invoice);
        }
        var record = {};
        var product = {};
        var products = [];
        product["商品名"] = items[i]["商品名"];
        product["受注数"] = Number(items[i]["受注数"]);
        product["メモ"] = items[i]["メモ"] ? items[i]["メモ"] : "";
        products.push(product);
        record["顧客名"] = items[i]["顧客名"];
        record["発送先名"] = items[i]["発送先名"];
        record["クール区分"] = items[i]["クール区分"];
        record["商品"] = products;
        tommorrowLists.push(record);
      }
    }
    if (
      Utilities.formatDate(new Date(items[i]["発送日"]), "JST", "yyyy/MM/dd") ==
        dayAfter2 &&
      items[i]["商品名"] != "送料"
    ) {
      const element = dayAfter2Lists.find(
        (value) =>
          value["発送先名"] === items[i]["発送先名"] &&
          value["クール区分"] === items[i]["クール区分"]
      );
      if (element) {
        const ele = element["商品"].find(
          (value) => value["商品名"] === items[i]["商品名"]
        );
        if (ele) {
          ele["受注数"] += Number(items[i]["受注数"]);
        } else {
          var product = {};
          var products = element["商品"].concat();
          product["商品名"] = items[i]["商品名"];
          product["受注数"] = Number(items[i]["受注数"]);
          product["メモ"] = items[i]["メモ"] ? items[i]["メモ"] : "";
          products.push(product);
          element["商品"] = products;
        }
      } else {
        if (
          dayAfter2InvoiceTypes.some(
            (value) => value["納品方法"] == items[i]["納品方法"]
          )
        ) {
          dayAfter2InvoiceTypes.forEach(function (wVal) {
            if (wVal["納品方法"] == items[i]["納品方法"]) {
              wVal["件数"]++;
              dayAfter2Cnt++;
            }
          });
        } else {
          var invoice = {};
          invoice["納品方法"] = items[i]["納品方法"];
          invoice["件数"] = 1;
          dayAfter2Cnt++;
          dayAfter2InvoiceTypes.push(invoice);
        }
        var record = {};
        var product = {};
        var products = [];
        product["商品名"] = items[i]["商品名"];
        product["受注数"] = Number(items[i]["受注数"]);
        product["メモ"] = items[i]["メモ"] ? items[i]["メモ"] : "";
        products.push(product);
        record["顧客名"] = items[i]["顧客名"];
        record["発送先名"] = items[i]["発送先名"];
        record["クール区分"] = items[i]["クール区分"];
        record["商品"] = products;
        dayAfter2Lists.push(record);
      }
    }
    if (
      Utilities.formatDate(new Date(items[i]["発送日"]), "JST", "yyyy/MM/dd") ==
        dayAfter3 &&
      items[i]["商品名"] != "送料"
    ) {
      const element = dayAfter3Lists.find(
        (value) =>
          value["発送先名"] === items[i]["発送先名"] &&
          value["クール区分"] === items[i]["クール区分"]
      );
      if (element) {
        const ele = element["商品"].find(
          (value) => value["商品名"] === items[i]["商品名"]
        );
        if (ele) {
          ele["受注数"] += Number(items[i]["受注数"]);
        } else {
          var product = {};
          var products = element["商品"].concat();
          product["商品名"] = items[i]["商品名"];
          product["受注数"] = Number(items[i]["受注数"]);
          product["メモ"] = items[i]["メモ"] ? items[i]["メモ"] : "";
          products.push(product);
          element["商品"] = products;
        }
      } else {
        if (
          dayAfter3InvoiceTypes.some(
            (value) => value["納品方法"] == items[i]["納品方法"]
          )
        ) {
          dayAfter3InvoiceTypes.forEach(function (wVal) {
            if (wVal["納品方法"] == items[i]["納品方法"]) {
              wVal["件数"]++;
              dayAfter3Cnt++;
            }
          });
        } else {
          var invoice = {};
          invoice["納品方法"] = items[i]["納品方法"];
          invoice["件数"] = 1;
          dayAfter3Cnt++;
          dayAfter3InvoiceTypes.push(invoice);
        }
        var record = {};
        var product = {};
        var products = [];
        product["商品名"] = items[i]["商品名"];
        product["受注数"] = Number(items[i]["受注数"]);
        product["メモ"] = items[i]["メモ"] ? items[i]["メモ"] : "";
        products.push(product);
        record["顧客名"] = items[i]["顧客名"];
        record["発送先名"] = items[i]["発送先名"];
        record["クール区分"] = items[i]["クール区分"];
        record["商品"] = products;
        dayAfter3Lists.push(record);
      }
    }
  }
  var dateStr = {};
  dateStr["昨日"] = Utilities.formatDate(new Date(yesterday), "JST", "MM/dd");
  dateStr["今日"] = Utilities.formatDate(new Date(today), "JST", "MM/dd");
  dateStr["明日"] = Utilities.formatDate(new Date(tommorrow), "JST", "MM/dd");
  dateStr["明後日"] = Utilities.formatDate(new Date(dayAfter2), "JST", "MM/dd");
  dateStr["明々後日"] = Utilities.formatDate(
    new Date(dayAfter3),
    "JST",
    "MM/dd"
  );

  var cnt = {};
  cnt["昨日"] = yesterdayCnt;
  cnt["今日"] = todayCnt;
  cnt["明日"] = tommorrowCnt;
  cnt["明後日"] = dayAfter2Cnt;
  cnt["明々後日"] = dayAfter3Cnt;
  var datas = [];
  datas.push(dateStr);
  datas.push(cnt);
  datas.push(yesterdayInvoiceTypes);
  datas.push(yesterdayLists);
  datas.push(todayInvoiceTypes);
  datas.push(todayLists);
  datas.push(tommorrowInvoiceTypes);
  datas.push(tommorrowLists);
  datas.push(dayAfter2InvoiceTypes);
  datas.push(dayAfter2Lists);
  datas.push(dayAfter3InvoiceTypes);
  datas.push(dayAfter3Lists);
  return JSON.stringify(datas);
}
