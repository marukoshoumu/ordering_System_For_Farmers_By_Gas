# モダンダイアログ統一化 設計書

## 概要

`window.confirm` (22箇所) と `window.alert` (約45箇所) を、HTML標準 `<dialog>` 要素とトースト通知に置き換え、UI を統一する。

## 方針

| 対象 | 置き換え先 | 方式 |
|------|-----------|------|
| `window.confirm` | `showConfirmDialog()` | `<dialog>` + `Promise<boolean>` |
| `alert`（バリデーション） | `showWarningToast()` + `return` | 既存トースト |
| `alert`（処理失敗） | `showErrorToast()` | 既存トースト |
| `alert`（情報通知） | `showInfoToast()` | 既存トースト |

## showConfirmDialog コンポーネント

### 配置場所

`utilities.html` に追加（既存トーストと同じファイル）

### API

```js
showConfirmDialog(message, options?) → Promise<boolean>
```

- `<dialog>` 要素を動的に生成し `showModal()` で表示
- 「キャンセル」「OK」の2ボタン（シンプル統一デザイン）
- `::backdrop` でオーバーレイ表示
- OKで `true`、キャンセルで `false` を resolve
- ESCキーでキャンセル扱い
- 既存トーストと統一感のあるスタイリング

### 呼び出し側の変更

```js
// Before
if (!confirm('削除しますか？')) return;

// After
if (!(await showConfirmDialog('削除しますか？'))) return;
// 親関数に async を付与
```

## alert 置き換えルール

- `return alert(...)` → `showWarningToast(...); return;`
- `catch` 内の `alert(...)` → `showErrorToast(...)`
- 情報通知の `alert(...)` → `showInfoToast(...)`
- `orderList.html` のローカル `showSuccessToast` は削除し、共通版に統一

## 変更対象ファイル

### 共通コンポーネント（1ファイル）

- `utilities.html` — `showConfirmDialog()` + CSS 追加

### confirm 置き換え（11ファイル・22箇所）

| ファイル | 箇所数 | 内容 |
|---|---|---|
| `orderList.html` | 7 | 一括操作（出荷済み、キャンセル、削除等） |
| `createFreeeDeliveryNote.html` | 2 | freee CSV作成確認 |
| `recurringOrderList.html` | 3 | 定期便の一時停止/再開/削除 |
| `aiImportList.html` | 2 | 仮受注の削除/再登録 |
| `createBill.html` | 1 | 請求書作成確認 |
| `masterImport.html` | 1 | インポート確認 |
| `salesDashboardScript.html` | 1 | 全分析実行確認 |
| `csvImport.html` | 1 | CSVインポート確認 |
| `createShippingSlips.html` | 1 | 発送伝票作成確認 |
| `shippingCustomerScript.html` | 2 | 顧客/発送先登録確認 |
| `quotation.html` | 2 | 顧客登録/見積書発行確認 |

### alert 置き換え（10ファイル・約45箇所）

| ファイル | 箇所数 |
|---|---|
| `orderList.html` | 18 |
| `recurringOrderList.html` | 8 |
| `home.html` | 5 |
| `scanner.html` | 5 |
| `HeatmapOrderList.html` | 3 |
| `productList.html` | 1 |
| `sidebar.html` | 1 |
| `salesDashboardScript.html` | 2 |
| `test-components.html` | 2 |

### 変更しないもの

- `aiAssistantUi.html` の `displayAlerts` 等（関数名であり `window.alert` ではない）
- 既にトーストを使用している箇所
