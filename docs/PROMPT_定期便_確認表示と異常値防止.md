# ClaudeCode 実装依頼プロンプト：定期便「確認画面表示」と「異常値防止」

以下をそのまま ClaudeCode（または実装担当AI）に渡して使ってください。

---

## 依頼文（コピー用）

```
このプロジェクトは Google Apps Script（GAS）の Web アプリで、受注管理・定期便機能があります。

次の2点を実装してください。

---

### 1. 確認画面に「次回発送日」「次回納品日」を表示する

**背景**
- 定期便として登録するとき、確認画面では「定期便として登録」「間隔（例: 毎月3日）」しか表示されておらず、計算された「次回発送日・次回納品日」が見えない。
- ユーザーが「次回はいつになるか」を確認できるようにしたい。

**対象ファイル・箇所**
- `src/orderCode.js` の確認画面HTML生成部分。定期便ブロック（「定期便として登録」＋間隔表示）を出力している箇所。
- 該当ブロックは `e.parameter.isRecurringOrder === 'true'` のときで、すでに `intervalObj`（間隔オブジェクト）、`intervalDisplay`、`intervalJson` を組み立てている直後です。
- 参考: `recurringOrderCode.js` の `createRecurringOrder` では、`baseShippingDate`（= shippingDate1）、`baseDeliveryDate`（= deliveryDate1）から「次回発送日」「次回納品日」を計算しています。

**実装内容**
- 確認画面の定期便ブロック内に、**次回発送日**と**次回納品日**を表示する行を追加する。
- 計算方法は `recurringOrderCode.js` と同様にする:
  - 基準: `e.parameter.shippingDate1`（発送日）、`e.parameter.deliveryDate1`（納品日）
  - 次回発送日 = `calcNextShippingDateFlexible(baseShippingDate, intervalObj)`（`recurringOrderCode.js` に定義済み）
  - 発送日〜納品日の日数差 `daysDiff` = (納品日 - 発送日) の日数
  - 次回納品日 = 次回発送日 + daysDiff（日）
- 表示形式: 「次回発送日: yyyy/MM/dd」「次回納品日: yyyy/MM/dd」のように読みやすい形でよい（既存の `Utilities.formatDate` や JST を利用してよい）。
- `calcNextShippingDateFlexible` は `recurringOrderCode.js` で定義されているので、orderCode.js から呼び出してよい（同一GASプロジェクトで共通スコープ）。

**注意**
- `shippingDate1` や `deliveryDate1` が空の場合は、次回日付の計算・表示をスキップするか、「-」などで表示すること。

---

### 2. 納品日が発送日より前になる入力を防ぐ（異常値防止）

**背景**
- 発送日より納品日が前（例: 発送 2/5・納品 2/3）だと、定期便の「発送日→納品日の間隔」が負になり、次回以降の計算がおかしくなる。登録時と間隔変更時の両方で警告したい。

**2-1. 受注登録時（受注画面 → 確認画面へ送信する前）**

- **対象ファイル**: `src/shippingValidation.html`
- **対象処理**: `performLocalCheck()` から呼ばれている `checkBasicValidation()` に、次のチェックを追加する（または `performLocalCheck()` 内で、`checkBasicValidation()` の結果に追加する形でも可）。
- **ロジック**:
  - 「定期便として登録する」にチェックが入っているか（`#isRecurringOrder` が checked）を判定する。
  - チェックが入っている場合にのみ、**1件目の日程**の発送日・納品日を取得する（`input[name="shippingDate1"]` と `input[name="deliveryDate1"]`、または id で `shippingDate1` / `deliveryDate1` があればそれを使用）。
  - 両方に値があるとき、納品日（Date）が発送日（Date）より**前**なら、警告を1件追加する。
  - 警告の種類: **critical** 推奨（納品日が発送日より前は業務上あり得ないため）。
  - メッセージ例: タイトル「発送日と納品日の関係を確認してください」、詳細「納品日は発送日以降の日付を指定してください。」など。
- 既存の `showWarningModal(warnings)` で表示されるので、`warnings` に同じ形式のオブジェクト（`type`, `icon`, `title`, `detail`）を追加すればよい。

**2-2. 定期便一覧での間隔・次回日付変更時**

- **対象ファイル**: `src/recurringOrderList.html`
- **対象処理**: `confirmIntervalChange()` 関数。すでに「次回納品日が未入力」のチェックがある。
- **追加内容**: 「次回発送日」「次回納品日」の両方に値が入っている場合、**納品日が発送日より前**なら `alert` で警告し、`return` してサーバーへ送信しない。
- 比較は、文字列 `yyyy-MM-dd` のまま比較するか、`new Date(値)` で Date に変換して比較してよい。

**注意**
- 受注画面では複数日程（shippingDate2/deliveryDate2 など）がある場合は、少なくとも 1 件目（shippingDate1 / deliveryDate1）について上記チェックをすればよい。
```

---

## 補足（実装者向け）

- **確認画面の表示位置**: `orderCode.js` で `html += ... 定期便として登録 ...` と `intervalDisplay` を出している直後、`</div>` の前などに「次回発送日」「次回納品日」の2行を追加するイメージです。
- **calcNextShippingDateFlexible**: `recurringOrderCode.js` の 38 行目付近。`baseDate` と `intervalObj` を受け取り、次回発送日（Date）を返します。
- **日数差の計算**: `(baseDeliveryDate - baseShippingDate) / (1000 * 60 * 60 * 24)` を `Math.round` した値が daysDiff です。createRecurringOrder 内の 257 行付近を参照してください。
- **shippingValidation のフォーム要素**: 受注画面は `shipping.html` で、`shippingValidation.html` は include されており、同じフォーム内の `#isRecurringOrder`、`input[name="shippingDate1"]`、`input[name="deliveryDate1"]` を参照できます。名前や id は必要に応じて shipping.html で確認してください。
```
