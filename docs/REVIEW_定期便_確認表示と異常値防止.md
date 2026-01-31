# レビュー：定期便「確認画面表示」と「異常値防止」の修正

## 対象修正

1. 確認画面に「次回発送日」「次回納品日」を表示する  
2. 納品日が発送日より前になる入力を防止する（受注登録時・間隔変更時）

---

## レビュー結果サマリ

| 項目 | 結果 | 備考 |
|------|------|------|
| 1. 確認画面の次回日付表示 | ✅ 問題なし | 計算・表示・空時は「-」まで妥当 |
| 2. 受注登録時の異常値防止 | ✅ 問題なし | 定期便時のみ・critical 警告で阻止 |
| 3. 間隔変更時の異常値防止 | ✅ 問題なし | alert で送信前に阻止 |

**総合: 要件を満たしており、そのまま採用してよい実装です。**

---

## 1. 確認画面に次回発送日・次回納品日を表示（orderCode.js）

**確認した内容**

- `e.parameter.shippingDate1` / `deliveryDate1` が両方ある場合のみ計算している。
- `new Date(...)` のあと `!isNaN(...getTime())` で無効日付を弾いている。
- `calcNextShippingDateFlexible(baseShippingDate, intervalObj)` で次回発送日を算出。
- `daysDiff = Math.round((baseDeliveryDate - baseShippingDate) / (1000*60*60*24))` で発送→納品の日数差を取得し、次回納品日 = 次回発送日 + daysDiff で計算。
- 表示は `Utilities.formatDate(..., 'JST', 'yyyy/MM/dd')` で統一。
- 日付が無い／無効のときは `nextShippingDateDisplay` / `nextDeliveryDateDisplay` を `'-'` のままにしている。
- 定期便ブロック内に「次回発送日」「次回納品日」の2行を追加し、既存の間隔表示の下に配置している。

**結論**: 仕様どおりで問題ありません。

---

## 2. 受注登録時の異常値防止（shippingValidation.html）

**確認した内容**

- `checkBasicValidation()` 内で、`#isRecurringOrder` が checked のときだけ発送日・納品日をチェックしている。
- `input[name="shippingDate1"]` / `input[name="deliveryDate1"]` を取得（`getElementById` のフォールバックあり）。
- 両方に値があるときだけ `new Date(...)` で比較し、`deliveryDateObj < shippingDateObj` のときだけ警告を追加。
- 警告は `type: 'critical'`, `icon: '🔴'` で、既存の `showWarningModal(warnings)` で表示されるため送信が止まる。
- タイトル・詳細メッセージで「納品日が発送日より前」「定期便の計算が正しく行われません」と説明されている。

**結論**: 要件どおりで問題ありません。

---

## 3. 間隔変更時の異常値防止（recurringOrderList.html）

**確認した内容**

- `confirmIntervalChange()` で「次回納品日を入力してください」チェックの直後に、次回発送日・次回納品日の両方がある場合のチェックを追加。
- `newNextShippingDate` / `newNextDeliveryDate` を `new Date(...)` にし、`getTime()` が有効な場合に `deliveryDateObj < shippingDateObj` を判定。
- 条件を満たすときは `alert` でメッセージを表示し `return` してサーバー送信を行わない。
- メッセージに発送日・納品日の値と「正しい日付を入力してください」が含まれている。

**結論**: 要件どおりで問題ありません。

---

## 軽微な補足（任意）

- **同日（発送日＝納品日）**: いずれも `deliveryDateObj < shippingDateObj` のみで判定しているため、同日はエラーにせず許可されている。業務上同日を許容するなら現状のままでよい。
- **間隔変更で「次回発送日」が未入力のとき**: 現状は納品日のみ必須で、発送日が空の場合は大小チェックをスキップしている。運用で「次回発送日も必須」にしたい場合は、別途「次回発送日を入力してください」などのチェックを追加するとよい。

---

## 総合

- 確認画面で次回発送日・次回納品日が正しく計算・表示されている。  
- 受注登録時・間隔変更時の両方で、納品日が発送日より前の場合は適切に警告され、送信が止まる。  

**修正内容は要件を満たしており、そのまま採用してよい水準です。**
