# モダンダイアログ統一化 実装計画

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** `window.confirm` (22箇所) と `window.alert` (約45箇所) を `<dialog>` 要素とトースト通知に置き換え、UIを統一する

**Architecture:** `utilities.html` に `showConfirmDialog()` を追加し、各画面の `confirm()` を `async/await` で呼び出す。`alert()` は既存の `showToast` 系関数に置き換える。

**Tech Stack:** HTML `<dialog>` 要素、既存 `utilities.html` トーストシステム、CSS変数（`components.html` 定義済み）

**設計書:** `docs/plans/2026-02-28-modern-dialog-design.md`

---

### Task 1: showConfirmDialog コンポーネント追加

**Files:**
- Modify: `src/components.html` — ダイアログCSS追加
- Modify: `src/utilities.html` — `showConfirmDialog()` 関数追加

**Step 1: `components.html` にダイアログCSS追加**

トーストCSS（`.toast-info` の直後、レスポンシブ対応の前）に以下を追加:

```css
  /* ===================================
     確認ダイアログコンポーネント
     =================================== */

  .confirm-dialog {
    border: none;
    border-radius: 12px;
    padding: 0;
    max-width: 400px;
    width: calc(100% - 32px);
    box-shadow: 0 8px 32px rgba(0, 0, 0, 0.15);
    font-family: inherit;
  }

  .confirm-dialog::backdrop {
    background: rgba(0, 0, 0, 0.4);
  }

  .confirm-dialog-body {
    padding: 24px;
  }

  .confirm-dialog-message {
    font-size: 0.95rem;
    line-height: 1.6;
    color: var(--text-primary);
    white-space: pre-line;
  }

  .confirm-dialog-actions {
    display: flex;
    justify-content: flex-end;
    gap: 8px;
    padding: 12px 24px;
    border-top: 1px solid #e5e7eb;
    background: var(--bg);
    border-radius: 0 0 12px 12px;
  }

  .confirm-dialog-btn {
    padding: 8px 20px;
    border-radius: 8px;
    font-size: 0.875rem;
    font-weight: 500;
    cursor: pointer;
    transition: all 0.15s ease;
    border: 1px solid #e5e7eb;
  }

  .confirm-dialog-btn-cancel {
    background: var(--card-bg);
    color: var(--text-secondary, #6b7280);
  }

  .confirm-dialog-btn-cancel:hover {
    background: #f3f4f6;
  }

  .confirm-dialog-btn-ok {
    background: var(--primary);
    color: #fff;
    border-color: var(--primary);
  }

  .confirm-dialog-btn-ok:hover {
    background: var(--primary-dark);
    border-color: var(--primary-dark);
  }
```

**Step 2: `utilities.html` にshowConfirmDialog関数追加**

`showInfoToast` 関数の直後に以下を追加（DOM APIでXSS安全に構築）:

```js
  /**
   * 確認ダイアログを表示（window.confirm の代替）
   * @param {string} message - 表示するメッセージ
   * @returns {Promise<boolean>} OKならtrue、キャンセルならfalse
   */
  function showConfirmDialog(message) {
    return new Promise(function(resolve) {
      var dialog = document.createElement('dialog');
      dialog.className = 'confirm-dialog';

      var body = document.createElement('div');
      body.className = 'confirm-dialog-body';
      var msg = document.createElement('div');
      msg.className = 'confirm-dialog-message';
      msg.textContent = message;
      body.appendChild(msg);

      var actions = document.createElement('div');
      actions.className = 'confirm-dialog-actions';
      var btnCancel = document.createElement('button');
      btnCancel.type = 'button';
      btnCancel.className = 'confirm-dialog-btn confirm-dialog-btn-cancel';
      btnCancel.textContent = 'キャンセル';
      var btnOk = document.createElement('button');
      btnOk.type = 'button';
      btnOk.className = 'confirm-dialog-btn confirm-dialog-btn-ok';
      btnOk.textContent = 'OK';
      actions.appendChild(btnCancel);
      actions.appendChild(btnOk);

      dialog.appendChild(body);
      dialog.appendChild(actions);

      function cleanup(result) {
        dialog.close();
        dialog.remove();
        resolve(result);
      }

      btnOk.addEventListener('click', function() { cleanup(true); });
      btnCancel.addEventListener('click', function() { cleanup(false); });
      dialog.addEventListener('cancel', function(e) {
        e.preventDefault();
        cleanup(false);
      });

      document.body.appendChild(dialog);
      dialog.showModal();
      btnOk.focus();
    });
  }
```

**Step 3: 動作確認**

`test-components.html` にテストボタンを追加して手動検証:
- OKクリック → `true` が返る
- キャンセルクリック → `false` が返る
- ESCキー → `false` が返る
- メッセージ内の改行 `\n` が正しく表示される

**Step 4: コミット**

```bash
git add src/components.html src/utilities.html
git commit -m "feat: <dialog>ベースのshowConfirmDialogコンポーネントを追加"
```

---

### Task 2: orderList.html — confirm 置き換え（7箇所）+ alert 置き換え（18箇所）

**Files:**
- Modify: `src/orderList.html`

**Step 1: confirm → showConfirmDialog（7箇所）**

各 `confirm()` を `await showConfirmDialog()` に変更し、親関数に `async` を付与。

対象箇所（行番号は目安、作業時に再確認）:
- L1688: `if (!confirm('全期間のデータを取得します...'))` → `if (!(await showConfirmDialog('全期間のデータを取得します...')))` + 親関数 async 化
- L2682: `if (!confirm(orderIds.length + '件の受注を出荷済みにしますか？'))` → 同様
- L2708: `if (!confirm(orderIds.length + '件の受注をキャンセルしますか？...'))` → 同様
- L2734: `if (!confirm(orderIds.length + '件の受注を「収穫待ち」に...'))` → 同様
- L2759: `if (!confirm('⚠️ ' + orderIds.length + '件の受注を完全に削除しますか？...'))` → 同様
- L2764: `if (!confirm('本当に削除してよろしいですか？...'))` → 同様
- L2790: `if (!confirm(orderIds.length + '件の受注のキャンセルを解除しますか？'))` → 同様

**Step 2: alert → トースト（18箇所）**

置き換えルール:
- `alert('受注を選択してください')` → `showWarningToast('受注を選択してください'); return;`
- `alert('更新エラー: ' + error.message)` → `showErrorToast('更新エラー: ' + error.message)`
- `alert('データ取得エラー: ...')` → `showErrorToast('データ取得エラー: ...')`
- `return alert(...)` → `showWarningToast(...); return;`
- OCR関連の `alert("認識不可: ...")` → `showWarningToast("認識不可: ...")`
- `alert("OCRエラー")` → `showErrorToast("OCRエラー")`
- `alert("キャプチャ失敗")` → `showErrorToast("キャプチャ失敗")`
- `alert('受注IDが取得できませんでした')` → `showErrorToast('受注IDが取得できませんでした')`

**Step 3: ローカル showSuccessToast 削除**

`orderList.html` L2809付近のローカル定義 `showSuccessToast` を削除（`utilities.html` の共通版を使用）。

**Step 4: コミット**

```bash
git add src/orderList.html
git commit -m "refactor(orderList): confirm/alertをモダンダイアログ/トーストに置き換え"
```

---

### Task 3: recurringOrderList.html — confirm（3箇所）+ alert（8箇所）

**Files:**
- Modify: `src/recurringOrderList.html`

**Step 1: confirm → showConfirmDialog（3箇所）**

- L1005: `if (!confirm('この定期便を一時停止しますか？')) return;` → `if (!(await showConfirmDialog('この定期便を一時停止しますか？'))) return;` + async 化
- L1021: `if (!confirm('この定期便を再開しますか？')) return;` → 同様
- L1037: `if (!confirm('この定期便を削除しますか？...')) return;` → 同様

**Step 2: alert → トースト（8箇所）**

- `alert('データの取得に失敗しました')` → `showErrorToast('データの取得に失敗しました')`
- `alert('次回納品日を入力してください')` → `showWarningToast('次回納品日を入力してください')`
- `alert('納品日が発送日より前になっています...')` → `showWarningToast('納品日が発送日より前になっています...')`
- `alert('変更に失敗しました')` → `showErrorToast('変更に失敗しました')`
- `alert('エラーが発生しました: ' + error)` → `showErrorToast('エラーが発生しました: ' + error)`
- `alert('一時停止に失敗しました')` → `showErrorToast('一時停止に失敗しました')`
- `alert('再開に失敗しました')` → `showErrorToast('再開に失敗しました')`
- `alert('削除に失敗しました')` → `showErrorToast('削除に失敗しました')`

**Step 3: コミット**

```bash
git add src/recurringOrderList.html
git commit -m "refactor(recurringOrderList): confirm/alertをモダンダイアログ/トーストに置き換え"
```

---

### Task 4: home.html — alert 置き換え（5箇所）

**Files:**
- Modify: `src/home.html`

**Step 1: alert → トースト（5箇所）**

- `alert('更新エラー: ' + error.message)` (2箇所) → `showErrorToast('更新エラー: ' + error.message)`
- `alert("認識不可: " + res.message)` → `showWarningToast("認識不可: " + res.message)`
- `alert("OCRエラー")` → `showErrorToast("OCRエラー")`
- `alert("キャプチャ失敗")` → `showErrorToast("キャプチャ失敗")`

**Step 2: コミット**

```bash
git add src/home.html
git commit -m "refactor(home): alertをトースト通知に置き換え"
```

---

### Task 5: aiImportList.html — confirm（2箇所）

**Files:**
- Modify: `src/aiImportList.html`

**Step 1: confirm → showConfirmDialog（2箇所）**

- L945: `if (!confirm('この仮受注データを完全に削除しますか？...'))` → `if (!(await showConfirmDialog('この仮受注データを完全に削除しますか？\n削除後は復元できません。')))` + async 化
- L972: `if (!confirm('この仮受注を再登録しますか？...'))` → 同様

**Step 2: コミット**

```bash
git add src/aiImportList.html
git commit -m "refactor(aiImportList): confirmをモダンダイアログに置き換え"
```

---

### Task 6: createFreeeDeliveryNote.html — confirm（2箇所）

**Files:**
- Modify: `src/createFreeeDeliveryNote.html`

**Step 1: confirm → showConfirmDialog（2箇所）**

- L343: `window.confirm('freee納品書CSVを作成してもよろしいですか？')` → `await showConfirmDialog(...)` + async 化
- L718: `window.confirm(confirmMsg)` → `await showConfirmDialog(confirmMsg)` + async 化

**Step 2: コミット**

```bash
git add src/createFreeeDeliveryNote.html
git commit -m "refactor(createFreeeDeliveryNote): confirmをモダンダイアログに置き換え"
```

---

### Task 7: 残りの小規模ファイル — confirm + alert 一括置き換え

**Files:**
- Modify: `src/createBill.html` — confirm 1箇所
- Modify: `src/masterImport.html` — confirm 1箇所
- Modify: `src/salesDashboardScript.html` — confirm 1箇所 + alert 2箇所
- Modify: `src/csvImport.html` — confirm 1箇所
- Modify: `src/createShippingSlips.html` — confirm 1箇所
- Modify: `src/shippingCustomerScript.html` — confirm 2箇所
- Modify: `src/quotation.html` — confirm 2箇所

**Step 1: 各ファイルの confirm を showConfirmDialog に置き換え**

全て同じパターン:
- `window.confirm(msg)` or `confirm(msg)` → `await showConfirmDialog(msg)`
- 親関数に `async` を付与

**Step 2: salesDashboardScript.html の alert 2箇所をトーストに置き換え**

- `alert('サーバー接続成功: ...')` → `showSuccessToast('サーバー接続成功')`
- `alert('サーバー接続失敗: ...')` → `showErrorToast('サーバー接続失敗: ' + error.message)`

**Step 3: コミット**

```bash
git add src/createBill.html src/masterImport.html src/salesDashboardScript.html src/csvImport.html src/createShippingSlips.html src/shippingCustomerScript.html src/quotation.html
git commit -m "refactor: 残り7ファイルのconfirm/alertをモダンダイアログ/トーストに置き換え"
```

---

### Task 8: 残りの alert のみのファイル

**Files:**
- Modify: `src/scanner.html` — alert 5箇所
- Modify: `src/HeatmapOrderList.html` — alert 3箇所
- Modify: `src/productList.html` — alert 1箇所
- Modify: `src/sidebar.html` — alert 1箇所
- Modify: `src/test-components.html` — alert 2箇所

**Step 1: 各ファイルの alert をトーストに置き換え**

- scanner.html: OCR/キャプチャ系エラーを `showErrorToast` / `showWarningToast` に
- HeatmapOrderList.html: バリデーション/検索エラーを `showWarningToast` / `showErrorToast` に
- productList.html: `alert('部門を先に選択してください')` → `showWarningToast(...); return;`
- sidebar.html: `alert("URLの取得に失敗しました。")` → `showErrorToast(...)`
- test-components.html: `alert(...)` → `showInfoToast(...)`

**Step 2: コミット**

```bash
git add src/scanner.html src/HeatmapOrderList.html src/productList.html src/sidebar.html src/test-components.html
git commit -m "refactor: 残り5ファイルのalertをトースト通知に置き換え"
```

---

### Task 9: 最終確認・クリーンアップ

**Step 1: 全ファイル検索で残留チェック**

```bash
grep -rn "window\.confirm\|[^a-zA-Z]confirm(" src/ --include="*.html" --include="*.js" | grep -v "showConfirmDialog\|confirmMsg\|confirmMessage\|Confirm\|confirm-dialog"
grep -rn "[^a-zA-Z]alert(" src/ --include="*.html" --include="*.js" | grep -v "showAlert\|alertEl\|displayAlert\|Alert\|alert-"
```

残留があれば対応。

**Step 2: コミット（残留があった場合のみ）**

```bash
git commit -m "refactor: confirm/alert残留箇所のクリーンアップ"
```
