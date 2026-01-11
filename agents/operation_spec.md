# Calc の操作仕様書

## 保存/読み込み

- `new_calc_document(hidden=True)` : private:factory/scalc で新規作成し `(desktop, CalcDocument, first_sheet)` を返す。
- `open_calc_document(path, hidden=True, read_only=False, as_template=False, filter_name=None)` : パスを UNO URL 化して開く。Hidden/ReadOnly などは `PropertyValue` で渡す。
- `CalcDocument.save()` : 上書き保存。
- `CalcDocument.save_as(path, filter_name=None, overwrite=True)` : `storeAsURL` を呼び出す。`Overwrite` をデフォルトで有効化。
- `CalcDocument.save_copy(path, filter_name=None, overwrite=True)` : `storeToURL` を呼び出すコピー保存。
- `CalcDocument.has_location` / `is_modified` : XStorable の `hasLocation` / `isModified` を返す。

## アクティブ/カレントの取得

- `active_document()` / `this_document()` : デスクトップの `getCurrentComponent()` を `CalcDocument` でラップして返す。
- `active_sheet()` / `this_sheet()` : 現在のコントローラが保持するアクティブシートを `Sheet` で返す。
- `CalcDocument.active_sheet` / `CalcDocument.this_sheet` : ドキュメントインスタンスから同様のシート取得を行う。

## シート操作

- `CalcDocument.add_sheet(name, index=None, from_sheet_name=None)` : 指定位置にシートを追加。`from_sheet_name` 指定時は `copyByName` があれば複製する。
- `CalcDocument.copy_sheet(source_name, new_name, index=None)` : UNO の `copyByName` を利用してシート複製（未サポート環境では AttributeError）。
- `CalcDocument.remove_sheet(name)` : シート削除。
- `Sheet.copy_to(new_name, index=None)` : ドキュメント参照付きシートでコピーを行う簡便メソッド。

