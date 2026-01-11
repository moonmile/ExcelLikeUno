## Calc 保存/読み込み実装計画

目的: Calc ドキュメントの新規作成・既存ファイルを開く・保存・名前を付けて保存を UNO ラッパーで提供し、接続/ドキュメント API から一貫した操作を可能にする。

### スコープ
- 新規作成: 新しい Calc ドキュメントを生成（Hidden オプション）。
- 既存を開く: パス指定で開く（Hidden/ReadOnly/Template オプション）。
- 保存: 上書き保存。
- 名前を付けて保存: 別名で保存（必要なら FilterName/Overwrite 指定）。
- コピー保存: storeToURL によるコピー保存（任意）。

### 実装ポイント
1) 型と定数
- InterfaceNames に X_STORABLE (com.sun.star.frame.XStorable) を追加。
- typing/calc.py に XStorable Protocol (store, storeAsURL, storeToURL, hasLocation, isModified) を追加。

2) CalcDocument 拡張
- storable プロパティ経由で XStorable を取得。
- save(): store()
- save_as(path, filter_name=None, overwrite=True): storeAsURL(URL, [PropertyValue])
- save_copy(path, filter_name=None): storeToURL(URL, [PropertyValue])
- is_modified / has_location を補助プロパティで提供。
- パスは systemPathToFileUrl で URL 化。PropertyValue 配列生成ヘルパーを小実装。

3) connection モジュール拡張
- get_desktop ヘルパー（既存 connect_calc ロジックを共通化）。
- new_calc_document(hidden=True): private:factory/scalc を loadComponentFromURL。
- open_calc_document(path, hidden=True, read_only=False, as_template=False, filter_name=None): パスを URL 化して loadComponentFromURL。
- 既存 connect_calc のシグネチャは維持しつつ内部で共有ロジックを使用。
- ActiveDocument/ThisDocument, ActiveSheet/ThisSheet を取得するヘルパーを追加（デスクトップ/現在のコントローラ経由）。

4) パス/オプション
- Windows パスを考慮し、常に systemPathToFileUrl で URL 化。
- Hidden/ReadOnly/AsTemplate/FilterName/Overwrite を PropertyValue 配列で渡す。

5) サンプル・テスト
- samples: 新規→保存→再オープンのデモスクリプト。
- tests: UNO 環境前提で save/load ラウンドトリップを検証（不可なら skip）。

7) Active/This エントリ
- CalcDocument に ActiveSheet/ThisSheet アクセサを追加（現在のコントローラ or ThisComponent から現在のシートを返す）。
- connection 経由で ActiveDocument/ThisDocument を返すヘルパーを用意し、既存ドキュメントをラップして CalcDocument を返す。

6) ドキュメント
- README と agents/operation_spec.md に使い方を追記。
- ToDo を完了・更新するタイミングを合わせる。

8) シートの追加と削除
- CalcDocument.sheets.add(name) / remove(name) メソッドを追加し、シートの追加・削除をサポートする。
- Sheet の複製をサポートする場合は、add(name, from_sheet_name) のような形で実装を検討する。
