## Calc 保存/読み込み、AcstiveCalcDocument/ActiveSheet 取得実装計画


目的: Calc ドキュメントの新規作成・既存ファイルを開く・保存・名前を付けて保存を UNO ラッパーで提供し、接続/ドキュメント API から一貫した操作を可能にする。

### スコープ

- Calc ドキュメントの新規作成/既存を開くなどのメソッドを実装する
- 現在アクティブな CalcDocument, Sheet を取得するメソッドを実装する

### 実装ポイント

1) global current_desktop を利用する

- excellikeuno.desktop モジュールに current_desktop 変数があるので、接続時に設定・切断時にクリアする

2) CalcDocument 保存/読み込み機能の実装

new_calc_document, open_calc_document を desktop のメソッドに変更する

- destkop.add_calc_document() -> CalcDocument      : 新規作成
- desktop.open_calc_document(path: str, ...) -> CalcDocument : 既存を開く

3) Active な CalcDocument/Sheet 取得機能の実装

- desktop.get_active_calc_document() -> CalcDocument : アクティブドキュメント取得
- CalcDocument.activate() : 指定の CalcDocument をアクティブにする
- Sheet.activate() : 現在の Sheet をアクティブにする 

4) エイリアス関数の追加

- ThisDesktop        : global desktop の alias
- ActiveCalcDocument : desktop.get_active_calc_document() の alias
- ActiveSheet       : アクティブシート取得の alias

- excellikeuno.ActiveCalcDocument にする
- ActiveCalcDocument = current_desktop.get_active_calc_document() で別名として定義する
- 型ヒントを付けるために CalcDocument 型と同じような protocol を定義する
