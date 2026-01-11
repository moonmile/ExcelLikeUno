## PivotTable (DataPilot) 実装計画

目的: Calc の DataPilot を Excel ライクに扱える PivotTable ラッパーを追加し、`sheet.pivot_tables.add(name, source_range, target_cell, layout=...)` のような自然なフローを提供する。既存パターン（UnoObject + typing Protocol + シート上のコレクションヘルパー）に倣う。

### マイルストーン
1) タイピング整備
- インターフェース定数/Protocol を追加: XDataPilotTables, XDataPilotTable2, XDataPilotDescriptor, XDataPilotField, XDataPilotFieldReference, フィールド/テーブル用 XPropertySet。
- 必要なら StructNames に CellAddress/CellRangeAddress を追加。

2) ラッパー設計
- 新モジュール `table/pivot_table.py` に `PivotTable` とコレクション `PivotTables`（ChartCollection と同様の形）を追加。
- `PivotTables.add(name, source_range, target_cell, filter_fields=None, row_fields=None, column_fields=None, data_fields=None)` でデスクリプタ作成→挿入する API を用意。
- `PivotTable` が持つ主要プロパティ: name, source_range, position(アンカーセル), refresh(), fields へのアクセス、show_details トグル、行/列/全体の合計フラグ。

3) Sheet への組み込み
- Sheet に `pivot_tables` プロパティを追加して PivotTables を返す。
- table/__init__.py とトップレベル __init__.py でエクスポートを追加。

4) サンプルとドキュメント
- 簡単なピボット（カテゴリ別行 + 値の合計）を作成するサンプルを samples/ に追加。
- README と agents/operation_spec.md に使い方を記載し、ToDo のスペルを PivotTable に修正。

5) テスト
- pytest（UNO 必須、環境なければ skip）：ピボット作成・名前確認・refresh・フィールド数確認、add/remove の冪等性を検証。

### オープンな検討事項
- フィールド指定は名前優先かインデックス許容か？ `fields["Name"]` ヘルパーを用意するか。
  → 名前優先で、インデックスは非推奨とする。
- レイアウト指定の簡易 DSL を用意するか（行/列/データ/フィルタをタプルや dict で渡す形式）。
    → 最初は個別引数で実装し、後で DSL を検討。
- グループ化・計算フィールド・小計/総計などの詳細を初期スコープに含めるか、後続にするか。
    → 初期スコープには含めず、後続で対応。
