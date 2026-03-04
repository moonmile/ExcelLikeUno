## Plan: Cell ラッパー再設計

Cell を新スタブ (src/stubs) 前提で型付けし直し、XCell / XPropertySet を統合した Excel ライク API を整理する。

**Steps**
1. 現状整理: [src/excellikeuno/table/cell.py](src/excellikeuno/table/cell.py) の責務（XCell 呼び出し、CellProperties/CharacterProperties 委譲、borders/font/row/col）を棚卸しし、不要な重複や例外ハンドリング箇所を洗う。
2. 依存整理: インターフェース定数経由の取得を整理し、型は [src/stubs/com/sun/star/table/XCell.pyi](src/stubs/com/sun/star/table/XCell.pyi) などのスタブを直接 import（InterfaceNames は定数用途に限定）。必要なら [src/stubs-add/com](src/stubs-add/com) の補助スタブも確認。
3. API 面の統合: `value/formula/text` の小文字アクセサを明示し、CellProperties/CharacterProperties の公開属性を直下プロパティで提供する方針を維持しつつ、__getattr__/__setattr__ のフォールバック範囲を最小化する（内部属性や property との競合を整理）。
4. ボーダー/フォント周り: `borders`/`font` プロキシと TableBorder(TableBorder2) の同期処理を見直し、BorderLine/BorderLine2 の正規化を共通化。stubs の型名を用いて引数・戻り値に型ヒントを付与。
5. シート関連ヘルパー: `row_height`/`column_width` など親シート解決ロジックで取得するインターフェースをスタブで型付けし、例外経路を整理。`position` は Point dataclass/stub に合わせて型明示。
6. インポート整理: PascalCase の UNO プロパティはそのまま、追加メソッドは snake_case に統一（design_guidelines 参照）。不要な古い typing/calc import を削り、stubs 由来のインターフェース/enum を採用。
7. ドキュメント反映: 方針変更・依存パス（src/stubs）を [agents/plan.md](agents/plan.md) と関連設計メモに追記し、利用例を更新（`sheet.cell(...).value` 等）。

**Verification**
- 単体: `python tools/idl_to_pyi.py` を再実行してスタブが最新であることを確認。
- テスト: `pytest tests/test_cell.py tests/test_cell_value.py tests/test_cell_properties.py tests/test_cell_border.py tests/test_cell_font.py` を LibreOffice Python で実行。
- IDE: VS Code で `from com.sun.star.table import XCell` 等が解決し、Cell の補完が出るか確認（必要ならウィンドウリロード）。

**Decisions**
- 型は src/stubs 配下の .pyi を優先利用し、InterfaceNames は定数参照に留める。
- 小文字アクセサは `value/formula/text` のみとし、それ以外は UNO の PascalCase を踏襲。
