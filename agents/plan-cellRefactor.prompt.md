## Plan: Cell ラッパー再設計

Cell を新スタブ (src/stubs) 前提で型付けし直し、XCell / XPropertySet を統合した Excel ライク API を整理する。

**Steps**
1. 現状整理: [src/excellikeuno/table/cell.py](src/excellikeuno/table/cell.py) の責務（XCell 呼び出し、CellProperties/CharacterProperties 委譲、borders/font/row/col）を棚卸しし、不要な重複や例外ハンドリング箇所を洗う。
2. 依存整理: インターフェース定数経由の取得を整理し、型は [src/stubs/com/sun/star/table/XCell.pyi](src/stubs/com/sun/star/table/XCell.pyi) などスタブを直接 import（InterfaceNames はクエリ用定数に限定）。不足型は [src/stubs-add/com](src/stubs-add/com) を確認し、旧 typing/calc からの import は段階的に削除。
3. API 面の統合: 公開サーフェスを整理する。
	- 小文字アクセサは `value/formula/text` のみ維持（それ以外は PascalCase UNO 名）。
	- CellProperties/CharacterProperties の公開属性は明示的な property 経由で提供し、__getattr__/__setattr__ のフォールバックを最小化（内部属性・既存 property との競合を避ける）。
	- CellProperties は get_property/set_property の最小 API に縮退し、個別の PascalCase property は Cell 側で必要分のみ持つ。
	- props/properties/character_properties/borders/font のアクセサは型付きで残し、フォールバックは CellProperties への単純委譲のみに限定。
	- Number 判定による setValue/setFormula の分岐を整理（数値は setValue、それ以外は setFormula）。
	- 型注釈は src/stubs の UNO 型を用いて IDE 補完を強化。
4. ボーダー/フォント周り: `borders`/`font` プロキシと TableBorder(TableBorder2) の同期処理を見直す。
	- BorderLine/BorderLine2 の正規化ヘルパーを Cell 内に共通化（CellProperties からは除去済み）。
	- TableBorder/TableBorder2 の set/get と borders プロキシの適用ロジックを整理し、型ヒントは stubs の型を使用。
	- font/borders セッターは渡されたプロキシの `_current()` または dict を解釈して apply するだけに簡素化。
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
