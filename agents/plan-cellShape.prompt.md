## Plan: Cell/Shape リファクタリング

Cell/Shape をスタブ前提で型付けし直し、Excelライク API と UNO ライク API を整理する。

**Steps**
1. 現状棚卸し  
   - [src/excellikeuno/table/cell.py](src/excellikeuno/table/cell.py) の責務・差分を整理（value/formula/text、borders/font、CellProperties/CharacterProperties 委譲）。  
   - [src/excellikeuno/drawing/shape.py](src/excellikeuno/drawing/shape.py) と line/fill helpers を確認し、cell の方針と揃える。  
   - 境界生成ロジックは [src/excellikeuno/style/border.py](src/excellikeuno/style/border.py) の `Borders`/`BorderStyle` をベースに再利用ポイントを洗う。  
   - テスト期待を [tests/test_cell*.py](tests/test_cell.py) / [tests/test_cell_border.py](tests/test_cell_border.py) / [tests/test_shape*.py](tests/test_shape.py) で確認。

2. 型付けと依存整理  
   - スタブ [src/stubs/com/sun/star](src/stubs/com/sun/star) を直接 import し、InterfaceNames は query 用定数に限定。  
   - Cell/Shape で使う UNO 型（XCell, XPropertySet, TableBorder2, XShape 等）を明示インポートし、typing/calc 依存を削減。  
   - import 循環を避けるために TYPE_CHECKING で型のみ参照する箇所を決める。

3. Cell API 再設計  
   - 公開サーフェス: Excelライクは `value/formula/text`, `borders`, `font`, `back_color` など snake_case；UNOライクは PascalCase を property 経由で限定公開。  
   - `properties/character_properties` の委譲を明示 property 化し、`__getattr__` フォールバックを最小化。  
   - `Borders` プロキシ利用: Cell 内部で `_border_getter/_border_setter` を単純化し、TableBorder/TableBorder2 を一括同期。  
   - Number 判定で setValue/setFormula 分岐を整理。  
   - row/column ヘルパーは型付けを stubs ベースに更新。

4. Shape API 再設計  
   - `line`/`fill`/`position/size` を Excelライク snake_case で提供し、UNO props は `props` 経由。  
   - 既存 line/fill helpers を border と同様のプロキシ構造に合わせるか、共通ヘルパーを utils に切り出し。  
   - Group shapes などの取得メソッドも型付けして補完強化。

5. 実装・整合性  
   - 重複ヘルパー（border 正規化など）を１カ所に集約し、Cell/Range で共有。  
   - docs 更新: [agents/plan.md](agents/plan.md) と設計メモに新 API 方針と PYTHONPATH (src, src/stubs) を記載。

**Verification**
- テスト: LibreOffice Python で `pytest tests/test_cell.py tests/test_cell_border.py tests/test_cell2.py tests/test_cell2_border.py tests/test_shape.py tests/test_shape_fill.py tests/test_shape_line.py`。  
- 補完確認: VS Code で `Cell`, `Shape` の PascalCase/SnakeCase プロパティが解決するか。  

**Decisions**
- 型は stubs 優先、InterfaceNames は query 定数のみ。  
- Excelライク API は snake_case 最小セット、UNO ライクは明示 property 経由で露出。  
- border/line/fill のプロキシは既存 `Borders/BorderStyle` を再利用し、複製ロジックを共通化。
