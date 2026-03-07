Plan: SheetCell replacement for Cell

VBAライクな SheetCell を新設し、最終的に既存 Cell を置き換える移行パスを作る。

Steps
1. 現状確認: src/excellikeuno/table/cell.py の値/式/文字列/Borders/Font 実装と src/excellikeuno/table/cell_properties.py のプロパティ委譲を再確認。依存する src/excellikeuno/style/font.py、src/excellikeuno/style/border.py、agents/design_guidelines.md も確認。
2. 新実装: SheetCell を新モジュールに定義し、raw/props/borders/font/value/formula/text を持つ。型は src/stubs を直接 import。InterfaceNames はクエリ定数のみ。UnoObject.iface でキャッシュ。
3. RawProps 統合: CharacterProperties/ParagraphProperties/TextProperties/FormulaProperties を束ねる RawProps を定義し、UnoObject.iface で各 PropertySet をキャッシュ。SheetCell.props から透過アクセスできるよう __getattr__/__setattr__ で委譲先を決定し、優先順位を決める。
4. 境界線/プロパティ: 既存 Cell のボーダー変換ロジックを抽出・共通化して _border_getter/_border_setter を実装。RawProps 経由で getPropertyValue/setPropertyValue にフォールバックし、CharHeight/BackgroundColor/NumberFormat/CharUnderline など主要プロパティを明示。
5. シートAPI追加: src/excellikeuno/table/sheet.py に sheet_cell(col,row) を追加して SheetCell を返す。Range/Sheet 内部の cell/セル取得もオプトインで SheetCell を使えるよう拡張（フラグまたは段階的切替）。
6. 段階的置換: Cell を内部で SheetCell に委譲するか、エイリアス化して既存 API を壊さず移行する。エクスポートは SheetCell を優先し、Cell は非推奨表示をコメントで案内。
7. 呼び出し側移行: tests/ や主要コードの呼び出しを sheet_cell/SheetCell へ置換。Range/Sheet が返す型の変更に伴う影響をテストで検証。
8. ドキュメント整備: agents/plan.md の SheetCell 章、agents/operation_spec.md、agents/naming_rules.md を更新し、VBAライク例を SheetCell ベースに差し替え。

Relevant files
- src/excellikeuno/table/cell.py, cell_properties.py — 既存ロジックの参照、cell_properties は利用しなくて良い
- src/excellikeuno/table/raw_props.py（新規想定） — Character/Paragraph/Text/Formula Properties を束ねる RawProps
- src/excellikeuno/table/__init__.py, sheet.py, range.py — SheetCell エクスポートと取得経路の切替
- src/excellikeuno/style/font.py, border.py — 再利用するプロキシ
- src/excellikeuno/typing/interfaces.py と src/stubs/** — 型とインターフェース定数
- tests/ 配下セル/範囲/シート関連 — 呼び出し更新と回帰検証
- agents/* docs — SheetCell 仕様反映

Verification
1. $env:PYTHONPATH='H:\LibreOffice-ExcelLike\src'; & 'C:\Program Files\LibreOffice\program\python' -m pytest tests
2. REPL で SheetCell の value/formula/props/borders/font を触るスモークチェック。

Decisions
- Cell を最終的に SheetCell へ置換する方針。移行中は Cell をエイリアス/ラッパーとして残し後方互換を確保。
- 型は src/stubs 直 import。InterfaceNames はクエリ定数のみ。
- Font/Borders の API は維持し、新規スタイル抽象は追加しない。

Further Considerations
1. Range/Sheet のデフォルト戻り値をいつ SheetCell に切り替えるか（即時/フラグ/リリース段階）。
    → SheetCell の試験作成なので、一旦保留。切り替えは即時の予定。

2. 非推奨化のサイクル（何リリースで Cell を削除するか）を決める必要あり。
    →　即時の予定

3. RawProps で属性名が衝突した際の優先順位（Character vs Paragraph vs Text vs Formula）をどう決めるか。

    → UNO API の時点で継承時に衝突しないように設計されている

4. 追加で扱うプロパティの優先順位（NumberFormat/BackColor/Underline 以外に何を先にサポートするか）を整理。

    → これは別途、既存のテストコードに合わせて作成する。
