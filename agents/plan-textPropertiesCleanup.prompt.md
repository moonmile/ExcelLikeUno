# Plan: Text/Character Properties整理

TextProperties を Text* 専用に絞り、Char* は CharacterProperties に一本化。Shape からの利用経路とテストを更新する。

## Steps

1. TextProperties から Char* アクセサを削除し、Text* とテキストパディング/オートグローのみ残す（text_properties.py）。
2. Shape に CharacterProperties ラッパー入口を追加し、既存の sub-wrapper (line/fill/shadow) に倣う（shape.py）。
3. Shape の文字属性アクセスを CharacterProperties 経由に付け替え、Justify は TextHorizontalAdjust/TextVerticalAdjust を優先し最小限のフォールバックに整理。
4. Shape 向けテストを追加/更新し、Char* が CharacterProperties 経由で効くことと Justify 振る舞いを検証（test_shape.py など）。
5. 設計ノートを同期し、TextProperties 範囲縮小と CharacterProperties 依存を明記（plan-textPropertiesCleanup.prompt.md ほか関連ガイド）。


## Further Considerations

1. 後方互換: Shape 上の Char* エイリアスを一時的に残すか即削除か？
   → Shape 側の Char* エイリアスを一時的に残す
2. Justify 互換: ParaAdjust/HoriJustify などのフォールバックをどこまで保持するか？
   → 最小限のフォールバックに整理し、主要な新プロパティを優先する
3. テスト戦略: headless UNO 実機 vs モックで CharacterProperties をどこまでカバーするか？
   → headless UNO 実機での統合テストを重視し、モックは補助的に利用する
   
Shape クラスに font プロパティを付ける。
Font クラスがアクセスする CharacterProperties を Cell.font と同じように Shape.font で公開する

- CharFontName : shape.font.name
- CharHeight : shape.font.size
- CharWeight: shape.font.weight
- CharPosture: shape.font.itralic
- CharUnderline: shape.font.underline
- CharStrikeout: shape.font.strikeout
- CharColor: shape.font.color
- CharBackColor: shape.font.back_color
- CharEscapement: shape.font.subscript/superscript/strikethrough