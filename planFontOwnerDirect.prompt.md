## Plan: Font で owner 直参照＋Range は現行維持

Font が owner の文字プロパティに直接アクセスし、Range だけは既存ブロードキャストを使う設計に整理する。

### Steps
1. owner 付き Font の分岐設計を定義し、共通インタフェース（`text_properties` or `character_properties`）で読む/書くキーを列挙して font.py に反映する。
2. Cell/Shape に「Font が参照可能な text/character プロパティ取得口」（既存の character_properties/text_properties を活用）を明示し、Range は `_font_broadcast` のままにする。
3. Font.apply で owner が Range なら既存 setter を使い、Cell/Shape なら直接 Char* を設定する経路を追加し、getter も同様に owner から読めるよう実装方針をまとめる。
4. テスト追加/更新案を用意する（Cell/Shape への直接適用、owner なしバッファ、Range ブロードキャストの回帰） tests/test_cell_font.py / tests/test_range_font.py。

### Further Considerations
1. backcolor の優先順位（CharBackColor vs CellBackColor）とキャッシュ挙動をどこで維持するか（Cell/Shape 側 or Font 側）。
   → Font 側で維持する
