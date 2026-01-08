## Plan: BorderStyleで罫線を属性指定

### Steps
1. border.py に `BorderStyle` を追加し、side名＋親 `Borders` を保持。`color/inner_width/outer_width/distance/line_style/width(weight)` プロパティで現行ラインをクローン→フィールド更新→親に適用。
2. `Borders` の `top/bottom/left/right/all` を `BorderStyle` を返すようにし、従来のライン構造体取得用に `top_line` などの後方互換アクセサを用意。
3. `Borders.apply` を拡張し、`BorderStyle` からの部分更新（color/width/line_style 等）を `BorderLine2` に正規化。従来の全ライン/辞書入力は維持。
4. テスト追加: `test_cell_border.py` に per-side 属性更新と range ブロードキャスト (`rng.borders.top.color` など) を追加し、既存テストも通ることを確認。
5. サンプル更新: `samples/calc_sample.py` に `cell.borders.top.color/weight/line_style` と `BorderStyle(...)` 代入例を記載し、LineStyle=0 が実線であることをコメント。

### Further Considerations
- 後方互換を維持（生の `BorderLine/BorderLine2` 取得・設定経路を残す）。
  → 残す。BoderLine 系は TopBorder 等の UNO 直結系で利用される。
- 新規 setter は `BorderLine2` 優先で LineStyle/LineWidth を保持し、`weight/width` のマッピングを一貫させる。
  → そうする
- UNO 幅丸めは ±2 程度の許容をテストに入れる。
  → そうする
