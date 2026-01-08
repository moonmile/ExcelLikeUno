## Plan: Border プロキシ追加（Font と同様パターン）

### Steps
1. `style/border.py`（新規）に `Border` プロキシを実装：`owner` 付き/なしで動作、`_current` 相当は最小アクセス、`top/left/right/bottom` を `BorderLine` で get/set。
2. Cell に `@property border` を追加し、`Border(owner=self)` を返す。setter で owner なし Border からのバッファも適用できるようにする。
3. Range には `border` を追加し、代表セル getter ＋ `_border_broadcast`（全セルに適用）で対応。
4. Shape はスキップ（要求が Cell なので）。必要なら後続対応。
5. テスト追加: `tests/test_cell_border.py` を作成し、Top/Bottom/Left/Right の roundtrip と、owner なし Border 再利用を確認。必要に応じ Range ブロードキャストも追加。

### Further Considerations
1. `BorderLine`/`BorderLine2` のどちらをサポートするか（とりあえず `BorderLine`）。
   → BorderLine で実装


- Cell.border プロキシを追加し、Top/Bottom/Left/Right ボーダーの get/set をサポートする。
- Range.border も追加し、左上の代表セル border を取得し、setter で全セルに適用できるようにする。
