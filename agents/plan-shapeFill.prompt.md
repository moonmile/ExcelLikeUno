## Plan: Shape.fill 実装

Shape の塗りつぶしを Color/Gradient/Bitmap/Hatch で扱えるようにし、既存の line/font と同じプロキシ API にまとめる。

### Steps
1) Shape.fill と共有 FillProperties を確認し、shape→fill プロキシの責務を決める（getter/setter 接続）。参照: src/excellikeuno/drawing/shape.py, src/excellikeuno/drawing/fill_properties.py。
2) UNO 型定義を拡充（FillStyle の NONE/SOLID/GRADIENT/BITMAP/HATCH、必要な bitmap mode/struct/gradient/hatch struct）して IDE 補完を整える。参照: src/excellikeuno/typing/calc.py, src/excellikeuno/typing/structs.py。
3) FillProperties を拡張し、FillStyle 切替と FillColor/FillGradient(名/struct)/FillBitmap(名/URL/Mode/Offset/Position/Size)/FillHatch、背景・透過系を正規化しつつ公開する。
4) Shape.fill を拡張した FillProperties に接続し、setter でスタイルを自動設定（SOLID/GRADIENT/BITMAP/HATCH/NONE）。必要なら defaults をバッファする。
5) ドキュメントとテストを追加（スモークでも可）。ToDo の該当項目を完了扱いに更新し、簡単な使用例または pytest を用意。

### Further Considerations
- FillBitmapURL もサポートするか？ Option A: 今回含める / Option B: 名前指定のみで後回し。
  → Option B: 名前指定のみで後回し


以下の形式でアクセスする。

Shape.fill.color
Shape.fill.gradient
Shape.fill.bitmap
Shape.fill.hatch
