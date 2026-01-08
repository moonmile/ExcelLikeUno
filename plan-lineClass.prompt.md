## Plan: Line クラスをFont風に追加

Line を Font/Borders と同じプロキシ構造に整備し、Shape から line.width などを安全に取得/設定できるようにする方針です。

### Steps
1. Font と同じ owner/getter/setter/buffer パターンで Line を実装し、width(=LineWidth/OuterLineWidth), color, line_style, dash系を適用するよう整理する（src/excellikeuno/style/line.py）。
2. Shape が Line プロキシを生成・再利用するようにし、既存の UNO ライン getter/setter を介して反映させる（src/excellikeuno/drawing/shape.py）。
3. LineStyle/Dash など対応プロパティの型ヒント・enum を typing に揃え、Line API と一致させる（src/excellikeuno/typing 配下）。
4. 代表的な shape を生成して line.width/color/style を往復させる回帰テストを追加する（tests 配下、既存 shape テストに倣う）。

### Further Considerations
1. どこまで対応するか: A) width+color だけ B) width+color+LineStyle まで C) dash 系含めフル対応。選択はどれにしますか？
   → C) フル対応でお願いします。

LineProperties 
- LineStyle : line.line_style
- LineDash  : line.dash
- LineDashName : line.dash_name
- LineColor : line.color
- LineTransparence : line.transparence
- LineWidth : line.width

までを実装