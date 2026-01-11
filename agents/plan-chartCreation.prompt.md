## Plan: Chart 作成・更新 API

Sheet.charts に Chart/ChartCollection を追加し、作成時は add(name, range, pos/size?)、更新時は chart.range/position/size で再設定できるようにする。既存 Point/Size セッターの形を流用し、Rectangle ヘルパーを追加する。

### Steps
1) UNO 型定義拡張: XTableCharts/XTableChart 等を typing/interfaces.py に追加し、Protocol を typing/calc.py に追加。StructNames に Rectangle を新設。
2) Rectangle dataclass (awt.Rectangle) を typing/structs.py に追加し、Point/Size と同様の to_uno_struct と fallback を実装。
3) Chart/ChartCollection 実装: table/chart.py を新設し、Base 継承で name/range/position/size/props を XTableChart + XPropertySet に委譲。range は CellRangeAddress/tuple を受け、position/size は Rectangle/Point/Size 互換で設定可能に。
4) Sheet.charts アクセサと作成 API: sheet.py に charts プロパティを追加。add(name, range, rect_or_pos/size, headers flags) で XTableCharts.addNewByName を呼び、ChartDocument を BarDiagram 等に差し替える汎用/バー用メソッドを用意。
5) エクスポート・ドキュメント更新: table/__init__.py, __init__.py, agents/class_design.md, agents/folder_structure.md, agents/ToDo.md を同期。
6) テスト追加: UNO 環境前提で add/add_bar_diagram を使い、作成後に chart.range/position/size を更新して反映を検証する pytest を tests に追加（UNO 不可なら skip）。

### Further Considerations
1) add の座標指定: Rectangle 一括指定 vs pos+size 分割の両方をサポートし、優先順位・型受け入れルールを決める。
2) 名前空間を excellikeuno/chart/ にする


### 実装例

- sheet.charts.add_area_chart( ... )
- sheet.charts.add_bar_diagram( ... )
- sheet.charts.add_bubble_diagram( ... )
- sheet.charts.add_dim2d_diagram( ... )
- sheet.charts.add_donut_diagram( ... )
- sheet.charts.add_filled_net_diagram( ... )
- sheet.charts.add_line_diagram( ... )
- sheet.charts.add_net_diagram( ... )
- sheet.charts.add_pie_diagram( ... )
- sheet.charts.add_stackable_diagram( ... )
- sheet.charts.add_stock_diagram( ... )
- sheet.charts.add_xy_diagram( ... )

汎用版
- sheeet.charts.add(x: int, y: int, width: int, height: int, chart_type: str) -> Chart
- sheeet.charts.add(pos: Position, size: Size, chart_type: str) -> Chart
- sheeet.charts.add(rect: Rectangle, chart_type: str) -> Chart

```python
chart = sheet.charts.add(x: 0, y: 0, width: 10000, height: 7000, chart_type: "bar")
chart.name = "Sample Chart"
chart.range = sheet.range("A1:B10")
chart.title = "Sample Chart"
```

