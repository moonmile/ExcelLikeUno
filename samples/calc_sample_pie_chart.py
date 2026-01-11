# 円グラフのサンプル

from excellikeuno.connection import connect_calc
from excellikeuno.typing.structs import Rectangle
(desktop, doc, sheet) = connect_calc()
data_range = sheet.range("A1:B6")
sz = 100 / 5
data = [
    ["Label", "Value"],
    ["佐藤君", sz],
    ["鈴木君", sz],
    ["吉田君", sz],
    ["森田君", sz],
    ["綾波さん x3", sz*3],
]
data_range.value = data
name = "Sample_Pie_Chart"
chart = sheet.charts.add_pie_diagram(
    name=name,
    data_range=data_range,
    rectangle=Rectangle(5000, 5000, 10000, 8000),
    column_headers=True,
    row_headers=True,
)
# 凡例の表示
chart.has_legend = True

chart_doc = chart.raw.getEmbeddedObject()
# タイトルの表示フラグ
chart_doc.setPropertyValue("HasMainTitle", True)
chart_doc.setPropertyValue("HasSubTitle", True)

# タイトル文字列
title = chart_doc.getTitle()
title.String = "メインタイトル"

sub = chart_doc.getSubTitle()
sub.String = "サブタイトル"

print("作成したチャート名:", chart.name)
