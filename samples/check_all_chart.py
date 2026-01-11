# doc に含まれるチャートの一覧を取得する

from excellikeuno.connection import connect_calc
(desktop, doc, sheet) = connect_calc()
charts = sheet.charts
print("シートに含まれるチャート一覧:")
for chart_name in charts._names():
    print(" -", chart_name)
print("チャート数:", len(charts))
for chart_name in charts._names():
    chart = charts[chart_name]
    print(f"チャート名: {chart.name}, 種類: {chart.chart_type}")