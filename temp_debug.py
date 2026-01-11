from excellikeuno.connection import connect_calc
from excellikeuno.typing.structs import Rectangle
from excellikeuno.table.range import Range
import traceback


desktop, doc, sheet = connect_calc()
charts = sheet.raw.getCharts()
data_range = sheet.range("A1:B6")
addr_from_wrapper = data_range.raw.getRangeAddress()
print("addr from wrapper type:", type(addr_from_wrapper))
print("addr fields:", addr_from_wrapper.Sheet, addr_from_wrapper.StartColumn, addr_from_wrapper.StartRow, addr_from_wrapper.EndColumn, addr_from_wrapper.EndRow)
rect = Rectangle(5000, 5000, 10000, 8000).to_raw()

name = "__debug_chart__"
try:
    if charts.hasByName(name):
        charts.removeByName(name)
except Exception:
    pass

try:
    charts.addNewByName(name, rect, (addr_from_wrapper,), True, True)
    print("addNewByName direct success")
except Exception as exc:
    print("direct failed", type(exc), exc)
    try:
        import uno
        addr_seq = uno.Any("[]com.sun.star.table.CellRangeAddress", (addr_from_wrapper,))
        charts.addNewByName(name, rect, addr_seq, True, True)
        print("addNewByName with Any success")
    except Exception as exc2:
        print("with Any failed", type(exc2), exc2)
        try:
            import uno
            uno.invoke(charts, "addNewByName", (name, rect, (addr_from_wrapper,), True, True))
            print("invoke success")
        except Exception as exc3:
            print("invoke failed", type(exc3), exc3)
            traceback.print_exc()

from excellikeuno.table.chart import ChartCollection
col = ChartCollection(sheet)
try:
    chart = col.add_pie_diagram("__debug_chart2__", data_range, rectangle=Rectangle(6000, 6000, 8000, 6000))
    print("ChartCollection.add_pie_diagram success, name:", chart.name)
except Exception as exc:
    print("ChartCollection.add_pie_diagram failed", type(exc), exc)
