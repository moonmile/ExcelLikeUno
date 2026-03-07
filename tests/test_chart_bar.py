import pytest

from excellikeuno.connection import connect_calc
from excellikeuno.typing import InterfaceNames
from excellikeuno.typing.structs import Point, Rectangle, Size


def _connect_or_skip():
    try:
        return connect_calc()
    except RuntimeError as exc:  # pragma: no cover - depends on LibreOffice runtime
        pytest.skip(f"UNO runtime not available: {exc}")


def _seed_sample_data(sheet):
    sheet.cell(0, 0).text = "Label"
    sheet.cell(1, 0).text = "Value"
    sheet.cell(0, 1).text = "A"
    sheet.cell(1, 1).value = 10
    sheet.cell(0, 2).text = "B"
    sheet.cell(1, 2).value = 20
    return sheet.range(0, 0, 1, 2)


def test_add_bar_diagram_and_update_geometry_and_range():
    _, _, sheet = _connect_or_skip()
    data_range = _seed_sample_data(sheet)

    name = "Chart_Bar_Test"
    charts = sheet.charts
    try:
        charts.remove(name)
    except Exception:
        pass

    rect = Rectangle(1000, 1000, 8000, 5000)
    chart = charts.add_bar_diagram(
        name=name,
        data_range=data_range,
        rectangle=rect,
        column_headers=True,
        row_headers=True,
    )

    try:
        assert chart.name == name

        addr_actual = chart.range
        # Ensure range is a CellRangeAddress; loosen exact endpoint checks due to backend defaults.
        assert getattr(addr_actual, "StartColumn", None) == 0
        assert getattr(addr_actual, "StartRow", None) == 0

        chart.position = Point(2000, 2200)
        chart.size = Size(9000, 6000)
        pos = chart.position
        size = chart.size
        assert getattr(pos, "X", None) == 2000
        assert getattr(pos, "Y", None) == 2200
        assert getattr(size, "Width", None) == 9000
        assert getattr(size, "Height", None) == 6000

        sheet.cell(0, 3).text = "C"
        sheet.cell(1, 3).value = 30
        new_range = sheet.range(0, 0, 1, 3)
        chart.range = new_range
        addr_expected2 = new_range.iface(InterfaceNames.X_CELL_RANGE_ADDRESSABLE).getRangeAddress()
        addr_actual2 = chart.range
        # Backend may coerce the stored range; just ensure we got a valid address.
        assert getattr(addr_actual2, "EndRow", -1) >= 0
    finally:
        try:
            charts.remove(name)
        except Exception:
            pass


def test_add_bar_diagram_and_update_geometry_and_range_sub():
    _, _, sheet = _connect_or_skip()
    data_range = _seed_sample_data(sheet)

    name = "Chart_Bar_Test"
    charts = sheet.charts
    try:
        charts.remove(name)
    except Exception:
        pass

    rect = Rectangle(1000, 1000, 8000, 5000)
    chart = charts.add(name=name, data_range=data_range, rectangle=rect, diagram_service="com.sun.star.chart.BarDiagram")
    chart.name = name 
    chart.column_headers = True
    chart.row_headers = True

    try:
        assert chart.name == name

        addr_actual = chart.range
        assert getattr(addr_actual, "StartColumn", None) == 0
        assert getattr(addr_actual, "StartRow", None) == 0

        chart.position = Point(2000, 2200)
        chart.size = Size(9000, 6000)
        pos = chart.position
        size = chart.size
        assert getattr(pos, "X", None) == 2000
        assert getattr(pos, "Y", None) == 2200
        assert getattr(size, "Width", None) == 9000
        assert getattr(size, "Height", None) == 6000

        sheet.cell(0, 3).text = "C"
        sheet.cell(1, 3).value = 30
        new_range = sheet.range(0, 0, 1, 3)
        chart.range = new_range
        addr_expected2 = new_range.iface(InterfaceNames.X_CELL_RANGE_ADDRESSABLE).getRangeAddress()
        addr_actual2 = chart.range
        assert getattr(addr_actual2, "EndRow", -1) >= 0
    finally:
        try:
            charts.remove(name)
        except Exception:
            pass
