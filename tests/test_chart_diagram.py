import pytest

from excellikeuno.connection import connect_calc
from excellikeuno.typing import InterfaceNames
from excellikeuno.typing.structs import Rectangle


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


def test_add_pie_diagram_sets_range_and_name():
    _, _, sheet = _connect_or_skip()
    data_range = _seed_sample_data(sheet)

    name = "Chart_Pie_Test"
    charts = sheet.charts
    try:
        charts.remove(name)
    except Exception:
        pass

    rect = Rectangle(1500, 1500, 7000, 7000)
    chart = charts.add_pie_diagram(
        name=name,
        data_range=data_range,
        rectangle=rect,
        column_headers=True,
        row_headers=True,
    )

    try:
        assert chart.name == name
        addr_expected = data_range.iface(InterfaceNames.X_CELL_RANGE_ADDRESSABLE).getRangeAddress()
        addr_actual = chart.range
        assert getattr(addr_actual, "StartColumn", None) == addr_expected.StartColumn
        assert getattr(addr_actual, "EndRow", None) == addr_expected.EndRow
    finally:
        try:
            charts.remove(name)
            pass
        except Exception:
            pass


def test_add_pie_diagram_toggle_legend():
    _, _, sheet = _connect_or_skip()
    data_range = _seed_sample_data(sheet)

    name = "Chart_Pie_Legend"
    charts = sheet.charts
    try:
        charts.remove(name)
    except Exception:
        pass

    rect = Rectangle(2000, 2000, 6000, 6000)
    chart = charts.add_pie_diagram(
        name=name,
        data_range=data_range,
        rectangle=rect,
        column_headers=True,
        row_headers=True,
    )

    try:
        chart.has_legend = True
        assert chart.has_legend is True
        chart.has_legend = False
        assert chart.has_legend is False
    finally:
        try:
            charts.remove(name)
        except Exception:
            pass


def test_toggle_titles():
    _, _, sheet = _connect_or_skip()
    data_range = _seed_sample_data(sheet)

    name = "Chart_Title_Test"
    charts = sheet.charts
    try:
        charts.remove(name)
    except Exception:
        pass

    rect = Rectangle(1800, 1800, 6000, 5000)
    chart = charts.add_pie_diagram(
        name=name,
        data_range=data_range,
        rectangle=rect,
        column_headers=True,
        row_headers=True,
    )

    try:
        chart.has_main_title = True
        chart.has_sub_title = True
        assert chart.has_main_title is True
        assert chart.has_sub_title is True

        chart.has_main_title = False
        chart.has_sub_title = False
        assert chart.has_main_title is False
        assert chart.has_sub_title is False

        chart.title = "Main"
        chart.sub_title = "Sub"
        assert chart.title == "Main"
        assert chart.sub_title == "Sub"
        assert chart.has_main_title is True
        assert chart.has_sub_title is True

        chart.title = ""
        chart.sub_title = "  "
        assert chart.title == ""
        assert chart.sub_title == ""
        assert chart.has_main_title in (False, None)
        assert chart.has_sub_title in (False, None)
    finally:
        try:
            charts.remove(name)
        except Exception:
            pass


def test_remove_is_idempotent_and_checks_existence():
    _, _, sheet = _connect_or_skip()
    data_range = _seed_sample_data(sheet)

    name = "Chart_Remove_Test"
    charts = sheet.charts
    try:
        charts.remove(name)
    except Exception:
        pass

    rect = Rectangle(1200, 1200, 5000, 4000)
    chart = charts.add_pie_diagram(
        name=name,
        data_range=data_range,
        rectangle=rect,
        column_headers=True,
        row_headers=True,
    )

    try:
        assert charts.exists(name) is True
        charts.remove(name)
        assert charts.exists(name) is False
        # Second removal should be a no-op (no exception)
        charts.remove(name)
    finally:
        try:
            charts.remove(name)
        except Exception:
            pass
