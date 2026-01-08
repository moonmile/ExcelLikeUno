import pytest

from excellikeuno.connection import connect_calc
from excellikeuno.typing import LineStyle


def _connect_or_skip():
    try:
        from com.sun.star.awt import Point, Size  # type: ignore
    except ImportError as exc:  # pragma: no cover - depends on LibreOffice runtime
        pytest.skip(f"UNO runtime not available: {exc}")

    try:
        desktop, doc, sheet = connect_calc()
    except RuntimeError as exc:  # pragma: no cover - depends on LibreOffice runtime
        pytest.skip(f"UNO runtime not available: {exc}")

    return desktop, doc, sheet, Point, Size


def _add_rectangle(doc, sheet, Point, Size):
    draw_page = sheet.raw.getDrawPage()
    rect = doc.createInstance("com.sun.star.drawing.RectangleShape")
    rect.setPosition(Point(1000, 1000))
    rect.setSize(Size(2000, 1200))
    draw_page.add(rect)
    return rect, draw_page


def test_shape_line_roundtrip():
    _, doc, sheet, Point, Size = _connect_or_skip()
    rect, draw_page = _add_rectangle(doc, sheet, Point, Size)
    try:
        shape = sheet.shapes()[-1]
        line = shape.line

        dash = None
        try:
            import uno  # type: ignore

            dash = uno.createUnoStruct("com.sun.star.drawing.LineDash")
            dash.Style = 2
            dash.Dots = 0
            dash.DotLen = 0
            dash.Dashes = 2
            dash.DashLen = 100
            dash.Distance = 100
        except Exception:
            pass

        new_color = 0x123456
        new_width = 200
        new_style = LineStyle.DASH
        new_trans = 10

        line.color = new_color
        line.width = new_width
        line.line_style = new_style
        line.transparence = new_trans
        if dash is not None:
            line.dash = dash
        line.dash_name = "TestDash"

        assert shape.line.color == new_color
        assert shape.line.width == new_width
        assert shape.line.line_style == new_style
        assert shape.line.transparence == new_trans
        if dash is not None:
            assert getattr(shape.line.dash, "DashLen", None) == getattr(dash, "DashLen", None)
        assert shape.line.dash_name == "TestDash"
    finally:
        try:
            draw_page.remove(rect)
        except Exception:
            pass
