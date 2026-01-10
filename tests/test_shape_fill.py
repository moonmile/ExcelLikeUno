import pytest

from excellikeuno.connection import connect_calc
from excellikeuno.typing.calc import FillStyle, GradientStyle, HatchStyle
from excellikeuno.typing.structs import Gradient, Hatch


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


def test_shape_fill_color_sets_style_solid():
    _, doc, sheet, Point, Size = _connect_or_skip()
    rect, draw_page = _add_rectangle(doc, sheet, Point, Size)
    try:
        shape = sheet.shapes()[-1]
        shape.fill.color = 0x00FF00
        assert shape.fill.color == 0x00FF00
        assert FillStyle(shape.FillStyle) == FillStyle.SOLID
        assert shape.FillColor == 0x00FF00
    finally:
        draw_page.remove(rect)


def test_shape_fill_gradient_sets_style_gradient():
    _, doc, sheet, Point, Size = _connect_or_skip()
    rect, draw_page = _add_rectangle(doc, sheet, Point, Size)
    try:
        shape = sheet.shapes()[-1]
        grad = Gradient(
            Style=GradientStyle.LINEAR,
            StartColor=0xFF0000,
            EndColor=0x0000FF,
            Angle=0,
            Border=0,
            XOffset=0,
            YOffset=0,
            StartIntensity=100,
            EndIntensity=100,
            StepCount=0,
        )
        shape.fill.gradient = grad
        assert FillStyle(shape.FillStyle) == FillStyle.GRADIENT
    finally:
        draw_page.remove(rect)


def test_shape_fill_hatch_sets_style_hatch():
    _, doc, sheet, Point, Size = _connect_or_skip()
    rect, draw_page = _add_rectangle(doc, sheet, Point, Size)
    try:
        shape = sheet.shapes()[-1]
        hatch = Hatch(Style=HatchStyle.SINGLE, Color=0x000000, Distance=100, Angle=450)
        shape.fill.hatch = hatch
        assert FillStyle(shape.FillStyle) == FillStyle.HATCH
    finally:
        draw_page.remove(rect)
