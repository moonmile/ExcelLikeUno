import pytest

from excellikeuno.connection import connect_calc
from excellikeuno.style.border import Borders, BorderStyle
from excellikeuno.typing import BorderLine


def _connect_or_skip():
    try:
        return connect_calc()
    except RuntimeError as exc:  # pragma: no cover - environment dependent
        pytest.skip(f"UNO runtime not available: {exc}")


def test_cell_border_roundtrip_via_proxy():
    _, _, sheet = _connect_or_skip()
    cell = sheet.cell(0, 8)

    originals = {
        "top": cell.props.TopBorder,
        "bottom": cell.props.BottomBorder,
        "left": cell.props.LeftBorder,
        "right": cell.props.RightBorder,
    }

    top_line = BorderLine(Color=0x112233, InnerLineWidth=20, OuterLineWidth=60, LineDistance=80)
    bottom_line = BorderLine(Color=0x445566, InnerLineWidth=10, OuterLineWidth=40, LineDistance=70)
    left_line = BorderLine(Color=0x778899, InnerLineWidth=5, OuterLineWidth=30, LineDistance=50)
    right_line = BorderLine(Color=0x99AABB, InnerLineWidth=15, OuterLineWidth=45, LineDistance=65)

    try:
        cell.borders.top = top_line
        cell.borders.bottom = bottom_line
        cell.borders.left = left_line
        cell.borders.right = right_line

        assert getattr(cell.props.TopBorder, "Color", None) == top_line.Color
        assert abs(getattr(cell.props.BottomBorder, "OuterLineWidth", 0) - bottom_line.OuterLineWidth) <= 2
        assert getattr(cell.props.LeftBorder, "Color", None) == left_line.Color
        assert getattr(cell.props.RightBorder, "Color", None) == right_line.Color
    finally:
        cell.props.TopBorder = originals["top"]
        cell.props.BottomBorder = originals["bottom"]
        cell.props.LeftBorder = originals["left"]
        cell.props.RightBorder = originals["right"]


def test_borderstyle_config_constructor_defaults():
    style = BorderStyle(color=0x010203, weight=22, line_style=3, inner_width=4, distance=6)

    assert style.color == 0x010203
    assert style.weight == 22
    assert style.inner_width == 4
    assert style.distance == 6
    assert style.line_style == 3


def test_border_config_holder_reuse_on_cells():
    _, _, sheet = _connect_or_skip()
    first = sheet.cell(1, 8)
    second = sheet.cell(1, 9)

    originals = [
        (first.props.TopBorder, first.props.BottomBorder, first.props.LeftBorder, first.props.RightBorder),
        (second.props.TopBorder, second.props.BottomBorder, second.props.LeftBorder, second.props.RightBorder),
    ]

    top_line = BorderLine(Color=0x334455, InnerLineWidth=10, OuterLineWidth=40, LineDistance=60)
    bottom_line = BorderLine(Color=0x556677, InnerLineWidth=15, OuterLineWidth=35, LineDistance=55)
    left_line = BorderLine(Color=0x778899, InnerLineWidth=5, OuterLineWidth=25, LineDistance=45)
    right_line = BorderLine(Color=0x99AABB, InnerLineWidth=8, OuterLineWidth=28, LineDistance=48)

    border_cfg = Borders(top=top_line, bottom=bottom_line, left=left_line, right=right_line)

    try:
        first.borders = border_cfg
        second.borders = border_cfg
    
        for cell in (first, second):
            assert getattr(cell.props.TopBorder, "Color", None) == top_line.Color
            assert getattr(cell.props.BottomBorder, "Color", None) == bottom_line.Color
            assert getattr(cell.props.LeftBorder, "Color", None) == left_line.Color
            assert getattr(cell.props.RightBorder, "Color", None) == right_line.Color

        assert border_cfg.top.Color == top_line.Color
        assert border_cfg.bottom.Color == bottom_line.Color
    finally:
        (f_top, f_bottom, f_left, f_right), (s_top, s_bottom, s_left, s_right) = originals
        first.props.TopBorder = f_top
        first.props.BottomBorder = f_bottom
        first.props.LeftBorder = f_left
        first.props.RightBorder = f_right

        second.props.TopBorder = s_top
        second.props.BottomBorder = s_bottom
        second.props.LeftBorder = s_left
        second.props.RightBorder = s_right


def test_border_all_sets_all_sides():
    _, _, sheet = _connect_or_skip()
    cell = sheet.cell(2, 8)

    originals = {
        "top": cell.props.TopBorder,
        "bottom": cell.props.BottomBorder,
        "left": cell.props.LeftBorder,
        "right": cell.props.RightBorder,
    }

    line = BorderLine(Color=0xABCDEF, InnerLineWidth=12, OuterLineWidth=24, LineDistance=36)

    try:
        cell.borders.all = line

        assert getattr(cell.props.TopBorder, "Color", None) == line.Color
        assert getattr(cell.props.BottomBorder, "Color", None) == line.Color
        assert getattr(cell.props.LeftBorder, "Color", None) == line.Color
        assert getattr(cell.props.RightBorder, "Color", None) == line.Color
    finally:
        cell.props.TopBorder = originals["top"]
        cell.props.BottomBorder = originals["bottom"]
        cell.props.LeftBorder = originals["left"]
        cell.props.RightBorder = originals["right"]


def test_cell_borderstyle_attribute_updates():
    _, _, sheet = _connect_or_skip()
    cell = sheet.cell(3, 8)

    originals = {
        "top": cell.props.TopBorder,
        "bottom": cell.props.BottomBorder,
        "left": cell.props.LeftBorder,
        "right": cell.props.RightBorder,
    }

    try:
        cell.borders.top.color = 0x123456
        cell.borders.top.weight = 40
        cell.borders.top.line_style = 1

        assert cell.borders.top.color == 0x123456
        assert abs(cell.borders.top.weight - 40) <= 2
        assert int(cell.borders.top.line_style) == 1
    finally:
        cell.props.TopBorder = originals["top"]
        cell.props.BottomBorder = originals["bottom"]
        cell.props.LeftBorder = originals["left"]
        cell.props.RightBorder = originals["right"]



def test_cell_borderstyle_constructor():
    _, _, sheet = _connect_or_skip()
    cell = sheet.cell(3, 8)

    originals = {
        "top": cell.props.TopBorder,
        "bottom": cell.props.BottomBorder,
        "left": cell.props.LeftBorder,
        "right": cell.props.RightBorder,
    }

    try:

        borderstyle = BorderStyle(color=0x123456, weight=40, line_style=0)
        cell.borders.top = borderstyle

        assert cell.borders.top.color == 0x123456
        assert abs(cell.borders.top.weight - 40) <= 2
        assert int(cell.borders.top.line_style) == 0
    finally:
        cell.props.TopBorder = originals["top"]
        cell.props.BottomBorder = originals["bottom"]
        cell.props.LeftBorder = originals["left"]
        cell.props.RightBorder = originals["right"]
