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
    cell = sheet.cell2(0, 0)
    
    top_color = cell.borders.top.color
    top_width = cell.borders.top.width

    cell.borders.top.color = 0x112233
    cell.borders.top.width = 20

    assert cell.borders.top.color == 0x112233
    assert cell.borders.top.width == 20

    cell.borders.top.color = top_color
    cell.borders.top.width = top_width


def test_cell_border_roundtrip_via_proxy2():
    _, _, sheet = _connect_or_skip()
    cell = sheet.cell2(0, 8)

    originals = {
        "top": cell.borders.top,
        "bottom": cell.borders.bottom,
        "left": cell.borders.left,
        "right": cell.borders.right,
    }

    top_line = BorderLine(Color=0x112233, InnerLineWidth=20, OuterLineWidth=60, LineDistance=80)
    bottom_line = BorderLine(Color=0x445566, InnerLineWidth=10, OuterLineWidth=40, LineDistance=70)
    left_line = BorderLine(Color=0x778899, InnerLineWidth=5, OuterLineWidth=30, LineDistance=50)
    right_line = BorderLine(Color=0x99AABB, InnerLineWidth=15, OuterLineWidth=45, LineDistance=65)

    cell.borders.top = top_line
    cell.borders.bottom = bottom_line
    cell.borders.left = left_line
    cell.borders.right = right_line

    assert cell.borders.top.Color == top_line.Color
    assert abs(cell.borders.bottom.OuterLineWidth - bottom_line.OuterLineWidth) <= 2
    assert cell.borders.left.Color == left_line.Color
    assert cell.borders.right.Color == right_line.Color

    cell.borders.top = originals["top"]
    cell.borders.bottom = originals["bottom"]
    cell.borders.left = originals["left"]
    cell.borders.right = originals["right"]
