import pytest

from excellikeuno.connection import connect_calc
from excellikeuno.style.border import Border
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
        "top": cell.TopBorder,
        "bottom": cell.BottomBorder,
        "left": cell.LeftBorder,
        "right": cell.RightBorder,
    }

    top_line = BorderLine(Color=0x112233, InnerLineWidth=20, OuterLineWidth=60, LineDistance=80)
    bottom_line = BorderLine(Color=0x445566, InnerLineWidth=10, OuterLineWidth=40, LineDistance=70)
    left_line = BorderLine(Color=0x778899, InnerLineWidth=5, OuterLineWidth=30, LineDistance=50)
    right_line = BorderLine(Color=0x99AABB, InnerLineWidth=15, OuterLineWidth=45, LineDistance=65)

    try:
        cell.border.top = top_line
        cell.border.bottom = bottom_line
        cell.border.left = left_line
        cell.border.right = right_line

        assert getattr(cell.TopBorder, "Color", None) == top_line.Color
        assert abs(getattr(cell.BottomBorder, "OuterLineWidth", 0) - bottom_line.OuterLineWidth) <= 2
        assert getattr(cell.LeftBorder, "Color", None) == left_line.Color
        assert getattr(cell.RightBorder, "Color", None) == right_line.Color
    finally:
        cell.TopBorder = originals["top"]
        cell.BottomBorder = originals["bottom"]
        cell.LeftBorder = originals["left"]
        cell.RightBorder = originals["right"]


def test_border_config_holder_reuse_on_cells():
    _, _, sheet = _connect_or_skip()
    first = sheet.cell(1, 8)
    second = sheet.cell(1, 9)

    originals = [
        (first.TopBorder, first.BottomBorder, first.LeftBorder, first.RightBorder),
        (second.TopBorder, second.BottomBorder, second.LeftBorder, second.RightBorder),
    ]

    top_line = BorderLine(Color=0x334455, InnerLineWidth=10, OuterLineWidth=40, LineDistance=60)
    bottom_line = BorderLine(Color=0x556677, InnerLineWidth=15, OuterLineWidth=35, LineDistance=55)
    left_line = BorderLine(Color=0x778899, InnerLineWidth=5, OuterLineWidth=25, LineDistance=45)
    right_line = BorderLine(Color=0x99AABB, InnerLineWidth=8, OuterLineWidth=28, LineDistance=48)

    border_cfg = Border(top=top_line, bottom=bottom_line, left=left_line, right=right_line)

    try:
        first.border = border_cfg
        second.border = border_cfg

        for cell in (first, second):
            assert getattr(cell.TopBorder, "Color", None) == top_line.Color
            assert getattr(cell.BottomBorder, "Color", None) == bottom_line.Color
            assert getattr(cell.LeftBorder, "Color", None) == left_line.Color
            assert getattr(cell.RightBorder, "Color", None) == right_line.Color

        assert border_cfg.top.Color == top_line.Color
        assert border_cfg.bottom.Color == bottom_line.Color
    finally:
        (f_top, f_bottom, f_left, f_right), (s_top, s_bottom, s_left, s_right) = originals
        first.TopBorder = f_top
        first.BottomBorder = f_bottom
        first.LeftBorder = f_left
        first.RightBorder = f_right

        second.TopBorder = s_top
        second.BottomBorder = s_bottom
        second.LeftBorder = s_left
        second.RightBorder = s_right


def test_border_all_sets_all_sides():
    _, _, sheet = _connect_or_skip()
    cell = sheet.cell(2, 8)

    originals = {
        "top": cell.TopBorder,
        "bottom": cell.BottomBorder,
        "left": cell.LeftBorder,
        "right": cell.RightBorder,
    }

    line = BorderLine(Color=0xABCDEF, InnerLineWidth=12, OuterLineWidth=24, LineDistance=36)

    try:
        cell.border.all = line

        assert getattr(cell.TopBorder, "Color", None) == line.Color
        assert getattr(cell.BottomBorder, "Color", None) == line.Color
        assert getattr(cell.LeftBorder, "Color", None) == line.Color
        assert getattr(cell.RightBorder, "Color", None) == line.Color
    finally:
        cell.TopBorder = originals["top"]
        cell.BottomBorder = originals["bottom"]
        cell.LeftBorder = originals["left"]
        cell.RightBorder = originals["right"]


def test_range_border_broadcast_via_proxy():
    _, _, sheet = _connect_or_skip()
    rng = sheet.range(3, 8, 4, 9)  # 2x2 block

    cells = [rng.cell(0, 0), rng.cell(1, 0), rng.cell(0, 1), rng.cell(1, 1)]
    originals = [
        (c.TopBorder, c.BottomBorder, c.LeftBorder, c.RightBorder) for c in cells
    ]

    line = BorderLine(Color=0xC0FFEE, InnerLineWidth=18, OuterLineWidth=36, LineDistance=54)

    try:
        rng.border = Border(all=line)

        for cell in cells:
            assert getattr(cell.TopBorder, "Color", None) == line.Color
            assert abs(getattr(cell.BottomBorder, "OuterLineWidth", 0) - line.OuterLineWidth) <= 2
            assert getattr(cell.LeftBorder, "Color", None) == line.Color
            assert getattr(cell.RightBorder, "Color", None) == line.Color
    finally:
        for cell, (top, bottom, left, right) in zip(cells, originals):
            cell.TopBorder = top
            cell.BottomBorder = bottom
            cell.LeftBorder = left
            cell.RightBorder = right
