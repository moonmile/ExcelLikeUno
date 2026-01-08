from excellikeuno.typing.calc import BorderLineStyle
import pytest

from excellikeuno.connection import connect_calc
from excellikeuno.style.border import Borders, BorderStyle
from excellikeuno.typing import BorderLine


def _connect_or_skip():
    try:
        return connect_calc()
    except RuntimeError as exc:  # pragma: no cover - environment dependent
        pytest.skip(f"UNO runtime not available: {exc}")


def test_range_border_broadcast_via_proxy():
    _, _, sheet = _connect_or_skip()
    rng = sheet.range(3, 8, 4, 9)  # 2x2 block

    cells = [rng.cell(0, 0), rng.cell(1, 0), rng.cell(0, 1), rng.cell(1, 1)]
    originals = [
        (c.TopBorder, c.BottomBorder, c.LeftBorder, c.RightBorder) for c in cells
    ]

    line = BorderLine(Color=0xC0FFEE, InnerLineWidth=18, OuterLineWidth=36, LineDistance=54)

    try:
        rng.borders = Borders(all=line)

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


def test_range_borderstyle_broadcast_fields():
    _, _, sheet = _connect_or_skip()
    rng = sheet.range(5, 8, 6, 9)

    cells = [rng.cell(0, 0), rng.cell(1, 0), rng.cell(0, 1), rng.cell(1, 1)]
    originals = [
        (c.TopBorder, c.BottomBorder, c.LeftBorder, c.RightBorder) for c in cells
    ]

    try:
        rng.borders.left.color = 0x0F0F0F
        rng.borders.left.weight = 30
        rng.borders.left.line_style = 1

        for cell in cells:
            assert cell.borders.left.color == 0x0F0F0F
            assert abs(cell.borders.left.weight - 30) <= 2
    finally:
        for cell, (top, bottom, left, right) in zip(cells, originals):
            cell.TopBorder = top
            cell.BottomBorder = bottom
            cell.LeftBorder = left
            cell.RightBorder = right


def test_range_borders_accept_borderstyle_config():
    _, _, sheet = _connect_or_skip()
    rng = sheet.range(7, 8, 8, 9)

    cells = [rng.cell(0, 0), rng.cell(1, 0), rng.cell(0, 1), rng.cell(1, 1)]
    originals = [
        (c.TopBorder, c.BottomBorder, c.LeftBorder, c.RightBorder) for c in cells
    ]

    cfg = Borders(all=BorderStyle(color=0x123456, weight=40, line_style=1))

    try:
        rng.borders = cfg

        for cell in cells:
            assert cell.borders.top.color == 0x123456
            assert cell.borders.bottom.color == 0x123456
            assert cell.borders.left.color == 0x123456
            assert cell.borders.right.color == 0x123456
    finally:
        for cell, (top, bottom, left, right) in zip(cells, originals):
            cell.TopBorder = top
            cell.BottomBorder = bottom
            cell.LeftBorder = left
            cell.RightBorder = right


# 外枠の罫線のみ引く場合
def test_range_borders_around():
    _, _, sheet = _connect_or_skip()
    rng = sheet.range("B2:D4")

    cells = [
        rng.cell(0, 0), rng.cell(1, 0), rng.cell(2, 0), 
        rng.cell(0, 1), rng.cell(1, 1), rng.cell(2, 1), 
        rng.cell(0, 2), rng.cell(1, 2), rng.cell(2, 2), 
        ]
    originals = [
        (c.TopBorder, c.BottomBorder, c.LeftBorder, c.RightBorder) for c in cells
    ]

    border_stlye_zero = BorderStyle(color=0xFFFFFF, weight=0, line_style=BorderLineStyle.NONE )
    border_style = BorderStyle(color=0x000000, weight=20, line_style=BorderLineStyle.SOLID )

    # すべての罫線を一旦消す
    rng.borders.all = border_stlye_zero

    try:
        # 外枠のみ罫線を引く
        rng.borders.around = border_style

        # ひとつずつセルの罫線をチェックする
        assert rng.cell(0, 0).borders.top.line_style == BorderLineStyle.SOLID
        assert rng.cell(0, 0).borders.right.line_style == BorderLineStyle.NONE
        assert rng.cell(0, 0).borders.bottom.line_style == BorderLineStyle.NONE
        assert rng.cell(0, 0).borders.left.line_style == BorderLineStyle.SOLID

        assert rng.cell(1, 0).borders.top.line_style == BorderLineStyle.SOLID
        assert rng.cell(1, 0).borders.right.line_style == BorderLineStyle.NONE
        assert rng.cell(1, 0).borders.bottom.line_style == BorderLineStyle.NONE
        assert rng.cell(1, 0).borders.left.line_style == BorderLineStyle.NONE

        assert rng.cell(2, 0).borders.top.line_style == BorderLineStyle.SOLID
        assert rng.cell(2, 0).borders.right.line_style == BorderLineStyle.SOLID
        assert rng.cell(2, 0).borders.bottom.line_style == BorderLineStyle.NONE
        assert rng.cell(2, 0).borders.left.line_style == BorderLineStyle.NONE

        assert rng.cell(0, 1).borders.top.line_style == BorderLineStyle.NONE
        assert rng.cell(0, 1).borders.right.line_style == BorderLineStyle.NONE
        assert rng.cell(0, 1).borders.bottom.line_style == BorderLineStyle.NONE
        assert rng.cell(0, 1).borders.left.line_style == BorderLineStyle.SOLID

        assert rng.cell(1, 1).borders.top.line_style == BorderLineStyle.NONE
        assert rng.cell(1, 1).borders.right.line_style == BorderLineStyle.NONE
        assert rng.cell(1, 1).borders.bottom.line_style == BorderLineStyle.NONE
        assert rng.cell(1, 1).borders.left.line_style == BorderLineStyle.NONE

        assert rng.cell(2, 1).borders.top.line_style == BorderLineStyle.NONE
        assert rng.cell(2, 1).borders.right.line_style == BorderLineStyle.SOLID
        assert rng.cell(2, 1).borders.bottom.line_style == BorderLineStyle.NONE
        assert rng.cell(2, 1).borders.left.line_style == BorderLineStyle.NONE

        assert rng.cell(0, 2).borders.top.line_style == BorderLineStyle.NONE
        assert rng.cell(0, 2).borders.right.line_style == BorderLineStyle.NONE
        assert rng.cell(0, 2).borders.bottom.line_style == BorderLineStyle.SOLID
        assert rng.cell(0, 2).borders.left.line_style == BorderLineStyle.SOLID

        assert rng.cell(1, 2).borders.top.line_style == BorderLineStyle.NONE
        assert rng.cell(1, 2).borders.right.line_style == BorderLineStyle.NONE
        assert rng.cell(1, 2).borders.bottom.line_style == BorderLineStyle.SOLID
        assert rng.cell(1, 2).borders.left.line_style == BorderLineStyle.NONE

        assert rng.cell(2, 2).borders.top.line_style == BorderLineStyle.NONE
        assert rng.cell(2, 2).borders.right.line_style == BorderLineStyle.SOLID
        assert rng.cell(2, 2).borders.bottom.line_style == BorderLineStyle.SOLID
        assert rng.cell(2, 2).borders.left.line_style == BorderLineStyle.NONE

    finally:
        for cell, (top, bottom, left, right) in zip(cells, originals):
            cell.TopBorder = top
            cell.BottomBorder = bottom
            cell.LeftBorder = left
            cell.RightBorder = right

# 外枠の罫線のみ引く場合（元の罫線は残る）
def test_range_borders_around_red():
    _, _, sheet = _connect_or_skip()
    rng = sheet.range("B2:D4")

    cells = [
        rng.cell(0, 0), rng.cell(1, 0), rng.cell(2, 0), 
        rng.cell(0, 1), rng.cell(1, 1), rng.cell(2, 1), 
        rng.cell(0, 2), rng.cell(1, 2), rng.cell(2, 2), 
        ]
    originals = [
        (c.TopBorder, c.BottomBorder, c.LeftBorder, c.RightBorder) for c in cells
    ]

    border_stlye_red = BorderStyle(color=0xFF0000, weight=20, line_style=BorderLineStyle.SOLID )
    border_style = BorderStyle(color=0x000000, weight=20, line_style=BorderLineStyle.SOLID )

    # すべての罫線を一旦赤線で引く
    rng.borders.all = border_stlye_red
    pass

    try:
        # 外枠のみ罫線を引く
        rng.borders.around = border_style

        # 内側の赤い罫線が消されてないことを確認する
        assert rng.cell(0, 0).borders.top.color == 0x000000
        assert rng.cell(0, 0).borders.right.color == 0xFF0000
        assert rng.cell(0, 0).borders.bottom.color == 0xFF0000
        assert rng.cell(0, 0).borders.left.color == 0x000000

        assert rng.cell(1, 1).borders.top.color == 0xFF0000
        assert rng.cell(1, 1).borders.right.color == 0xFF0000
        assert rng.cell(1, 1).borders.bottom.color == 0xFF0000
        assert rng.cell(1, 1).borders.left.color == 0xFF0000

        assert rng.cell(2, 2).borders.top.color == 0xFF0000
        assert rng.cell(2, 2).borders.right.color == 0x000000
        assert rng.cell(2, 2).borders.bottom.color == 0x000000
        assert rng.cell(2, 2).borders.left.color == 0xFF0000

    finally:
        for cell, (top, bottom, left, right) in zip(cells, originals):
            cell.TopBorder = top
            cell.BottomBorder = bottom
            cell.LeftBorder = left
            cell.RightBorder = right

# 内枠の罫線のみ引く場合
def test_range_borders_inner():
    _, _, sheet = _connect_or_skip()
    rng = sheet.range("B2:D4")

    cells = [
        rng.cell(0, 0), rng.cell(1, 0), rng.cell(2, 0), 
        rng.cell(0, 1), rng.cell(1, 1), rng.cell(2, 1), 
        rng.cell(0, 2), rng.cell(1, 2), rng.cell(2, 2), 
        ]
    originals = [
        (c.TopBorder, c.BottomBorder, c.LeftBorder, c.RightBorder) for c in cells
    ]

    border_stlye_zero = BorderStyle(color=0xFFFFFF, weight=0, line_style=BorderLineStyle.NONE )
    border_style = BorderStyle(color=0x000000, weight=20, line_style=BorderLineStyle.SOLID )

    # すべての罫線を一旦消す
    rng.borders.all = border_stlye_zero

    try:
        # 外枠のみ罫線を引く
        rng.borders.inner = border_style

        # ひとつずつセルの罫線をチェックする
        assert rng.cell(0, 0).borders.top.line_style == BorderLineStyle.NONE
        assert rng.cell(0, 0).borders.right.line_style == BorderLineStyle.SOLID
        assert rng.cell(0, 0).borders.bottom.line_style == BorderLineStyle.SOLID
        assert rng.cell(0, 0).borders.left.line_style == BorderLineStyle.NONE

        assert rng.cell(1, 0).borders.top.line_style == BorderLineStyle.NONE
        assert rng.cell(1, 0).borders.right.line_style == BorderLineStyle.SOLID
        assert rng.cell(1, 0).borders.bottom.line_style == BorderLineStyle.SOLID
        assert rng.cell(1, 0).borders.left.line_style == BorderLineStyle.SOLID

        assert rng.cell(2, 0).borders.top.line_style == BorderLineStyle.NONE
        assert rng.cell(2, 0).borders.right.line_style == BorderLineStyle.NONE
        assert rng.cell(2, 0).borders.bottom.line_style == BorderLineStyle.SOLID
        assert rng.cell(2, 0).borders.left.line_style == BorderLineStyle.SOLID

        assert rng.cell(0, 1).borders.top.line_style == BorderLineStyle.SOLID
        assert rng.cell(0, 1).borders.right.line_style == BorderLineStyle.SOLID
        assert rng.cell(0, 1).borders.bottom.line_style == BorderLineStyle.SOLID
        assert rng.cell(0, 1).borders.left.line_style == BorderLineStyle.NONE

        assert rng.cell(1, 1).borders.top.line_style == BorderLineStyle.SOLID
        assert rng.cell(1, 1).borders.right.line_style == BorderLineStyle.SOLID
        assert rng.cell(1, 1).borders.bottom.line_style == BorderLineStyle.SOLID
        assert rng.cell(1, 1).borders.left.line_style == BorderLineStyle.SOLID

        assert rng.cell(2, 1).borders.top.line_style == BorderLineStyle.SOLID
        assert rng.cell(2, 1).borders.right.line_style == BorderLineStyle.NONE
        assert rng.cell(2, 1).borders.bottom.line_style == BorderLineStyle.SOLID
        assert rng.cell(2, 1).borders.left.line_style == BorderLineStyle.SOLID

        assert rng.cell(0, 2).borders.top.line_style == BorderLineStyle.SOLID
        assert rng.cell(0, 2).borders.right.line_style == BorderLineStyle.SOLID
        assert rng.cell(0, 2).borders.bottom.line_style == BorderLineStyle.NONE
        assert rng.cell(0, 2).borders.left.line_style == BorderLineStyle.NONE

        assert rng.cell(1, 2).borders.top.line_style == BorderLineStyle.SOLID
        assert rng.cell(1, 2).borders.right.line_style == BorderLineStyle.SOLID
        assert rng.cell(1, 2).borders.bottom.line_style == BorderLineStyle.NONE
        assert rng.cell(1, 2).borders.left.line_style == BorderLineStyle.SOLID

        assert rng.cell(2, 2).borders.top.line_style == BorderLineStyle.SOLID
        assert rng.cell(2, 2).borders.right.line_style == BorderLineStyle.NONE
        assert rng.cell(2, 2).borders.bottom.line_style == BorderLineStyle.NONE
        assert rng.cell(2, 2).borders.left.line_style == BorderLineStyle.SOLID

    finally:
        for cell, (top, bottom, left, right) in zip(cells, originals):
            cell.TopBorder = top
            cell.BottomBorder = bottom
            cell.LeftBorder = left
            cell.RightBorder = right