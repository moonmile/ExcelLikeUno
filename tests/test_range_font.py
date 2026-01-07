import pytest

from excellikeuno.connection import connect_calc


def _connect_or_skip():
    try:
        return connect_calc()
    except RuntimeError as exc:  # pragma: no cover - environment dependent
        pytest.skip(f"UNO runtime not available: {exc}")


def test_range_font_broadcast_size_and_bold():
    _, _, sheet = _connect_or_skip()
    rng = sheet.range(0, 5, 1, 6)  # 2x2 block

    cells = [rng.cell(0, 0), rng.cell(1, 0), rng.cell(0, 1), rng.cell(1, 1)]
    originals = [(c.CharHeight, c.CharWeight) for c in cells]

    new_size = 15.0
    new_bold = True

    try:
        rng.font.size = new_size
        rng.font.bold = new_bold

        for cell in cells:
            assert abs(cell.CharHeight - new_size) <= 0.5
            assert (cell.CharWeight >= 150.0) is new_bold
    finally:
        for cell, (h, w) in zip(cells, originals):
            cell.CharHeight = h
            cell.CharWeight = w


def test_range_font_broadcast_color_and_backcolor():
    _, _, sheet = _connect_or_skip()
    rng = sheet.range(2, 5, 3, 6)  # another 2x2 block

    cells = [rng.cell(0, 0), rng.cell(1, 0), rng.cell(0, 1), rng.cell(1, 1)]
    original_colors = [c.CharColor for c in cells]
    try:
        original_backs = [c.CellBackColor for c in cells]
    except Exception:
        original_backs = [None for _ in cells]

    new_color = 0x334455
    new_back = 0x99AABB

    try:
        rng.font.color = new_color
        rng.font.backcolor = new_back

        for cell in cells:
            assert cell.CharColor == new_color
            assert cell.font.backcolor == new_back
    finally:
        for cell, color in zip(cells, original_colors):
            cell.CharColor = color
        for cell, back in zip(cells, original_backs):
            if back is not None:
                cell.CellBackColor = back
