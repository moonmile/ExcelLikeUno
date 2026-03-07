# from com.sun.star.awt.Size import Size
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
    originals = [(c.font.size, c.font.bold) for c in cells]

    new_size = 15.0
    new_bold = True

    try:
        rng.font.apply(size=new_size, bold=new_bold)

        for cell in cells:
            assert abs(cell.font.size - new_size) <= 0.5
            assert cell.font.bold is new_bold
    finally:
        for cell, (h, w) in zip(cells, originals):
            cell.font.apply(size=h, bold=w)


def test_range_font_size_roundtrip_via_proxy():
    _, _, sheet = _connect_or_skip()
    rng = sheet.range(4, 5, 5, 6)  # 2x2 block separate from other test data

    cells = [rng.cell(0, 0), rng.cell(1, 0), rng.cell(0, 1), rng.cell(1, 1)]
    originals = [float(c.font.size) for c in cells]

    new_size = 18.5

    try:
        rng.font.apply(size=new_size)

        for cell in cells:
            assert abs(cell.font.size - new_size) <= 0.5

        assert abs(float(rng.font.size) - new_size) <= 0.5
    finally:
        for cell, height in zip(cells, originals):
            cell.font.apply(size=height)


def test_range_font_broadcast_color_and_backcolor():
    _, _, sheet = _connect_or_skip()
    rng = sheet.range(2, 5, 3, 6)  # another 2x2 block

    cells = [rng.cell(0, 0), rng.cell(1, 0), rng.cell(0, 1), rng.cell(1, 1)]
    original_colors = [c.font.color for c in cells]
    original_backs = [c.font.backcolor for c in cells]

    new_color = 0x334455
    new_back = 0x99AABB

    try:
        rng.font.apply(color=new_color, backcolor=new_back)

        for cell in cells:
            assert cell.font.color == new_color
            assert cell.font.backcolor == new_back
    finally:
        for cell, color in zip(cells, original_colors):
            cell.font.apply(color=color)
        for cell, back in zip(cells, original_backs):
            cell.font.apply(backcolor=back)


def test_range_font_broadcast_size_kanji():
    _, _, sheet = _connect_or_skip()
    rng = sheet.range(2, 5, 3, 6)  # another 2x2 block

    cells = [rng.cell(0, 0), rng.cell(1, 0), rng.cell(0, 1), rng.cell(1, 1)]
    orizinal_sizes = [float(c.font.size) for c in cells]

    new_size = 20.0

    try:
        rng.font.apply(size=new_size)
        for cell in cells:
            assert abs(cell.font.size - new_size) <= 0.5
        rng.value = [["あ", "あ"], ["あ", "あ"]]
        for cell in cells:
            assert cell.text == "あ"
            assert abs(cell.font.size - new_size) <= 0.5

    finally:
        # for cell, height in zip(cells, orizinal_sizes):
        #    cell.char_props.CharHeight = height
        pass
