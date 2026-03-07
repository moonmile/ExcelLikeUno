from calendar import c

import pytest

from excellikeuno.sheet import SheetCell, SheetCellRange
from excellikeuno.sheet.sheet_cell import RawProps
from excellikeuno.connection import connect_calc


def _connect_or_skip():
    try:
        return connect_calc()
    except RuntimeError as exc:
        pytest.skip(f"UNO runtime not available: {exc}")


def test_sheet_cell_value_roundtrip():
    _, _, sheet = _connect_or_skip()
    cell = sheet.cell(0, 0)
    assert isinstance(cell, SheetCell)
    original = cell.value
    cell.value = 42.5
    try:
        assert cell.value == 42.5
    finally:
        cell.value = original


def test_sheet_cell_props_background_color():
    _, _, sheet = _connect_or_skip()
    cell = sheet.cell(1, 0)
    props = cell.props
    assert isinstance(props, RawProps)
    original_color = props.CellBackColor  # type: ignore[attr-defined]
    new_color = 0x112233 if original_color != 0x112233 else 0x445566
    try:
        props.CellBackColor = new_color  # type: ignore[attr-defined]
        assert props.CellBackColor == new_color  # type: ignore[attr-defined]
    finally:
        props.CellBackColor = original_color  # type: ignore[attr-defined]


def test_sheet_cell_font_roundtrip():
    _, _, sheet = _connect_or_skip()
    cell = sheet.cell(2, 0)
    original_size = cell.font.size
    cell.font.size = 15
    try:
        assert cell.font.size == 15
    finally:
        cell.font.size = original_size


def test_sheet_cell_range_returns_sheetcell():
    _, _, sheet = _connect_or_skip()
    rng = sheet.range(0, 0, 1, 1)
    assert isinstance(rng, SheetCellRange)
    cells = rng.getCells()
    assert isinstance(cells[0][0], SheetCell)


def test_sheet_cell_attribute_properties():
    _, _, sheet = _connect_or_skip()
    cell = sheet.cell(3, 0)

    original_backcolor = cell.backcolor  # type: ignore[attr-defined]
    original_color     = cell.color  # type: ignore[attr-defined]
    original_vertical_align = cell.vertical_align
    original_horizontal_align = cell.horizontal_align

    try:
        cell.backcolor = 0x123456  # type: ignore[attr-defined]
        cell.color = 0x654321  # type: ignore[attr-defined]
        cell.vertical_align = 1 # TOP
        cell.horizontal_align = 1 # LEFT

        assert cell.backcolor == 0x123456  # type: ignore[attr-defined]
        assert cell.color == 0x654321  # type: ignore[attr-defined]
        assert cell.vertical_align == 1
        assert cell.horizontal_align == 1
    finally:
        cell.backcolor = original_backcolor  # type: ignore[attr-defined]
        cell.color = original_color  # type: ignore[attr-defined]
        cell.vertical_align = original_vertical_align
        cell.horizontal_align = original_horizontal_align
