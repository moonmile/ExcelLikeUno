import pytest

from excellikeuno.table import SheetCell, RawProps
from excellikeuno.connection import connect_calc


def _connect_or_skip():
    try:
        return connect_calc()
    except RuntimeError as exc:
        pytest.skip(f"UNO runtime not available: {exc}")


def test_sheet_cell_value_roundtrip():
    _, _, sheet = _connect_or_skip()
    cell = sheet.sheet_cell(0, 0)
    assert isinstance(cell, SheetCell)
    original = cell.value
    cell.value = 42.5
    try:
        assert cell.value == 42.5
    finally:
        cell.value = original


def test_sheet_cell_props_background_color():
    _, _, sheet = _connect_or_skip()
    cell = sheet.sheet_cell(1, 0)
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
    cell = sheet.sheet_cell(2, 0)
    original_size = cell.font.size
    cell.font.size = 15
    try:
        assert cell.font.size == 15
    finally:
        cell.font.size = original_size
