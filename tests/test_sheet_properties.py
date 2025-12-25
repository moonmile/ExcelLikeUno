import pytest

from excellikeuno.connection import connect_calc


def _connect_or_skip():
    try:
        return connect_calc()
    except RuntimeError as exc:
        pytest.skip(f"UNO runtime not available: {exc}")


def test_sheet_name_roundtrip():
    _, _, sheet = _connect_or_skip()
    original = sheet.name
    new_name = f"{original[:25]}_tmp"
    try:
        sheet.name = new_name
        assert sheet.name == new_name
    finally:
        sheet.name = original

"""
def test_sheet_visibility_property():
    _, _, sheet = _connect_or_skip()
    original = sheet.is_visible
    try:
        sheet.is_visible = not original
        assert sheet.is_visible is (not original)
    finally:
        sheet.is_visible = original
"""

def test_sheet_page_style_and_tab_color():
    _, _, sheet = _connect_or_skip()
    original_style = sheet.page_style
    original_color = sheet.tab_color
    new_color = 0x112233 if original_color != 0x112233 else 0x445566
    try:
        sheet.page_style = original_style or "Default"
        assert sheet.page_style == (original_style or "Default")
        sheet.tab_color = new_color
        assert sheet.tab_color == new_color
    finally:
        sheet.page_style = original_style
        sheet.tab_color = original_color


def test_sheet_table_layout_and_print_area():
    _, _, sheet = _connect_or_skip()
    original_layout = sheet.table_layout
    original_print_area = sheet.automatic_print_area
    try:
        sheet.table_layout = original_layout
        assert sheet.table_layout == original_layout
        sheet.automatic_print_area = not original_print_area
        assert sheet.automatic_print_area is (not original_print_area)
    finally:
        sheet.table_layout = original_layout
        sheet.automatic_print_area = original_print_area


def test_sheet_conditional_formats_passthrough():
    _, _, sheet = _connect_or_skip()
    assert sheet.conditional_formats is not None


def test_sheet_cell_uses_xspreadsheet():
    _, _, sheet = _connect_or_skip()
    cell = sheet.cell(2, 3)
    original_value = cell.value
    try:
        cell.value = 99.0
        assert cell.value == 99.0
    finally:
        cell.value = original_value
