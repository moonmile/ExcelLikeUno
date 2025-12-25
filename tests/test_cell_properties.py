import pytest

from excellikeuno.connection import connect_calc


def _connect_or_skip():
    try:
        return connect_calc()
    except RuntimeError as exc:
        pytest.skip(f"UNO runtime not available: {exc}")


def test_cellproperties_get_set_property_methods():
    _, _, sheet = _connect_or_skip()
    cell = sheet.cell(0, 2)
    original_color = cell.CellBackColor
    new_color = 0x778899 if original_color != 0x778899 else 0x99aabb
    try:
        cell.CellBackColor = new_color
        assert cell.CellBackColor == new_color
    finally:
        cell.CellBackColor = original_color


def test_cellproperties_uno_method_aliases():
    _, _, sheet = _connect_or_skip()
    cell = sheet.cell(1, 2)
    props = cell.props
    original_color = props.getPropertyValue("CellBackColor")
    new_color = 0xaabbcc if original_color != 0xaabbcc else 0xccbbaa
    try:
        props.setPropertyValue("CellBackColor", new_color)
        assert cell.CellBackColor == new_color
    finally:
        props.setPropertyValue("CellBackColor", original_color)


def test_cellproperties_attribute_passthrough():
    _, _, sheet = _connect_or_skip()
    cell = sheet.cell(2, 2)
    original_wrap = cell.IsTextWrapped
    new_wrap = not original_wrap
    try:
        cell.IsTextWrapped = new_wrap
        assert cell.IsTextWrapped is new_wrap
    finally:
        cell.IsTextWrapped = original_wrap
