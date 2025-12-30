import pytest

from excellikeuno.connection import connect_calc


def _connect_or_skip():
    try:
        return connect_calc()
    except RuntimeError as exc:
        pytest.skip(f"UNO runtime not available: {exc}")


def test_value_roundtrip():
    _, _, sheet = _connect_or_skip()
    cell = sheet.cell(0, 0)
    original = cell.value
    cell.value = 12.34
    assert cell.value == 12.34
    cell.value = original  # restore


def test_formula_roundtrip():
    _, _, sheet = _connect_or_skip()
    cell = sheet.cell(0, 1)
    original_formula = cell.formula
    cell.formula = "=1+2"
    assert cell.formula == "=1+2"
    cell.formula = original_formula  # restore


def test_props_passthrough():
    _, _, sheet = _connect_or_skip()
    cell = sheet.cell(1, 0)
    props = cell.props
    original_color = props.getPropertyValue("CellBackColor")
    props.setPropertyValue("CellBackColor", 0x112233)
    assert props.getPropertyValue("CellBackColor") == 0x112233
    props.setPropertyValue("CellBackColor", original_color)


def test_cellproperties_attribute_access():
    _, _, sheet = _connect_or_skip()
    cell = sheet.cell(2, 0)
    original_color = cell.CellBackColor
    new_color = 0x223344 if original_color != 0x223344 else 0x556677
    try:
        cell.CellBackColor = new_color
        assert cell.CellBackColor == new_color
    finally:
        cell.CellBackColor = original_color


def test_cell_a1_alias():
    _, _, sheet = _connect_or_skip()
    a1 = sheet.cell("A1")
    zero_zero = sheet.cell(0, 0)
    a1.value = 99.0
    try:
        assert a1.value == zero_zero.value == 99.0
    finally:
        a1.value = 0.0


def test_cell_a1_with_dollar():
    _, _, sheet = _connect_or_skip()
    b3 = sheet.cell("$B$3")
    b3.text = "hello"
    try:
        assert b3.text == "hello"
    finally:
        b3.text = ""
