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
