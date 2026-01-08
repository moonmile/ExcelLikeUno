import pytest


def _connect_or_skip():
    try:
        from excellikeuno.connection import connect_calc
    except Exception:
        pytest.skip("connect_calc not available")
    try:
        return connect_calc()
    except Exception as exc:
        pytest.skip(f"UNO runtime not available: {exc}")


def test_cell_value_numeric_roundtrip():
    _, __, sheet = _connect_or_skip()
    cell = sheet.cell(0, 5)
    original_formula = cell.formula
    try:
        target = 123.45
        cell.value = target
        assert abs(cell.value - target) < 0.001
    finally:
        try:
            cell.formula = original_formula
        except Exception:
            pass


def test_cell_value_string_roundtrip():
    _, __, sheet = _connect_or_skip()
    cell = sheet.cell(1, 5)
    original_formula = cell.formula
    try:
        text = "hello"
        cell.value = text
        assert cell.text == text
    finally:
        try:
            cell.formula = original_formula
        except Exception:
            pass
