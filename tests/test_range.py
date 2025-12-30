import os
import pytest

from excellikeuno.table import Sheet
from excellikeuno.connection import connect_calc


def _connect_or_skip():
    try:
        return connect_calc()
    except RuntimeError as exc:
        pytest.skip(f"UNO runtime not available: {exc}")


def test_range_cell_access():
    _, doc, sheet = _connect_or_skip()
    sheet.cell(0, 0).text = "id"
    sheet.cell(1, 1).value = 1
    rng = sheet.range(0, 0, 1, 1)
    assert rng.cell(0, 0).text == "id"
    assert rng.cell(1, 1).value == 1

def test_range_subrange_and_aliases():
    _, doc, sheet = _connect_or_skip()
    sheet.cell(0, 0).text = "A1"
    sheet.cell(1, 0).text = "B1"
    sheet.cell(2, 1).text = "C2"

    rng = sheet.range(0, 0, 2, 2)
    sub = rng.subrange(0, 0, 0, 0)
    assert sub.cell(0, 0).text == "A1"
    # aliases should behave the same
    assert rng.getCellByPosition(1, 0).text == "B1"
    nested = rng.getCellRangeByPosition(2, 1, 2, 1)
    assert nested.cell(0, 0).text == "C2"
    
    def test_range_a1_notation():
        _, doc, sheet = _connect_or_skip()
        sheet.cell("A1").text = "top-left"
        sheet.cell("B2").text = "center"
        sheet.cell("C3").text = "bottom-right"

        rng = sheet.range("A1:C3")
        assert rng.cell(0, 0).text == "top-left"
        assert rng.cell(1, 1).text == "center"
        assert rng.cell(2, 2).text == "bottom-right"

        sub = sheet.range("A1", "B2")
        assert sub.cell(1, 1).text == "center"
