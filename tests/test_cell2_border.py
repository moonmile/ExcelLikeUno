import pytest

from excellikeuno.connection import connect_calc
from excellikeuno.style.border import Borders, BorderStyle
from excellikeuno.typing import BorderLine


def _connect_or_skip():
    try:
        return connect_calc()
    except RuntimeError as exc:  # pragma: no cover - environment dependent
        pytest.skip(f"UNO runtime not available: {exc}")


def test_cell_border_roundtrip_via_proxy():
    _, _, sheet = _connect_or_skip()
    cell = sheet.cell2(0, 0)
    
    top_color = cell.borders.top.color
    top_width = cell.borders.top.width

    cell.borders.top.color = 0x112233
    cell.borders.top.width = 20

    assert cell.borders.top.color == 0x112233
    assert cell.borders.top.width == 20

    cell.borders.top.color = top_color
    cell.borders.top.width = top_width
