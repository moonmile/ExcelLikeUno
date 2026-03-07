import pytest

from excellikeuno.connection import connect_calc
from excellikeuno.style.font import Font


def _connect_or_skip():
    try:
        return connect_calc()
    except RuntimeError as exc:  # pragma: no cover - environment dependent
        pytest.skip(f"UNO runtime not available: {exc}")


def test_cell_font_size_and_bold_roundtrip():
    _, _, sheet = _connect_or_skip()
    cell = sheet.cell(0, 4)

    original_size = float(cell.props.CharHeight)
    original_weight = float(cell.props.CharWeight)

    new_size = 16.0 if abs(original_size - 16.0) > 0.01 else 14.0
    new_bold = False if original_weight >= 150.0 else True

    try:
        cell.font.apply(size=new_size, bold=new_bold)

        assert abs(cell.font.size - new_size) <= 0.5
        assert cell.font.bold is new_bold
    finally:
        cell.font.apply(size=original_size, bold=(original_weight >= 150.0))

        cell.props.setPropertyValue("CharHeight", original_size)
        cell.props.setPropertyValue("CharWeight", original_weight)


def test_cell_font_color_and_backcolor_roundtrip():
    _, _, sheet = _connect_or_skip()
    cell = sheet.cell(1, 4)

    original_color = cell.font.color
    original_back = cell.font.backcolor

    new_color = 0x112233 if original_color != 0x112233 else 0x445566
    new_back = 0xAABBCC if original_back != 0xAABBCC else 0xCCBBAA

    try:
        cell.font.apply(color=new_color, backcolor=new_back)

        assert cell.font.color == new_color
        assert cell.font.backcolor == new_back
    finally:
        cell.font.apply(color=original_color, backcolor=original_back)


def test_font_as_config_holder_without_setter():
    font = Font(size=11, bold=True)
    assert font.size == 11
    assert font.bold is True

    font.italic = True
    font.underline = 1

    assert font.italic is True
    assert font.underline == 1


def test_font_buffers_when_setter_raises():
    def getter():
        return {"size": 8, "bold": False}

    def failing_setter(**kwargs):
        raise RuntimeError("setter failed")

    font = Font(getter=getter, setter=failing_setter)
    font.size = 12
    font.bold = True

    # buffered values should override getter results even when setter fails
    assert font.size == 12
    assert font.bold is True


def test_font_without_owner_reusable_on_multiple_cells():
    _, _, sheet = _connect_or_skip()
    c1 = sheet.cell(2, 6)
    c2 = sheet.cell(2, 7)

    originals = [
        (float(c1.font.size), bool(c1.font.bold)),
        (float(c2.font.size), bool(c2.font.bold)),
    ]

    font = Font(size=13, bold=True)

    try:
        c1.font = font
        c2.font = font

        assert abs(c1.font.size - 13.0) <= 0.5
        assert c1.font.bold is True
        assert abs(c2.font.size - 13.0) <= 0.5
        assert c2.font.bold is True

        # font remains reusable/config-holder after applications
        assert font.size == 13
        assert font.bold is True
    finally:
        (c1_height, c1_bold), (c2_height, c2_bold) = originals
        c1.font.apply(size=c1_height, bold=c1_bold)
        c2.font.apply(size=c2_height, bold=c2_bold)
