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


def test_cell_topborder_roundtrip():
    _, _, sheet = _connect_or_skip()
    cell = sheet.cell(3, 2)

    try:
        import uno
        from com.sun.star.table import BorderLine  # type: ignore
    except Exception as exc:  # pragma: no cover - depends on LibreOffice runtime
        pytest.skip(f"UNO runtime not available: {exc}")

    # capture original border and prepare a new one
    original_border = cell.TopBorder
    new_border = uno.createUnoStruct("com.sun.star.table.BorderLine")
    new_border.Color = 0x123456
    new_border.InnerLineWidth = 0
    new_border.OuterLineWidth = 50
    new_border.LineDistance = 0

    try:
        cell.TopBorder = new_border
        updated = cell.TopBorder
        assert updated.Color == 0x123456
        # LibreOffice may normalize to the nearest supported width; allow small drift.
        assert abs(int(updated.OuterLineWidth) - 50) <= 2
    finally:
        cell.TopBorder = original_border


def test_cell_topborder2_roundtrip():
    _, _, sheet = _connect_or_skip()
    cell = sheet.cell(4, 2)

    try:
        import uno
        from com.sun.star.table import BorderLine2  # type: ignore
    except Exception as exc:  # pragma: no cover - depends on LibreOffice runtime
        pytest.skip(f"UNO runtime not available: {exc}")

    original_border = cell.TopBorder2
    new_border = uno.createUnoStruct("com.sun.star.table.BorderLine2")
    new_border.Color = 0x654321
    new_border.InnerLineWidth = 0
    new_border.OuterLineWidth = 60
    new_border.LineDistance = 0
    new_border.LineStyle = 0
    new_border.LineWidth = 60

    try:
        cell.TopBorder2 = new_border
        updated = cell.TopBorder2
        assert updated.Color == 0x654321
        assert abs(int(updated.OuterLineWidth) - 60) <= 2
        assert abs(int(updated.LineWidth) - 60) <= 2
    finally:
        cell.TopBorder2 = original_border


def test_cell_horijustify_enum_roundtrip():
    _, _, sheet = _connect_or_skip()
    from excellikeuno.typing import CellHoriJustify

    cell = sheet.cell(5, 2)
    original = cell.HoriJustify
    new_value = CellHoriJustify.LEFT if original != CellHoriJustify.LEFT else CellHoriJustify.RIGHT
    try:
        cell.HoriJustify = new_value
        assert cell.HoriJustify == new_value
    finally:
        cell.HoriJustify = original


def test_cell_vertjustify_enum_roundtrip():
    _, _, sheet = _connect_or_skip()
    from excellikeuno.typing import CellVertJustify

    cell = sheet.cell(5, 3)
    original = cell.VertJustify
    new_value = CellVertJustify.TOP if original != CellVertJustify.TOP else CellVertJustify.BOTTOM
    try:
        cell.VertJustify = new_value
        assert cell.VertJustify == new_value
    finally:
        cell.VertJustify = original
