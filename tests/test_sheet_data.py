import os
import pathlib
import pytest

# Note: このテストは test_sheet_001.ods を開いてテストするので、skip しても問題なし。

from excellikeuno.connection import open_calc_document


def _open_or_skip():
    path = pathlib.Path(__file__).parent / "data" / "test_sheet_001.ods"
    if os.environ.get("ENABLE_OPEN_CALC_DOC_TESTS") != "1":
        pytest.skip("Set ENABLE_OPEN_CALC_DOC_TESTS=1 to run open_calc_document-based tests")
    try:
        desktop, doc, sheet = open_calc_document(str(path))
    except RuntimeError as exc:  # pragma: no cover - depends on LibreOffice runtime
        pytest.skip(f"UNO runtime not available: {exc}")
    return desktop, doc, sheet


def _close_safe(doc):
    try:
        doc.close(True)
    except Exception:
        pass


def test_sheet_sample_text_and_numbers():
    _, doc, sheet = _open_or_skip()
    try:
        assert sheet.cell(0, 0).text == "id"
        assert sheet.cell(1, 0).text == "name"
        assert sheet.cell(2, 0).text == "address"

        assert sheet.cell(0, 1).value == 1
        assert sheet.cell(1, 1).text == "masuda"
        assert sheet.cell(2, 1).text == "tokyo"
    finally:
        _close_safe(doc)


def test_sheet_sample_cell_border_and_fill():
    _, doc, sheet = _open_or_skip()
    try:
        cell = sheet.cell(0, 0)
        assert int(cell.CellBackColor) == 0xFFBF00

        border = cell.TopBorder
        assert int(border.Color) == 0  # black
        # 0.74pt border is roughly 26 twips (1/100 mm); allow small variation
        assert abs(int(border.OuterLineWidth) - 26) <= 3

        cell = sheet.cell(7, 0)
        border = cell.TopBorder
        assert int(border.Color) == 0  # black
        assert int(border.OuterLineWidth) == 0  # no border

    finally:
        _close_safe(doc)
