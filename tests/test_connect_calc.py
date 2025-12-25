import pytest
import uno

from excellikeuno import InterfaceNames, UnoObject, connect_calc


def test_uno_connect_calc_is_registered():
    """Importing excellikeuno should register uno.connect_calc when UNO is present."""
    assert hasattr(uno, "connect_calc")

def _connect_or_skip():
    try:
        return connect_calc()
    except RuntimeError as exc:
        pytest.skip(f"UNO runtime not available: {exc}")


def test_connect_calc_returns_handles():
    """connect_calc should return desktop, document, and Sheet wrapper."""
    desktop, doc, sheet = _connect_or_skip()
    assert desktop is not None
    assert doc is not None
    assert sheet is not None


def test_connect_calc_document_has_sheets():
    """Connected document should expose sheets interface via UnoObject."""
    _, doc, _ = uno.connect_calc()
    doc_wrapper = UnoObject(doc)
    sheets_iface = doc_wrapper.iface(InterfaceNames.X_SPREADSHEET_DOCUMENT)
    assert sheets_iface is not None
    assert hasattr(sheets_iface, "getSheets")


def test_connect_calc_get_active_sheet():
    """Should retrieve active sheet through the returned Sheet wrapper."""
    _, _, sheet = uno.connect_calc()
    cell = sheet.cell(0, 0)
    assert cell is not None


def test_connect_calc_cell_value_read_write():
    """Should read and write cell values via Sheet/Cell wrappers."""
    _, _, sheet = uno.connect_calc()
    cell = sheet.cell(0, 0)
    cell.value = 42.0
    assert cell.value == 42.0


def test_connect_calc_multiple_calls_same_instance():
    """Multiple connect_calc calls should work consistently."""
    first = _connect_or_skip()
    second = _connect_or_skip()
    assert first is not None
    assert second is not None

