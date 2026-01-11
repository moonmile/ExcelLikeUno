import pytest

from excellikeuno.connection import connect_calc


def _connect_or_skip():
    try:
        return connect_calc()
    except RuntimeError as exc:
        pytest.skip(f"UNO runtime not available: {exc}")


def test_add_and_remove_sheet_roundtrip():
    _, doc, _ = _connect_or_skip()
    base_names = set(doc.sheet_names)
    name = "tmp_add_sheet"
    if name in base_names:
        doc.remove_sheet(name)
    added = doc.add_sheet(name)
    try:
        assert added.name == name
        assert name in doc.sheet_names
    finally:
        doc.remove_sheet(name)
        assert set(doc.sheet_names) == base_names


def test_copy_sheet_when_supported():
    _, doc, sheet = _connect_or_skip()
    sheets = doc._sheets()
    if not hasattr(sheets, "copyByName"):
        pytest.skip("copyByName not available in this UNO runtime")

    source_name = sheet.name
    copy_name = f"{source_name}_copy_tmp"
    if copy_name in doc.sheet_names:
        doc.remove_sheet(copy_name)

    copied = doc.copy_sheet(source_name, copy_name)
    try:
        assert copied.name == copy_name
        assert copy_name in doc.sheet_names
    finally:
        doc.remove_sheet(copy_name)


def test_add_sheet_from_existing_when_supported():
    _, doc, sheet = _connect_or_skip()
    sheets = doc._sheets()
    if not hasattr(sheets, "copyByName"):
        pytest.skip("copyByName not available in this UNO runtime")

    source_name = sheet.name
    new_name = f"{source_name}_dup_tmp"
    if new_name in doc.sheet_names:
        doc.remove_sheet(new_name)

    duplicated = doc.add_sheet(new_name, from_sheet_name=source_name)
    try:
        assert duplicated.name == new_name
        assert new_name in doc.sheet_names
    finally:
        doc.remove_sheet(new_name)
