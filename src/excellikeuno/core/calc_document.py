from __future__ import annotations

from typing import List, cast

from ..table import Sheet
from ..typing import InterfaceNames, XSpreadsheet, XSpreadsheetDocument
from .base import UnoObject


class Document(UnoObject):
    """Wraps a Calc XSpreadsheetDocument."""

    def _sheets(self):
        doc = cast(XSpreadsheetDocument, self.iface(InterfaceNames.X_SPREADSHEET_DOCUMENT))
        return doc.getSheets()

    def sheet(self, index: int) -> Sheet:
        sheets = self._sheets()
        sheet_obj = cast(XSpreadsheet, sheets.getByIndex(index))
        return Sheet(sheet_obj)

    def sheet_by_name(self, name: str) -> Sheet:
        sheets = self._sheets()
        sheet_obj = cast(XSpreadsheet, sheets.getByName(name))
        return Sheet(sheet_obj)

    def add_sheet(self, name: str, index: int | None = None) -> Sheet:
        sheets = self._sheets()
        position = sheets.getCount() if index is None else int(index)
        sheets.insertNewByName(name, position)
        return self.sheet_by_name(name)

    def remove_sheet(self, name: str) -> None:
        sheets = self._sheets()
        sheets.removeByName(name)

    @property
    def sheet_names(self) -> List[str]:
        sheets = self._sheets()
        return list(sheets.getElementNames())

    @property
    def active_sheet(self) -> Sheet:
        doc = cast(XSpreadsheetDocument, self.iface(InterfaceNames.X_SPREADSHEET_DOCUMENT))
        controller = doc.getCurrentController()
        sheet_obj = cast(XSpreadsheet, controller.getActiveSheet())
        return Sheet(sheet_obj)
