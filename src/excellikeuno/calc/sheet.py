from __future__ import annotations

from typing import cast

from ..core import InterfaceNames, UnoObject
from ..typing import XSpreadsheet
from .cell import Cell


class Sheet(UnoObject):
    def cell(self, column: int, row: int) -> Cell:
        sheet = cast(XSpreadsheet, self.iface(InterfaceNames.X_SPREADSHEET))
        return Cell(sheet.getCellByPosition(column, row))
