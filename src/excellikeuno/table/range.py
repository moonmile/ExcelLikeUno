from __future__ import annotations

from typing import Any, cast

from ..core import UnoObject
from ..typing import InterfaceNames, XCellRange, XTableRows, XTableColumns
from .cell import Cell
from .columns import TableColumns
from .rows import TableRows


class Range(UnoObject):
    """Wraps a UNO cell range and exposes cell-level access."""

    def cell(self, column: int, row: int) -> Cell:
        rng = cast(XCellRange, self.iface(InterfaceNames.X_CELL_RANGE))
        return Cell(rng.getCellByPosition(column, row))

    def subrange(self, start_column: int, start_row: int, end_column: int, end_row: int) -> "Range":
        rng = cast(XCellRange, self.iface(InterfaceNames.X_CELL_RANGE))
        return Range(rng.getCellRangeByPosition(start_column, start_row, end_column, end_row))

    # Convenience aliases matching spreadsheet terminology
    getCellByPosition = cell  # noqa: N815 - UNO-style alias
    getCellRangeByPosition = subrange  # noqa: N815 - UNO-style alias

    @property
    def rows(self) -> XTableRows:
        rng = cast(XCellRange, self.iface(InterfaceNames.X_CELL_RANGE))
        return TableRows(rng.getRows())

    @property
    def columns(self) -> XTableColumns:
        rng = cast(XCellRange, self.iface(InterfaceNames.X_CELL_RANGE))
        return TableColumns(rng.getColumns())

    def __iter__(self):
        # Iterate rows then columns by yielding Cell wrappers
        rng = cast(XCellRange, self.iface(InterfaceNames.X_CELL_RANGE))
        # Attempt to derive bounds; UNO ranges lack a simple API, so rely on getCellByPosition until it fails
        # This keeps iteration defensive without assuming size metadata.
        col = 0
        row = 0
        while True:  # pragma: no cover - iteration is best-effort
            try:
                yield self.cell(col, row)
                col += 1
            except Exception:
                # move to next row when column lookup fails
                col = 0
                row += 1
                try:
                    _ = rng.getCellByPosition(col, row)
                except Exception:
                    break

    def __repr__(self) -> str:  # pragma: no cover - debugging helper
        return f"Range({self.raw!r})"
