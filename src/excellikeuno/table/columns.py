from __future__ import annotations

from typing import TYPE_CHECKING, cast

from ..core import UnoObject
from ..typing import XTableColumns

if TYPE_CHECKING:  # for type hints only
    from .range import TableColumn


class TableColumns(UnoObject):
    """Lightweight wrapper for XTableColumns with friendly aliases."""

    @property
    def count(self) -> int:
        cols = cast(XTableColumns, self.raw)
        return int(cols.getCount())

    def insert(self, index: int, count: int = 1) -> None:
        cols = cast(XTableColumns, self.raw)
        cols.insertByIndex(int(index), int(count))

    def remove(self, index: int, count: int = 1) -> None:
        cols = cast(XTableColumns, self.raw)
        cols.removeByIndex(int(index), int(count))

    def getByIndex(self, index: int) -> "TableColumn":
        from .range import TableColumn  # local import to avoid circular dependency

        cols = cast(XTableColumns, self.raw)
        col_obj = cols.getByIndex(int(index))
        return TableColumn(col_obj)

    insertByIndex = insert  # noqa: N815 - UNO naming alias
    removeByIndex = remove  # noqa: N815 - UNO naming alias

class TableColumn(UnoObject):
    """Lightweight wrapper for a single table column."""

    # Additional methods and properties for TableColumn can be added here