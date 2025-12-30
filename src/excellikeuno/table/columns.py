from __future__ import annotations

from typing import cast

from ..core import UnoObject
from ..typing import XTableColumns


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

    insertByIndex = insert  # noqa: N815 - UNO naming alias
    removeByIndex = remove  # noqa: N815 - UNO naming alias
