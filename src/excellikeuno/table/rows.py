from __future__ import annotations

from typing import cast

from ..core import UnoObject
from ..typing import XTableRows


class TableRows(UnoObject):
    """Lightweight wrapper for XTableRows with friendly aliases."""

    @property
    def count(self) -> int:
        rows = cast(XTableRows, self.raw)
        return int(rows.getCount())

    def insert(self, index: int, count: int = 1) -> None:
        rows = cast(XTableRows, self.raw)
        rows.insertByIndex(int(index), int(count))

    def remove(self, index: int, count: int = 1) -> None:
        rows = cast(XTableRows, self.raw)
        rows.removeByIndex(int(index), int(count))

    # UNO-style aliases
    insertByIndex = insert  # noqa: N815 - UNO naming alias
    removeByIndex = remove  # noqa: N815 - UNO naming alias
