from __future__ import annotations

from typing import Any

from ..typing import InterfaceNames
from .base import UnoObject


class WriterDocument(UnoObject):
    """Wraps a Writer XTextDocument."""

    @property
    def text(self) -> Any:
        """Return the document text (XText)."""
        doc = self.iface(InterfaceNames.X_TEXT_DOCUMENT)
        return doc.getText()

    def create_cursor(self) -> Any:
        """Create a text cursor at the start of the document."""
        return self.text.createTextCursor()

    def append_string(self, value: str, paragraph_break: bool = True) -> None:
        """Append string to the end; optionally add a paragraph break."""
        end = self.text.getEnd()
        self.text.insertString(end, value, False)
        if paragraph_break:
            # 0 == PARAGRAPH_BREAK
            self.text.insertControlCharacter(end, 0, False)

    def __getattr__(self, name: str) -> Any:
        """Delegate unknown attributes to the raw UNO document."""
        try:
            return getattr(self.raw, name)
        except AttributeError as exc:
            raise AttributeError(f"WriterDocument has no attribute {name}") from exc
