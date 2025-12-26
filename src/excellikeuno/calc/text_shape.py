from __future__ import annotations

from .shape import Shape


class TextShape(Shape):
    """Wraps com.sun.star.drawing.TextShape service."""

    @property
    def string(self) -> str:
        value = self._get_prop("String")
        return str(value) if value is not None else ""

    @string.setter
    def string(self, value: str) -> None:
        self._set_prop("String", value)
*** End File