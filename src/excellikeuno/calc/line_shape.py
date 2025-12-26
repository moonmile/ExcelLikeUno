from __future__ import annotations

from typing import Any

from .shape import Shape


class LineShape(Shape):
    """Wraps com.sun.star.drawing.LineShape service."""

    @property
    def start_position(self) -> Any:
        return self._get_prop("StartPosition")

    @start_position.setter
    def start_position(self, value: Any) -> None:
        self._set_prop("StartPosition", value)

    @property
    def end_position(self) -> Any:
        return self._get_prop("EndPosition")

    @end_position.setter
    def end_position(self, value: Any) -> None:
        self._set_prop("EndPosition", value)
