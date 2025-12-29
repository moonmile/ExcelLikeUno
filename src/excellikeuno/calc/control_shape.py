from __future__ import annotations

from typing import Any

from .shape import Shape


class ControlShape(Shape):
    """Wraps com.sun.star.drawing.ControlShape service."""

    @property
    def control(self) -> Any:
        return self._get_prop("Control")

    @control.setter
    def control(self, value: Any) -> None:
        self._set_prop("Control", value)
