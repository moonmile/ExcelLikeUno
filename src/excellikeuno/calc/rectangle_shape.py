from __future__ import annotations

from .shape import Shape


class RectangleShape(Shape):
    """Wraps com.sun.star.drawing.RectangleShape service."""

    @property
    def corner_radius(self) -> int:
        return int(self._get_prop("CornerRadius"))

    @corner_radius.setter
    def corner_radius(self, value: int) -> None:
        self._set_prop("CornerRadius", int(value))

