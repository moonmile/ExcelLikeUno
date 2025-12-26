from __future__ import annotations

from typing import Any

from .shape import Shape


class PolyLineShape(Shape):
    """Wraps com.sun.star.drawing.PolyLineShape service."""

    @property
    def poly_polygon(self) -> Any:
        return self._get_prop("PolyPolygon")

    @poly_polygon.setter
    def poly_polygon(self, value: Any) -> None:
        self._set_prop("PolyPolygon", value)
*** End File