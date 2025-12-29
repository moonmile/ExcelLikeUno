from __future__ import annotations

from typing import Any

from .shape import Shape


class PolyPolygonBezierShape(Shape):
    """Wraps com.sun.star.drawing.PolyPolygonBezierShape service."""

    @property
    def poly_polygon_bezier(self) -> Any:
        return self._get_prop("PolyPolygonBezier")

    @poly_polygon_bezier.setter
    def poly_polygon_bezier(self, value: Any) -> None:
        self._set_prop("PolyPolygonBezier", value)
