from __future__ import annotations

from typing import Any

from .shape import Shape


class CustomShape(Shape):
    """Wraps com.sun.star.drawing.CustomShape service."""

    @property
    def custom_shape_engine(self) -> str:
        value = self._get_prop("CustomShapeEngine")
        return str(value) if value is not None else ""

    @custom_shape_engine.setter
    def custom_shape_engine(self, value: str) -> None:
        self._set_prop("CustomShapeEngine", value)

    @property
    def custom_shape_data(self) -> Any:
        return self._get_prop("CustomShapeData")

    @custom_shape_data.setter
    def custom_shape_data(self, value: Any) -> None:
        self._set_prop("CustomShapeData", value)
