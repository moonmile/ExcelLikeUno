from __future__ import annotations

from typing import Any

from .shape import Shape


class MeasureShape(Shape):
    """Wraps com.sun.star.drawing.MeasureShape service."""

    @property
    def measure_text(self) -> str:
        value = self._get_prop("MeasureText")
        return str(value) if value is not None else ""

    @measure_text.setter
    def measure_text(self, value: str) -> None:
        self._set_prop("MeasureText", value)

    @property
    def measure_value(self) -> Any:
        return self._get_prop("MeasureValue")

    @measure_value.setter
    def measure_value(self, value: Any) -> None:
        self._set_prop("MeasureValue", value)
