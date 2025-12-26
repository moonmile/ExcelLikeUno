from __future__ import annotations

from typing import Any

from .shape import Shape


class EllipseShape(Shape):
    """Wraps com.sun.star.drawing.EllipseShape service."""

    @property
    def circle_kind(self) -> int:
        return int(self._get_prop("CircleKind"))

    @circle_kind.setter
    def circle_kind(self, value: int) -> None:
        self._set_prop("CircleKind", int(value))

    @property
    def start_angle(self) -> int:
        value = self._get_prop("StartAngle")
        return int(value) if value is not None else 0

    @start_angle.setter
    def start_angle(self, value: int) -> None:
        self._set_prop("StartAngle", int(value))

    @property
    def end_angle(self) -> int:
        value = self._get_prop("EndAngle")
        return int(value) if value is not None else 0

    @end_angle.setter
    def end_angle(self, value: int) -> None:
        self._set_prop("EndAngle", int(value))

    @property
    def circle_center(self) -> Any:
        return self._get_prop("CircleCenter")

    @circle_center.setter
    def circle_center(self, value: Any) -> None:
        self._set_prop("CircleCenter", value)

    @property
    def circle_radius(self) -> Any:
        return self._get_prop("CircleRadius")

    @circle_radius.setter
    def circle_radius(self, value: Any) -> None:
        self._set_prop("CircleRadius", value)
