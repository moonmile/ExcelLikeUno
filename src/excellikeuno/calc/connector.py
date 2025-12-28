from __future__ import annotations

from typing import Any, Union

from .shape import Shape
from ..typing import LineDash, LineStyle


class ConnectorShape(Shape):
    """Wraps com.sun.star.drawing.ConnectorShape service."""

    @property
    def start_shape(self) -> Any:
        return self._get_prop("StartShape")

    @start_shape.setter
    def start_shape(self, value: Any) -> None:
        self._set_prop("StartShape", value)

    @property
    def end_shape(self) -> Any:
        return self._get_prop("EndShape")

    @end_shape.setter
    def end_shape(self, value: Any) -> None:
        self._set_prop("EndShape", value)

    @property
    def start_glue_point_index(self) -> int:
        return int(self._get_prop("StartGluePointIndex"))

    @start_glue_point_index.setter
    def start_glue_point_index(self, value: int) -> None:
        self._set_prop("StartGluePointIndex", int(value))

    @property
    def end_glue_point_index(self) -> int:
        return int(self._get_prop("EndGluePointIndex"))

    @end_glue_point_index.setter
    def end_glue_point_index(self, value: int) -> None:
        self._set_prop("EndGluePointIndex", int(value))

    @property
    def edge_kind(self) -> int:
        return int(self._get_prop("EdgeKind"))

    @edge_kind.setter
    def edge_kind(self, value: int) -> None:
        self._set_prop("EdgeKind", int(value))

    @property
    def line_style(self) -> LineStyle:
        value = self._get_prop("LineStyle")
        try:
            return LineStyle(int(value))
        except Exception:
            return LineStyle.SOLID

    @line_style.setter
    def line_style(self, value: Union[int, LineStyle]) -> None:
        self._set_prop("LineStyle", int(LineStyle(value)))

    @property
    def line_start_name(self) -> str:
        return str(self._get_prop("LineStartName"))

    @line_start_name.setter
    def line_start_name(self, value: str) -> None:
        self._set_prop("LineStartName", value)

    @property
    def line_end_name(self) -> str:
        return str(self._get_prop("LineEndName"))

    @line_end_name.setter
    def line_end_name(self, value: str) -> None:
        self._set_prop("LineEndName", value)

    @property
    def line_start_center(self) -> bool:
        return bool(self._get_prop("LineStartCenter"))

    @line_start_center.setter
    def line_start_center(self, value: bool) -> None:
        self._set_prop("LineStartCenter", bool(value))

    @property
    def line_end_center(self) -> bool:
        return bool(self._get_prop("LineEndCenter"))

    @line_end_center.setter
    def line_end_center(self, value: bool) -> None:
        self._set_prop("LineEndCenter", bool(value))

    @property
    def line_dash(self) -> LineDash:
        return self._get_prop("LineDash")

    @line_dash.setter
    def line_dash(self, value: LineDash) -> None:
        self._set_prop("LineDash", value)
