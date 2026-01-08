from __future__ import annotations
from typing import Any, Callable, Dict
from excellikeuno.typing.calc import Color, FontSlant

class Line:
    def __init__(self, set_property: Callable[[str, Any], None], get_property: Callable[[str], Any]) -> None:
        self._set_property = set_property
        self._get_property = get_property

    @property
    def color(self) -> Color:
        return self._get_property("LineColor")

    @color.setter
    def color(self, value: Color) -> None:
        self._set_property("LineColor", value)

    @property
    def line_style(self) -> FontSlant:
        return self._get_property("LineStyle")

    @line_style.setter
    def line_style(self, value: FontSlant) -> None:
        self._set_property("LineStyle", value)