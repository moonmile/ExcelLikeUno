from __future__ import annotations

from .shape import Shape


class PageShape(Shape):
    """Wraps com.sun.star.drawing.PageShape service."""

    @property
    def page_number(self) -> int:
        value = self._get_prop("PageNumber")
        return int(value) if value is not None else 0

    @page_number.setter
    def page_number(self, value: int) -> None:
        self._set_prop("PageNumber", int(value))
