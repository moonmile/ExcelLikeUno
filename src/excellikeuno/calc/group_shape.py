from __future__ import annotations

from typing import List, cast

from ..core import InterfaceNames, UnoObject
from ..typing import XShapes
from .shape import Shape


class GroupShape(Shape):
    """Wraps com.sun.star.drawing.GroupShape service."""

    def _shapes(self) -> XShapes:
        return cast(XShapes, self.iface(InterfaceNames.X_SHAPES))

    def shapes(self) -> List[Shape]:
        xs = self._shapes()
        return [Shape(xs.getByIndex(i)) for i in range(xs.getCount())]

    def shape(self, index: int) -> Shape:
        xs = self._shapes()
        return Shape(xs.getByIndex(index))

    def add(self, shape: Shape) -> None:
        xs = self._shapes()
        xs.add(shape.raw)

    def remove(self, shape: Shape) -> None:
        xs = self._shapes()
        xs.remove(shape.raw)
