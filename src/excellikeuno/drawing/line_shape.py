from __future__ import annotations

from excellikeuno.typing.structs import Point, Size
from .shape import Shape


class LineShape(Shape):
    """Wraps com.sun.star.drawing.LineShape service."""

    @property
    def start_position(self) -> Point:
        return self.Position

    @start_position.setter
    def start_position(self, value: Point) -> None:
        self.Position = value

    @property
    def end_position(self) -> Point:
        return Point( self.Position.X + self.Size.Width, self.Position.Y + self.Size.Height)

    @end_position.setter
    def end_position(self, value: Point) -> None:
        # Convert the desired end point into a size delta so UNO gets a Size struct
        width = value.X - self.Position.X
        height = value.Y - self.Position.Y
        self.Size = Size(width, height)
