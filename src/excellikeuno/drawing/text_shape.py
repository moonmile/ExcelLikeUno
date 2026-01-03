from __future__ import annotations

from .shape import Shape


class TextShape(Shape):
    """Wraps com.sun.star.drawing.TextShape service."""

    @property
    def string(self) -> str:
        try:
            value = self._get_prop("String")
            return str(value) if value is not None else ""
        except Exception:
            try:
                return str(self.raw.getText().getString())  # type: ignore[attr-defined]
            except Exception:
                return ""

    @string.setter
    def string(self, value: str) -> None:
        try:
            self._set_prop("String", value)
            return
        except Exception:
            pass
        try:
            self.raw.setString(value)  # type: ignore[attr-defined]
            return
        except Exception:
            pass
        try:
            self.raw.getText().setString(value)  # type: ignore[attr-defined]
        except Exception:
            # Best effort; swallow if unsupported
            pass

    @property
    def String(self) -> str:
        return self.string

    @String.setter
    def String(self, value: str) -> None:
        self.string = value

