from __future__ import annotations

from typing import Any, Callable, Dict


class Font:
    """Proxy-style font wrapper.

    When constructed with getter/setter callables, attribute access updates UNO immediately.
    When constructed without them, it behaves as a simple config holder (used for assignment).
    """

    def __init__(
        self,
        getter: Callable[[], Dict[str, Any]] | None = None,
        setter: Callable[..., None] | None = None,
        **kwargs: Any,
    ) -> None:
        object.__setattr__(self, "_getter", getter)
        object.__setattr__(self, "_setter", setter)
        object.__setattr__(self, "_buffer", {})
        if kwargs:
            self.apply(**kwargs)

    def _current(self) -> Dict[str, Any]:
        base: Dict[str, Any] = {}
        if self._getter:
            try:
                base = dict(self._getter())
            except Exception:
                base = {}
        # buffered values take precedence
        if self._buffer:
            base.update(self._buffer)
        return base

    def apply(self, **kwargs: Any) -> "Font":
        # Remove None updates so they don't clobber existing values silently
        updates = {k: v for k, v in kwargs.items() if v is not None}
        if not updates:
            return self
        if self._setter:
            try:
                self._setter(**updates)
            except Exception:
                # fall back to local buffer when setter fails
                self._buffer.update(updates)
        else:
            self._buffer.update(updates)
        return self

    def __getattr__(self, name: str) -> Any:  # noqa: D401
        if name.startswith("_"):
            raise AttributeError(name)
        cur = self._current()
        if name in cur:
            return cur[name]
        raise AttributeError(name)

    def __setattr__(self, name: str, value: Any) -> None:  # noqa: D401
        if name.startswith("_"):
            object.__setattr__(self, name, value)
            return
        self.apply(**{name: value})

    # Allow iteration over items when needed (e.g., for debug)
    def items(self):  # pragma: no cover - helper
        return self._current().items()

    # --- typed convenience properties for IDE completion -----------------
    @property
    def name(self) -> Any:
        return self._current().get("name")

    @name.setter
    def name(self, value: Any) -> None:
        self.apply(name=value)

    @property
    def size(self) -> Any:
        return self._current().get("size")

    @size.setter
    def size(self, value: Any) -> None:
        self.apply(size=value)

    @property
    def bold(self) -> Any:
        return self._current().get("bold")

    @bold.setter
    def bold(self, value: Any) -> None:
        self.apply(bold=value)

    @property
    def italic(self) -> Any:
        return self._current().get("italic")

    @italic.setter
    def italic(self, value: Any) -> None:
        self.apply(italic=value)

    @property
    def underline(self) -> Any:
        return self._current().get("underline")

    @underline.setter
    def underline(self, value: Any) -> None:
        self.apply(underline=value)

    @property
    def strikeout(self) -> Any:
        return self._current().get("strikeout")

    @strikeout.setter
    def strikeout(self, value: Any) -> None:
        self.apply(strikeout=value)

    @property
    def color(self) -> Any:
        return self._current().get("color")

    @color.setter
    def color(self, value: Any) -> None:
        self.apply(color=value)

    @property
    def backcolor(self) -> Any:
        return self._current().get("backcolor")

    @backcolor.setter
    def backcolor(self, value: Any) -> None:
        self.apply(backcolor=value)

    @property
    def subscript(self) -> Any:
        return self._current().get("subscript")

    @subscript.setter
    def subscript(self, value: Any) -> None:
        self.apply(subscript=value)

    @property
    def superscript(self) -> Any:
        return self._current().get("superscript")

    @superscript.setter
    def superscript(self, value: Any) -> None:
        self.apply(superscript=value)

    @property
    def font_style(self) -> Any:
        return self._current().get("font_style")

    @font_style.setter
    def font_style(self, value: Any) -> None:
        self.apply(font_style=value)

    @property
    def strikthrough(self) -> Any:
        return self._current().get("strikthrough")

    @strikthrough.setter
    def strikthrough(self, value: Any) -> None:
        self.apply(strikthrough=value)


__all__ = ["Font"]
