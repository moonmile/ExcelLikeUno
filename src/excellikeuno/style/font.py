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
        owner: Any | None = None,
        **kwargs: Any,
    ) -> None:
        # Allow constructing with owner only; fallback to owner's font getter/setter if provided.
        if owner is not None and getter is None:
            getter = getattr(owner, "_font_getter", None)
        if owner is not None and setter is None:
            setter = getattr(owner, "_font_setter", None)

        object.__setattr__(self, "_owner", owner)
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
        elif self._owner is not None:
            try:
                base = self._read_from_owner()
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
        elif self._owner is not None:
            try:
                self._write_to_owner(**updates)
            except Exception:
                self._buffer.update(updates)
        else:
            self._buffer.update(updates)
        return self

    # --- owner direct access helpers -------------------------------------
    def _read_from_owner(self) -> Dict[str, Any]:
        owner = getattr(self, "_owner", None)
        if owner is None:
            return {}
        # Prefer owner's dedicated getter when available (Range keeps broadcast semantics).
        og = getattr(owner, "_font_getter", None)
        if callable(og):
            try:
                return dict(og())
            except Exception:
                pass

        cp = getattr(owner, "character_properties", None)

        def _get(name: str) -> Any:
            if cp is not None:
                try:
                    return cp.get_property(name)
                except Exception:
                    # fall through to owner
                    pass
            try:
                return getattr(owner, name)
            except Exception:
                return None

        def _as_float(val: Any) -> float:
            try:
                return float(val)
            except Exception:
                return 0.0

        def _as_int(val: Any) -> int:
            try:
                return int(val)
            except Exception:
                return 0

        esc = _as_int(_get("CharEscapement"))
        backcolor = _get("CharBackColor")
        if backcolor is None:
            backcolor = _get("CellBackColor")

        return {
            "name": _get("CharFontName"),
            "size": _as_float(_get("CharHeight")),
            "bold": _as_float(_get("CharWeight")) >= 150.0,
            "italic": bool(_as_int(_get("CharPosture"))),
            "underline": _as_int(_get("CharUnderline")),
            "strikeout": _as_int(_get("CharStrikeout")),
            "color": _get("CharColor"),
            "backcolor": backcolor,
            "subscript": esc < 0,
            "superscript": esc > 0,
            "font_style": _as_int(_get("CharPosture")),
            "strikthrough": _as_int(_get("CharStrikeout")) != 0,
        }

    def _write_to_owner(self, **updates: Any) -> None:
        owner = getattr(self, "_owner", None)
        if owner is None:
            raise AttributeError("owner not set")

        os = getattr(owner, "_font_setter", None)
        if callable(os):
            os(**updates)
            return

        cp = getattr(owner, "character_properties", None)

        def _set(name: str, value: Any) -> bool:
            if cp is not None:
                try:
                    cp.set_property(name, value)
                    return True
                except Exception:
                    pass
            try:
                setattr(owner, name, value)
                return True
            except Exception:
                return False

        if "name" in updates:
            _set("CharFontName", updates["name"])
        if "size" in updates:
            size_val = float(updates["size"])
            _set("CharHeight", size_val)
            _set("CharHeightAsian", size_val)
            _set("CharHeightComplex", size_val)
        if "bold" in updates:
            _set("CharWeight", 150.0 if updates["bold"] else 100.0)
        if "italic" in updates:
            _set("CharPosture", 2 if updates["italic"] else 0)
        if "font_style" in updates:
            try:
                _set("CharPosture", int(updates["font_style"]))
            except Exception:
                pass
        if "underline" in updates:
            _set("CharUnderline", int(updates["underline"]))
        if "strikeout" in updates:
            _set("CharStrikeout", int(updates["strikeout"]))
        if "color" in updates:
            _set("CharColor", updates["color"])
        if "backcolor" in updates:
            value = updates["backcolor"]
            # Try character background first, fallback to cell background
            try:
                _set("CharBackTransparent", False)
            except Exception:
                pass
            _set("CharBackColor", value)
            _set("CellBackColor", value)
        if "subscript" in updates or "superscript" in updates:
            if updates.get("superscript"):
                _set("CharEscapement", 58)
            elif updates.get("subscript"):
                _set("CharEscapement", -25)
            else:
                _set("CharEscapement", 0)
        if "strikthrough" in updates:
            try:
                _set("CharStrikeout", 1 if updates["strikthrough"] else 0)
            except Exception:
                pass

    # 型無しプロパティ設定の互換のため
    def __getattr__(self, name: str) -> Any:  # noqa: D401
        if name.startswith("_"):
            raise AttributeError(name)
        cur = self._current()
        if name in cur:
            return cur[name]
        raise AttributeError(name)

    # 型無しプロパティ設定の互換のため
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
