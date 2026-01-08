from __future__ import annotations

from typing import Any, Callable, Dict

from excellikeuno.typing.calc import Color, FontSlant


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
    def _value_from_owner(self, field: str) -> Any:
        owner = getattr(self, "_owner", None)
        if owner is None:
            raise AttributeError("owner not set")

        cp = getattr(owner, "character_properties", None)

        def _get_prop(name: str) -> Any:
            if cp is not None:
                try:
                    return cp.get_property(name)
                except Exception:
                    pass
            return getattr(owner, name)

        if field == "name":
            return _get_prop("CharFontName")
        if field == "size":
            try:
                return float(_get_prop("CharHeight"))
            except Exception:
                return 0.0
        if field == "bold":
            try:
                return float(_get_prop("CharWeight")) >= 150.0
            except Exception:
                return False
        if field in {"italic", "font_style"}:
            try:
                posture = int(_get_prop("CharPosture"))
                return posture if field == "font_style" else bool(posture)
            except Exception:
                return 0 if field == "font_style" else False
        if field == "underline":
            try:
                return int(_get_prop("CharUnderline"))
            except Exception:
                return 0
        if field == "strikeout":
            try:
                return int(_get_prop("CharStrikeout"))
            except Exception:
                return 0
        if field == "color":
            try:
                return _get_prop("CharColor")
            except Exception:
                return None
        if field == "backcolor":
            val = None
            try:
                val = _get_prop("CharBackColor")
            except Exception:
                val = None
            if val is None:
                try:
                    val = _get_prop("CellBackColor")
                except Exception:
                    val = None
            return val
        if field in {"subscript", "superscript", "strikthrough"}:
            try:
                esc = int(_get_prop("CharEscapement"))
            except Exception:
                esc = 0
            if field == "subscript":
                return esc < 0
            if field == "superscript":
                return esc > 0
            if field == "strikthrough":
                try:
                    return int(_get_prop("CharStrikeout")) != 0
                except Exception:
                    return False
        # Fallback to cached mapping
        cur = self._current()
        if field in cur:
            return cur[field]
        raise AttributeError(field)

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
        if name in self._buffer:
            return self._buffer[name]
        try:
            return self._value_from_owner(name)
        except Exception:
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
    def name(self) -> str:
        if "name" in self._buffer:
            return self._buffer.get("name")
        try:
            return self._value_from_owner("name")
        except Exception:
            return self._current().get("name")

    @name.setter
    def name(self, value: str) -> None:
        self.apply(name=value)

    @property
    def size(self) -> float:
        if "size" in self._buffer:
            return self._buffer.get("size")
        try:
            return self._value_from_owner("size")
        except Exception:
            return self._current().get("size")

    @size.setter
    def size(self, value: float) -> None:
        self.apply(size=value)

    @property
    def bold(self) -> bool:
        if "bold" in self._buffer:
            return self._buffer.get("bold")
        try:
            return self._value_from_owner("bold")
        except Exception:
            return self._current().get("bold")

    @bold.setter
    def bold(self, value: bool) -> None:
        self.apply(bold=value)

    @property
    def italic(self) -> bool:
        if "italic" in self._buffer:
            return self._buffer.get("italic")
        try:
            return self._value_from_owner("italic")
        except Exception:
            return self._current().get("italic")

    @italic.setter
    def italic(self, value: bool) -> None:
        self.apply(italic=value)

    @property
    def underline(self) -> bool:
        if "underline" in self._buffer:
            return self._buffer.get("underline")
        try:
            return self._value_from_owner("underline")
        except Exception:
            return self._current().get("underline")

    @underline.setter
    def underline(self, value: bool) -> None:
        self.apply(underline=value)

    @property
    def strikeout(self) -> bool:
        if "strikeout" in self._buffer:
            return self._buffer.get("strikeout")
        try:
            return self._value_from_owner("strikeout")
        except Exception:
            return self._current().get("strikeout")

    @strikeout.setter
    def strikeout(self, value: bool) -> None:
        self.apply(strikeout=value)

    @property
    def color(self) -> Color:
        if "color" in self._buffer:
            return self._buffer.get("color")
        try:
            return self._value_from_owner("color")
        except Exception:
            return self._current().get("color")

    @color.setter
    def color(self, value: Color) -> None:
        self.apply(color=value)

    @property
    def backcolor(self) -> Color:
        if "backcolor" in self._buffer:
            return self._buffer.get("backcolor")
        try:
            return self._value_from_owner("backcolor")
        except Exception:
            return self._current().get("backcolor")

    @backcolor.setter
    def backcolor(self, value: Color) -> None:
        self.apply(backcolor=value)

    @property
    def subscript(self) -> bool:
        if "subscript" in self._buffer:
            return self._buffer.get("subscript")
        try:
            return self._value_from_owner("subscript")
        except Exception:
            return self._current().get("subscript")

    @subscript.setter
    def subscript(self, value: bool) -> None:
        self.apply(subscript=value)

    @property
    def superscript(self) -> bool:
        if "superscript" in self._buffer:
            return self._buffer.get("superscript")
        try:
            return self._value_from_owner("superscript")
        except Exception:
            return self._current().get("superscript")

    @superscript.setter
    def superscript(self, value: bool) -> None:
        self.apply(superscript=value)

    @property
    def font_style(self) -> FontSlant:
        if "font_style" in self._buffer:
            return self._buffer.get("font_style")
        try:
            return self._value_from_owner("font_style")
        except Exception:
            return self._current().get("font_style")

    @font_style.setter
    def font_style(self, value: FontSlant) -> None:
        self.apply(font_style=value)

    @property
    def strikthrough(self) -> int:
        if "strikthrough" in self._buffer:
            return self._buffer.get("strikthrough")
        try:
            return self._value_from_owner("strikthrough")
        except Exception:
            return self._current().get("strikthrough")

    @strikthrough.setter
    def strikthrough(self, value: int) -> None:
        self.apply(strikthrough=value)


__all__ = ["Font"]
