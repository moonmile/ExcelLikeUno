from __future__ import annotations

from typing import Any, Callable, Dict

from excellikeuno.typing import LineDash, LineStyle
from excellikeuno.typing.calc import Color


class Line:
    """Line-property proxy following the Font/Borders apply/buffer pattern."""

    def __init__(
        self,
        getter: Callable[[], Dict[str, Any]] | None = None,
        setter: Callable[..., None] | None = None,
        owner: Any | None = None,
        **kwargs: Any,
    ) -> None:
        if owner is not None and getter is None:
            getter = getattr(owner, "_line_getter", None)
        if owner is not None and setter is None:
            setter = getattr(owner, "_line_setter", None)

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
        if self._buffer:
            base.update(self._buffer)
        return base

    def apply(self, **kwargs: Any) -> "Line":
        updates = {k: v for k, v in kwargs.items() if v is not None}
        if not updates:
            return self
        if self._setter:
            try:
                self._setter(**updates)
            except Exception:
                self._buffer.update(updates)
            else:
                self._buffer.update(updates)
        elif self._owner is not None:
            try:
                self._write_to_owner(**updates)
            except Exception:
                self._buffer.update(updates)
            else:
                self._buffer.update(updates)
        else:
            self._buffer.update(updates)
        return self

    def _value_from_owner(self, field: str) -> Any:
        owner = getattr(self, "_owner", None)
        if owner is None:
            raise AttributeError("owner not set")

        lp = getattr(owner, "line_properties", None)

        def _get(name: str) -> Any:
            if lp is not None:
                try:
                    return lp.get_property(name)
                except Exception:
                    pass
            try:
                return getattr(owner, name)
            except Exception:
                return None

        if field == "color":
            return _get("LineColor")
        if field == "line_style":
            try:
                return LineStyle(int(_get("LineStyle")))
            except Exception:
                return LineStyle(0)
        if field == "dash":
            return _get("LineDash")
        if field == "dash_name":
            return _get("LineDashName")
        if field == "transparence":
            try:
                return int(_get("LineTransparence"))
            except Exception:
                return 0
        if field in {"width", "weight"}:
            try:
                return int(_get("LineWidth"))
            except Exception:
                return 0

        cur = self._current()
        if field in cur:
            return cur[field]
        raise AttributeError(field)

    def _read_from_owner(self) -> Dict[str, Any]:
        owner = getattr(self, "_owner", None)
        if owner is None:
            return {}

        og = getattr(owner, "_line_getter", None)
        if callable(og):
            try:
                return dict(og())
            except Exception:
                pass

        lp = getattr(owner, "line_properties", None)

        def _get(name: str) -> Any:
            if lp is not None:
                try:
                    return lp.get_property(name)
                except Exception:
                    pass
            try:
                return getattr(owner, name)
            except Exception:
                return None

        def _as_int(val: Any) -> int:
            try:
                return int(val)
            except Exception:
                return 0

        try:
            ls_val = LineStyle(int(_get("LineStyle")))
        except Exception:
            ls_val = LineStyle(0)

        return {
            "color": _get("LineColor"),
            "line_style": ls_val,
            "dash": _get("LineDash"),
            "dash_name": _get("LineDashName"),
            "transparence": _as_int(_get("LineTransparence")),
            "width": _as_int(_get("LineWidth")),
        }

    def _write_to_owner(self, **updates: Any) -> None:
        owner = getattr(self, "_owner", None)
        if owner is None:
            raise AttributeError("owner not set")

        os = getattr(owner, "_line_setter", None)
        if callable(os):
            os(**updates)
            return

        lp = getattr(owner, "line_properties", None)

        def _set(name: str, value: Any) -> bool:
            if lp is not None:
                try:
                    lp.set_property(name, value)
                    return True
                except Exception:
                    pass
            try:
                setattr(owner, name, value)
                return True
            except Exception:
                return False

        if "color" in updates:
            _set("LineColor", int(updates["color"]))
        if "line_style" in updates:
            _set("LineStyle", int(updates["line_style"]))
        if "dash" in updates:
            _set("LineDash", updates["dash"])
        if "dash_name" in updates:
            _set("LineDashName", updates["dash_name"])
        if "transparence" in updates:
            _set("LineTransparence", int(updates["transparence"]))
        if "width" in updates:
            _set("LineWidth", int(updates["width"]))
        if "weight" in updates:
            _set("LineWidth", int(updates["weight"]))

    @property
    def color(self) -> Color:
        cur = self._current()
        if "color" in cur:
            try:
                return Color(cur["color"])
            except Exception:
                pass
        try:
            return Color(self._value_from_owner("color"))
        except Exception:
            return Color(cur.get("color", 0))

    @color.setter
    def color(self, value: Color) -> None:
        self.apply(color=value)

    @property
    def line_style(self) -> LineStyle:
        cur = self._current()
        if "line_style" in cur:
            try:
                return LineStyle(int(cur["line_style"]))
            except Exception:
                pass
        try:
            return LineStyle(int(self._value_from_owner("line_style")))
        except Exception:
            try:
                return LineStyle(int(cur.get("line_style", 0)))
            except Exception:
                return LineStyle(0)

    @line_style.setter
    def line_style(self, value: LineStyle | int) -> None:
        self.apply(line_style=int(value))

    @property
    def dash(self) -> LineDash:
        cur = self._current()
        if "dash" in cur:
            return cur.get("dash")
        try:
            return self._value_from_owner("dash")
        except Exception:
            return cur.get("dash")

    @dash.setter
    def dash(self, value: LineDash) -> None:
        self.apply(dash=value)

    @property
    def dash_name(self) -> str:
        cur = self._current()
        if "dash_name" in cur:
            try:
                return str(cur.get("dash_name", ""))
            except Exception:
                pass
        try:
            return str(self._value_from_owner("dash_name"))
        except Exception:
            return str(cur.get("dash_name", ""))

    @dash_name.setter
    def dash_name(self, value: str) -> None:
        self.apply(dash_name=value)

    @property
    def transparence(self) -> int:
        cur = self._current()
        if "transparence" in cur:
            try:
                return int(cur.get("transparence", 0))
            except Exception:
                pass
        try:
            return int(self._value_from_owner("transparence"))
        except Exception:
            return int(cur.get("transparence", 0))

    @transparence.setter
    def transparence(self, value: int) -> None:
        self.apply(transparence=value)

    @property
    def width(self) -> int:
        cur = self._current()
        if "width" in cur:
            try:
                return int(cur.get("width", 0))
            except Exception:
                pass
        try:
            return int(self._value_from_owner("width"))
        except Exception:
            return int(cur.get("width", 0))

    @width.setter
    def width(self, value: int) -> None:
        self.apply(width=value)

    @property
    def weight(self) -> int:
        return self.width

    @weight.setter
    def weight(self, value: int) -> None:
        self.apply(weight=value)

    def items(self):  # pragma: no cover
        return self._current().items()

    def __getattr__(self, name: str) -> Any:
        cur = self._current()
        if name in cur:
            return cur[name]
        raise AttributeError(name)

    def __setattr__(self, name: str, value: Any) -> None:
        if name.startswith("_"):
            object.__setattr__(self, name, value)
            return
        self.apply(**{name: value})