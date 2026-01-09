from __future__ import annotations

from typing import Any, Callable, Dict

from excellikeuno.typing import BorderLine, BorderLine2
from excellikeuno.typing.calc import BorderLineStyle, Color


class BorderStyle:
    """Per-side border style proxy with field-level properties."""

    def __init__(self, parent: "Borders" | None = None, side: str = "all", **kwargs: Any) -> None:
        object.__setattr__(self, "_parent", parent)
        object.__setattr__(self, "_side", side)
        if parent is None:
            object.__setattr__(self, "_local_line", BorderLine2(LineStyle=0))
        if kwargs:
            self._apply(**kwargs)

    def _line(self) -> BorderLine | BorderLine2:
        parent = getattr(self, "_parent", None)
        if parent is None:
            return getattr(self, "_local_line")

        cur = parent._current()
        if self._side == "all":
            for key in ("top", "bottom", "left", "right"):
                if key in cur:
                    return parent._clone_line(cur[key])
            return BorderLine2(LineStyle=0)
        if self._side in cur:
            return parent._clone_line(cur[self._side])
        return BorderLine2(LineStyle=0)

    def _apply(self, **fields: Any) -> None:
        parent = getattr(self, "_parent", None)
        line = parent._clone_line(self._line()) if parent is not None else self._line()

        def _as_int(val: Any, default: int = 0) -> int:
            try:
                return int(val)
            except Exception:
                return default

        # When the underlying struct lacks LineStyle/LineWidth (BorderLine), upgrade to BorderLine2
        if "line_style" in fields and not hasattr(line, "LineStyle"):
            line = BorderLine2(
                Color=_as_int(getattr(line, "Color", 0)),
                InnerLineWidth=_as_int(getattr(line, "InnerLineWidth", 0)),
                OuterLineWidth=_as_int(getattr(line, "OuterLineWidth", 0)),
                LineDistance=_as_int(getattr(line, "LineDistance", 0)),
                LineStyle=_as_int(fields["line_style"]),
                LineWidth=_as_int(getattr(line, "LineWidth", getattr(line, "OuterLineWidth", 0))),
            )

        def _set(name: str, value: Any) -> None:
            try:
                setattr(line, name, value)
            except Exception:
                pass

        if "color" in fields:
            _set("Color", int(fields["color"]))
        if "inner_width" in fields:
            _set("InnerLineWidth", int(fields["inner_width"]))
        if "outer_width" in fields:
            val = int(fields["outer_width"])
            _set("OuterLineWidth", val)
            _set("LineWidth", val)
        if "distance" in fields:
            _set("LineDistance", int(fields["distance"]))
        if "line_style" in fields:
            _set("LineStyle", int(fields["line_style"]))
        if "width" in fields:
            val = int(fields["width"])
            _set("OuterLineWidth", val)
            _set("LineWidth", val)
        if "weight" in fields:
            val = int(fields["weight"])
            _set("OuterLineWidth", val)
            _set("LineWidth", val)

        if parent is None:
            object.__setattr__(self, "_local_line", line)
            return

        parent.apply(**{self._side: line})

    @property
    def color(self) -> Color:
        return int(getattr(self._line(), "Color", 0))

    @color.setter
    def color(self, value: Color) -> None:
        self._apply(color=value)

    @property
    def inner_width(self) -> int:
        return int(getattr(self._line(), "InnerLineWidth", 0))

    @inner_width.setter
    def inner_width(self, value: int) -> None:
        self._apply(inner_width=value)

    @property
    def outer_width(self) -> int:
        line = self._line()
        try:
            return int(getattr(line, "OuterLineWidth", 0))
        except Exception:
            return 0

    @outer_width.setter
    def outer_width(self, value: int) -> None:
        self._apply(outer_width=value)

    @property
    def width(self) -> int:
        return self.outer_width

    @width.setter
    def width(self, value: int) -> None:
        self._apply(width=value)

    @property
    def weight(self) -> int:
        return self.outer_width

    @weight.setter
    def weight(self, value: int) -> None:
        self._apply(weight=value)

    @property
    def distance(self) -> int:
        return int(getattr(self._line(), "LineDistance", 0))

    @distance.setter
    def distance(self, value: int) -> None:
        self._apply(distance=value)

    @property
    def line_style(self) -> BorderLineStyle:
        val = int(getattr(self._line(), "LineStyle", 0))
        try:
            width = int(getattr(self._line(), "OuterLineWidth", getattr(self._line(), "LineWidth", 0)) or 0)
        except Exception:
            width = 0
        try:
            color = int(getattr(self._line(), "Color", 0))
        except Exception:
            color = 0

        if val == int(BorderLineStyle.NONE):
            return BorderLineStyle.NONE
        if val == 0 and width == 0:
            return BorderLineStyle.NONE
        if val == 0 and color == 0xFFFFFF and width <= 1:
            # Treat zero-width, white line with no style as cleared.
            return BorderLineStyle.NONE
        return val

    @line_style.setter
    def line_style(self, value: BorderLineStyle) -> None:
        self._apply(line_style=value)

    @property
    def line(self) -> BorderLine | BorderLine2:
        return self._line()

    def __getattr__(self, name: str) -> Any:
        line = self._line()
        if hasattr(line, name):
            return getattr(line, name)
        raise AttributeError(name)

    def __setattr__(self, name: str, value: Any) -> None:
        if name.startswith("_"):
            object.__setattr__(self, name, value)
            return
        if name in {"Color", "InnerLineWidth", "OuterLineWidth", "LineDistance", "LineStyle", "LineWidth"}:
            # direct struct-like fields
            self._apply(**{name.lower(): value})
            return
        if name in {"color", "inner_width", "outer_width", "distance", "line_style", "width", "weight"}:
            getattr(type(self), name).fset(self, value)  # type: ignore[arg-type]
            return
        raise AttributeError(name)


class Borders:
    """Proxy-style border wrapper for cell/range borders."""

    def __init__(
        self,
        getter: Callable[[], Dict[str, Any]] | None = None,
        setter: Callable[..., None] | None = None,
        owner: Any | None = None,
        **kwargs: Any,
    ) -> None:
        if owner is not None and getter is None:
            getter = getattr(owner, "_border_getter", None)
        if owner is not None and setter is None:
            setter = getattr(owner, "_border_setter", None)

        object.__setattr__(self, "_owner", owner)
        object.__setattr__(self, "_getter", getter)
        object.__setattr__(self, "_setter", setter)
        object.__setattr__(self, "_buffer", {})
        if kwargs:
            self.apply(**kwargs)

    def _clone_line(self, value: Any) -> BorderLine | BorderLine2:
        def _as_int(val: Any, default: int = 0) -> int:
            try:
                return int(val)
            except Exception:
                return default

        if isinstance(value, BorderLine2):
            return BorderLine2(
                Color=_as_int(getattr(value, "Color", 0)),
                InnerLineWidth=_as_int(getattr(value, "InnerLineWidth", 0)),
                OuterLineWidth=_as_int(getattr(value, "OuterLineWidth", 0)),
                LineDistance=_as_int(getattr(value, "LineDistance", 0)),
                LineStyle=_as_int(getattr(value, "LineStyle", 0)),
                LineWidth=_as_int(getattr(value, "LineWidth", getattr(value, "OuterLineWidth", 0))),
            )
        if isinstance(value, BorderLine):
            return BorderLine(
                Color=_as_int(getattr(value, "Color", 0)),
                InnerLineWidth=_as_int(getattr(value, "InnerLineWidth", 0)),
                OuterLineWidth=_as_int(getattr(value, "OuterLineWidth", 0)),
                LineDistance=_as_int(getattr(value, "LineDistance", 0)),
            )
        try:
            return BorderLine(
                Color=_as_int(getattr(value, "Color", 0)),
                InnerLineWidth=_as_int(getattr(value, "InnerLineWidth", 0)),
                OuterLineWidth=_as_int(getattr(value, "OuterLineWidth", 0)),
                LineDistance=_as_int(getattr(value, "LineDistance", 0)),
            )
        except Exception:
            return BorderLine()

    def _line_from_dict(self, value: Dict[str, Any]) -> BorderLine | BorderLine2:
        color = value.get("color", value.get("Color", 0))
        inner = value.get("inner_width", value.get("InnerLineWidth", 0))
        outer = value.get("outer_width", value.get("OuterLineWidth", value.get("width", value.get("weight", 0))))
        distance = value.get("distance", value.get("LineDistance", 0))
        line_style = value.get("line_style", value.get("LineStyle", 0))
        line_width = value.get("line_width", value.get("LineWidth", outer))
        try:
            return BorderLine2(
                Color=int(color),
                InnerLineWidth=int(inner),
                OuterLineWidth=int(outer),
                LineDistance=int(distance),
                LineStyle=int(line_style),
                LineWidth=int(line_width),
            )
        except Exception:
            return BorderLine(Color=int(color), InnerLineWidth=int(inner), OuterLineWidth=int(outer), LineDistance=int(distance))

    def _normalize_updates(self, updates: Dict[str, Any]) -> Dict[str, BorderLine | BorderLine2]:
        normalized = {k: v for k, v in updates.items() if v is not None}
        if "all" in normalized:
            all_line = self._clone_line(normalized.pop("all"))
            for key in ("top", "bottom", "left", "right"):
                normalized.setdefault(key, self._clone_line(all_line))
        if "around" in normalized:
            value = normalized.pop("around")
            if isinstance(value, dict):
                normalized["around"] = self._line_from_dict(value)
            elif isinstance(value, BorderStyle):
                normalized["around"] = self._clone_line(value.line)
            else:
                normalized["around"] = self._clone_line(value)
        if "inner" in normalized:
            value = normalized.pop("inner")
            if isinstance(value, dict):
                normalized["inner"] = self._line_from_dict(value)
            elif isinstance(value, BorderStyle):
                normalized["inner"] = self._clone_line(value.line)
            else:
                normalized["inner"] = self._clone_line(value)
        for key, value in list(normalized.items()):
            if isinstance(value, dict):
                normalized[key] = self._line_from_dict(value)
            elif isinstance(value, BorderStyle):
                normalized[key] = self._clone_line(value.line)
            else:
                normalized[key] = self._clone_line(value)
        return normalized

    def _current(self) -> Dict[str, BorderLine | BorderLine2]:
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
        normalized = {k: self._clone_line(v) for k, v in base.items()}
        if self._buffer:
            normalized.update({k: self._clone_line(v) for k, v in self._buffer.items()})
        return normalized

    def apply(self, **kwargs: Any) -> "Borders":
        updates = self._normalize_updates(kwargs)
        if not updates:
            return self
        if self._setter:
            try:
                self._setter(**updates)
                self._buffer.update(updates)
                return self
            except Exception:
                self._buffer.update(updates)
                return self
        if self._owner is not None:
            try:
                self._write_to_owner(**updates)
                # Keep a local buffer copy so fields unsupported by UNO (e.g., line_style on BorderLine) remain readable.
                self._buffer.update(updates)
                return self
            except Exception:
                self._buffer.update(updates)
                return self
        self._buffer.update(updates)
        return self

    def _read_from_owner(self) -> Dict[str, BorderLine | BorderLine2]:
        owner = getattr(self, "_owner", None)
        if owner is None:
            return {}

        getter = getattr(owner, "_border_getter", None)
        if callable(getter):
            try:
                values = getter()
                return {k: self._clone_line(v) for k, v in values.items()}
            except Exception:
                pass

        attr_map = {
            "top": "TopBorder",
            "bottom": "BottomBorder",
            "left": "LeftBorder",
            "right": "RightBorder",
        }
        result: Dict[str, BorderLine | BorderLine2] = {}
        for key, attr in attr_map.items():
            try:
                base_val = getattr(owner, attr)
                alt_attr = f"{attr}2"
                try:
                    alt_val = getattr(owner, alt_attr)
                    if hasattr(alt_val, "LineStyle"):
                        base_val = alt_val
                except Exception:
                    pass
                result[key] = self._clone_line(base_val)
            except Exception:
                continue
        return result

    def _value_from_owner(self, field: str) -> BorderLine | BorderLine2:
        owner = getattr(self, "_owner", None)
        if owner is None:
            raise AttributeError("owner not set")

        getter = getattr(owner, "_border_getter", None)
        if callable(getter):
            try:
                data = getter()
                if field in data:
                    return self._clone_line(data[field])
            except Exception:
                pass

        attr_map = {
            "top": "TopBorder",
            "bottom": "BottomBorder",
            "left": "LeftBorder",
            "right": "RightBorder",
        }
        attr = attr_map.get(field)
        if attr is None:
            raise AttributeError(field)
        base_val = getattr(owner, attr)
        alt_attr = f"{attr}2"
        try:
            alt_val = getattr(owner, alt_attr)
            if hasattr(alt_val, "LineStyle"):
                base_val = alt_val
        except Exception:
            pass
        return self._clone_line(base_val)

    def _write_to_owner(self, **updates: BorderLine | BorderLine2) -> None:
        owner = getattr(self, "_owner", None)
        if owner is None:
            raise AttributeError("owner not set")

        setter = getattr(owner, "_border_setter", None)
        if callable(setter):
            setter(**updates)
            return

        attr_map = {
            "top": "TopBorder",
            "bottom": "BottomBorder",
            "left": "LeftBorder",
            "right": "RightBorder",
        }
        for key, line in updates.items():
            attr = attr_map.get(key)
            if attr is None:
                continue
            setattr(owner, attr, self._clone_line(line))

    def _style(self, side: str) -> BorderStyle:
        existing = self.__dict__.get(f"_{side}_style")
        if existing is None:
            existing = BorderStyle(self, side)
            object.__setattr__(self, f"_{side}_style", existing)
        return existing

    @property
    def top(self) -> BorderStyle:
        return self._style("top")

    @top.setter
    def top(self, value: Any) -> None:
        self.apply(top=value)

    @property
    def bottom(self) -> BorderStyle:
        return self._style("bottom")

    @bottom.setter
    def bottom(self, value: Any) -> None:
        self.apply(bottom=value)

    @property
    def left(self) -> BorderStyle:
        return self._style("left")

    @left.setter
    def left(self, value: Any) -> None:
        self.apply(left=value)

    @property
    def right(self) -> BorderStyle:
        return self._style("right")

    @right.setter
    def right(self, value: Any) -> None:
        self.apply(right=value)

    @property
    def all(self) -> BorderStyle:
        return self._style("all")

    @all.setter
    def all(self, value: Any) -> None:
        self.apply(all=value)

    @property
    def around(self) -> BorderStyle:
        return self._style("around")

    @around.setter
    def around(self, value: Any) -> None:
        self.apply(around=value)

    @property
    def inner(self) -> BorderStyle:
        return self._style("inner")

    @inner.setter
    def inner(self, value: Any) -> None:
        self.apply(inner=value)

    @property
    def top_line(self) -> BorderLine | BorderLine2:
        return self._value_from_owner("top") if self._owner is not None else self._current().get("top", BorderLine2())

    @property
    def bottom_line(self) -> BorderLine | BorderLine2:
        return self._value_from_owner("bottom") if self._owner is not None else self._current().get("bottom", BorderLine2())

    @property
    def left_line(self) -> BorderLine | BorderLine2:
        return self._value_from_owner("left") if self._owner is not None else self._current().get("left", BorderLine2())

    @property
    def right_line(self) -> BorderLine | BorderLine2:
        return self._value_from_owner("right") if self._owner is not None else self._current().get("right", BorderLine2())

    def items(self):  # pragma: no cover - helper
        return self._current().items()

    def __setattr__(self, name: str, value: Any) -> None:  # noqa: D401
        if name.startswith("_"):
            object.__setattr__(self, name, value)
            return
        self.apply(**{name: value})

