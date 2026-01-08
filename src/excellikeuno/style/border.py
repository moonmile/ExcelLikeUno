from __future__ import annotations

from typing import Any, Callable, Dict

from excellikeuno.typing import BorderLine, BorderLine2


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
                LineStyle=_as_int(getattr(value, "LineStyle", 1)),
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

    def _normalize_updates(self, updates: Dict[str, Any]) -> Dict[str, BorderLine]:
        normalized = {k: v for k, v in updates.items() if v is not None}
        if "all" in normalized:
            all_line = self._clone_line(normalized.pop("all"))
            for key in ("top", "bottom", "left", "right"):
                normalized.setdefault(key, self._clone_line(all_line))
        for key, value in list(normalized.items()):
            normalized[key] = self._clone_line(value)
        return normalized

    def _current(self) -> Dict[str, BorderLine]:
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
                return self
            except Exception:
                self._buffer.update(updates)
                return self
        if self._owner is not None:
            try:
                self._write_to_owner(**updates)
                return self
            except Exception:
                self._buffer.update(updates)
                return self
        self._buffer.update(updates)
        return self

    def _read_from_owner(self) -> Dict[str, BorderLine]:
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
        result: Dict[str, BorderLine] = {}
        for key, attr in attr_map.items():
            try:
                result[key] = self._clone_line(getattr(owner, attr))
            except Exception:
                continue
        return result

    def _value_from_owner(self, field: str) -> BorderLine:
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
        return self._clone_line(getattr(owner, attr))

    def _write_to_owner(self, **updates: BorderLine) -> None:
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

    @property
    def top(self) -> BorderLine2:
        if "top" in self._buffer:
            return self._buffer["top"]
        try:
            return self._value_from_owner("top")
        except Exception:
            return self._current().get("top", BorderLine2())

    @top.setter
    def top(self, value: Any) -> None:
        self.apply(top=value)

    @property
    def bottom(self) ->     BorderLine2:
        if "bottom" in self._buffer:
            return self._buffer["bottom"]
        try:
            return self._value_from_owner("bottom")
        except Exception:
            return self._current().get("bottom", BorderLine2())

    @bottom.setter
    def bottom(self, value: Any) -> None:
        self.apply(bottom=value)

    @property
    def left(self) ->   BorderLine2:
        if "left" in self._buffer:
            return self._buffer["left"]
        try:
            return self._value_from_owner("left")
        except Exception:
            return self._current().get("left", BorderLine2())

    @left.setter
    def left(self, value: Any) -> None:
        self.apply(left=value)

    @property
    def right(self) -> BorderLine2:
        if "right" in self._buffer:
            return self._buffer["right"]
        try:
            return self._value_from_owner("right")
        except Exception:
            return self._current().get("right", BorderLine2())

    @right.setter
    def right(self, value: Any) -> None:
        self.apply(right=value)

    @property
    def all(self) -> BorderLine2:
        current = self._current()
        if "top" in current:
            return current["top"]
        if "bottom" in current:
            return current["bottom"]
        if "left" in current:
            return current["left"]
        if "right" in current:
            return current["right"]
        return BorderLine2()

    @all.setter
    def all(self, value: Any) -> None:
        self.apply(all=value)

    def items(self):  # pragma: no cover - helper
        return self._current().items()

    def __setattr__(self, name: str, value: Any) -> None:  # noqa: D401
        if name.startswith("_"):
            object.__setattr__(self, name, value)
            return
        self.apply(**{name: value})

