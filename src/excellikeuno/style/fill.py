from __future__ import annotations

from typing import Any, Callable, Dict

from excellikeuno.typing.calc import BitmapMode, Color, FillStyle
from excellikeuno.typing.structs import Gradient, Hatch


class Fill:
    """Fill-property proxy following the Line/Borders apply/buffer pattern."""

    @staticmethod
    def _coerce_style(value: Any) -> FillStyle:
        """Normalize UNO FillStyle values to our FillStyle enum."""
        if isinstance(value, FillStyle):
            return value

        candidates: list[Any] = []
        if value is not None:
            candidates.append(value)
            candidates.append(getattr(value, "value", None))
            candidates.append(getattr(value, "Name", None))

        for cand in candidates:
            if cand is None:
                continue
            try:
                return FillStyle(int(cand))
            except Exception:
                pass
            try:
                return FillStyle[str(cand)]
            except Exception:
                pass

        return FillStyle.NONE

    def __init__(
        self,
        getter: Callable[[], Dict[str, Any]] | None = None,
        setter: Callable[..., None] | None = None,
        owner: Any | None = None,
        **kwargs: Any,
    ) -> None:
        if owner is not None and getter is None:
            getter = getattr(owner, "_fill_getter", None)
        if owner is not None and setter is None:
            setter = getattr(owner, "_fill_setter", None)

        object.__setattr__(self, "_owner", owner)
        object.__setattr__(self, "_getter", getter)
        object.__setattr__(self, "_setter", setter)
        object.__setattr__(self, "_buffer", {})
        if kwargs:
            self.apply(**kwargs)

    # --- internals ---------------------------------------------------------
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

    def apply(self, **kwargs: Any) -> "Fill":
        updates = {k: v for k, v in kwargs.items() if v is not None}
        if not updates:
            return self

        # auto style inference
        style_hint = updates.get("style")
        if style_hint is None:
            if any(k in updates for k in ("color", "backcolor")):
                style_hint = FillStyle.SOLID
            elif any(k in updates for k in ("gradient", "gradient_name")):
                style_hint = FillStyle.GRADIENT
            elif any(k in updates for k in ("hatch", "hatch_name")):
                style_hint = FillStyle.HATCH
            elif any(k in updates for k in ("bitmap", "bitmap_name", "bitmap_mode", "bitmap_offset", "bitmap_position", "bitmap_size")):
                style_hint = FillStyle.BITMAP
        if style_hint is not None:
            updates.setdefault("style", style_hint)

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

    def _read_from_owner(self) -> Dict[str, Any]:
        owner = getattr(self, "_owner", None)
        if owner is None:
            return {}

        og = getattr(owner, "_fill_getter", None)
        if callable(og):
            try:
                return dict(og())
            except Exception:
                pass

        fp = getattr(owner, "fill_properties", None)

        def _get(name: str) -> Any:
            if fp is not None:
                try:
                    return fp.get_property(name)
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

        style_val = self._coerce_style(_get("FillStyle"))

        return {
            "style": style_val,
            "color": _get("FillColor"),
            "transparence": _as_int(_get("FillTransparence")),
            "gradient_name": _get("FillGradientName"),
            "gradient": _get("FillGradient"),
            "hatch_name": _get("FillHatchName"),
            "hatch": _get("FillHatch"),
            "bitmap_name": _get("FillBitmapName"),
            "bitmap_mode": _as_int(_get("FillBitmapMode")),
            "bitmap_offset": (
                _as_int(_get("FillBitmapOffsetX")),
                _as_int(_get("FillBitmapOffsetY")),
            ),
            "bitmap_position": (
                _as_int(_get("FillBitmapPositionX")),
                _as_int(_get("FillBitmapPositionY")),
            ),
            "bitmap_size": (
                _as_int(_get("FillBitmapSizeX")),
                _as_int(_get("FillBitmapSizeY")),
            ),
            "background": bool(_get("FillBackground")),
        }

    def _write_to_owner(self, **updates: Any) -> None:
        owner = getattr(self, "_owner", None)
        if owner is None:
            raise AttributeError("owner not set")

        os = getattr(owner, "_fill_setter", None)
        if callable(os):
            os(**updates)
            return

        fp = getattr(owner, "fill_properties", None)

        def _set(name: str, value: Any) -> bool:
            if fp is not None:
                try:
                    fp.set_property(name, value)
                    return True
                except Exception:
                    pass
            try:
                setattr(owner, name, value)
                return True
            except Exception:
                return False

        if "style" in updates:
            try:
                target_style = self._coerce_style(updates["style"])
                _set("FillStyle", int(target_style))
            except Exception:
                _set("FillStyle", updates["style"])
        if "color" in updates:
            _set("FillColor", Color(updates["color"]))
        if "transparence" in updates:
            _set("FillTransparence", int(updates["transparence"]))
        if "gradient_name" in updates:
            _set("FillGradientName", updates["gradient_name"])
            _set("FillStyle", int(FillStyle.GRADIENT))
        if "gradient" in updates:
            grad = updates["gradient"]
            if hasattr(grad, "to_raw"):
                grad = grad.to_raw()
            _set("FillGradient", grad)
            _set("FillStyle", int(FillStyle.GRADIENT))
        if "hatch_name" in updates:
            _set("FillHatchName", updates["hatch_name"])
            _set("FillStyle", int(FillStyle.HATCH))
        if "hatch" in updates:
            hval = updates["hatch"]
            if hasattr(hval, "to_raw"):
                hval = hval.to_raw()
            _set("FillHatch", hval)
            _set("FillStyle", int(FillStyle.HATCH))
        if "bitmap_name" in updates:
            _set("FillBitmapName", updates["bitmap_name"])
            _set("FillStyle", int(FillStyle.BITMAP))
        if "bitmap_mode" in updates:
            try:
                _set("FillBitmapMode", int(BitmapMode(updates["bitmap_mode"])))
            except Exception:
                _set("FillBitmapMode", int(updates["bitmap_mode"]))
        if "bitmap_offset" in updates:
            try:
                ox, oy = updates["bitmap_offset"]
                _set("FillBitmapOffsetX", int(ox))
                _set("FillBitmapOffsetY", int(oy))
            except Exception:
                pass
        if "bitmap_position" in updates:
            try:
                px, py = updates["bitmap_position"]
                _set("FillBitmapPositionX", int(px))
                _set("FillBitmapPositionY", int(py))
            except Exception:
                pass
        if "bitmap_size" in updates:
            try:
                sx, sy = updates["bitmap_size"]
                _set("FillBitmapSizeX", int(sx))
                _set("FillBitmapSizeY", int(sy))
            except Exception:
                pass
        if "background" in updates:
            _set("FillBackground", bool(updates["background"]))

    # --- properties -------------------------------------------------------
    @property
    def style(self) -> FillStyle:
        cur = self._current()
        if "style" in cur:
            return self._coerce_style(cur["style"])
        return FillStyle.NONE

    @style.setter
    def style(self, value: FillStyle) -> None:
        self.apply(style=value)

    @property
    def color(self) -> Color:
        cur = self._current()
        try:
            return Color(cur.get("color", 0))
        except Exception:
            return Color(0)

    @color.setter
    def color(self, value: Color) -> None:
        self.apply(color=value)

    @property
    def transparence(self) -> int:
        cur = self._current()
        return int(cur.get("transparence", 0))

    @transparence.setter
    def transparence(self, value: int) -> None:
        self.apply(transparence=value)

    @property
    def gradient_name(self) -> str:
        cur = self._current()
        return str(cur.get("gradient_name", ""))

    @gradient_name.setter
    def gradient_name(self, value: str) -> None:
        self.apply(gradient_name=value)

    @property
    def gradient(self) -> Gradient | None:
        cur = self._current()
        val = cur.get("gradient")
        if isinstance(val, Gradient):
            return val
        return None

    @gradient.setter
    def gradient(self, value: Gradient | Any) -> None:
        self.apply(gradient=value)

    @property
    def hatch_name(self) -> str:
        cur = self._current()
        return str(cur.get("hatch_name", ""))

    @hatch_name.setter
    def hatch_name(self, value: str) -> None:
        self.apply(hatch_name=value)

    @property
    def hatch(self) -> Hatch | None:
        cur = self._current()
        val = cur.get("hatch")
        if isinstance(val, Hatch):
            return val
        return None

    @hatch.setter
    def hatch(self, value: Hatch | Any) -> None:
        self.apply(hatch=value)

    @property
    def bitmap_name(self) -> str:
        cur = self._current()
        return str(cur.get("bitmap_name", ""))

    @bitmap_name.setter
    def bitmap_name(self, value: str) -> None:
        self.apply(bitmap_name=value)

    @property
    def bitmap_mode(self) -> BitmapMode:
        cur = self._current()
        try:
            return BitmapMode(cur.get("bitmap_mode", 0))
        except Exception:
            return BitmapMode(0)

    @bitmap_mode.setter
    def bitmap_mode(self, value: BitmapMode) -> None:
        self.apply(bitmap_mode=value)

    @property
    def bitmap_offset(self) -> tuple[int, int]:
        cur = self._current()
        return tuple(cur.get("bitmap_offset", (0, 0)))  # type: ignore[return-value]

    @bitmap_offset.setter
    def bitmap_offset(self, value: tuple[int, int]) -> None:
        self.apply(bitmap_offset=value)

    @property
    def bitmap_position(self) -> tuple[int, int]:
        cur = self._current()
        return tuple(cur.get("bitmap_position", (0, 0)))  # type: ignore[return-value]

    @bitmap_position.setter
    def bitmap_position(self, value: tuple[int, int]) -> None:
        self.apply(bitmap_position=value)

    @property
    def bitmap_size(self) -> tuple[int, int]:
        cur = self._current()
        return tuple(cur.get("bitmap_size", (0, 0)))  # type: ignore[return-value]

    @bitmap_size.setter
    def bitmap_size(self, value: tuple[int, int]) -> None:
        self.apply(bitmap_size=value)

    @property
    def background(self) -> bool:
        cur = self._current()
        return bool(cur.get("background", False))

    @background.setter
    def background(self, value: bool) -> None:
        self.apply(background=value)

    # convenience alias
    def __getattr__(self, name: str) -> Any:  # pragma: no cover - passthrough
        cur = self._current()
        if name in cur:
            return cur[name]
        raise AttributeError(name)

    def __setattr__(self, name: str, value: Any) -> None:  # pragma: no cover - passthrough
        if name.startswith("_"):
            object.__setattr__(self, name, value)
            return
        self.apply(**{name: value})
