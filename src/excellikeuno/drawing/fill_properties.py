from __future__ import annotations

from typing import Any, cast

from excellikeuno.typing.calc import Color

from ..core import UnoObject
from ..typing import XPropertySet


class FillProperties(UnoObject):
    """Attribute-style wrapper over drawing FillProperties."""

    def _props(self) -> XPropertySet:
        return cast(XPropertySet, self.raw)

    def get_property(self, name: str) -> Any:
        try:
            return self._props().getPropertyValue(name)
        except BaseException:
            pass
        try:
            return getattr(self.raw, name)
        except BaseException:
            return None

    def set_property(self, name: str, value: Any) -> None:
        try:
            self._props().setPropertyValue(name, value)
            return
        except BaseException:
            pass
        try:
            setattr(self.raw, name, value)
        except BaseException:
            # Best-effort; swallow when the fill interface is missing
            pass

    # Common FillProperties for convenience
    @property
    def FillColor(self) -> Color:
        try:
            return Color(self.get_property("FillColor"))
        except BaseException:
            return Color(0)

    @FillColor.setter
    def FillColor(self, value: Color) -> None:
        self.set_property("FillColor", Color(value))

    @property
    def FillStyle(self) -> Any:
        try:
            return self.get_property("FillStyle")
        except BaseException:
            return None

    @FillStyle.setter
    def FillStyle(self, value: Any) -> None:
        self.set_property("FillStyle", value)

    @property
    def FillTransparence(self) -> int:
        try:
            return int(self.get_property("FillTransparence"))
        except BaseException:
            return 0

    @FillTransparence.setter
    def FillTransparence(self, value: int) -> None:
        self.set_property("FillTransparence", int(value))

    @property
    def FillGradientName(self) -> str:
        try:
            return cast(str, self.get_property("FillGradientName"))
        except BaseException:
            return ""

    @FillGradientName.setter
    def FillGradientName(self, value: str) -> None:
        self.set_property("FillGradientName", value)

    @property
    def FillHatchName(self) -> str:
        try:
            return cast(str, self.get_property("FillHatchName"))
        except BaseException:
            return ""

    @FillHatchName.setter
    def FillHatchName(self, value: str) -> None:
        self.set_property("FillHatchName", value)

    @property
    def FillBitmapName(self) -> str:
        try:
            return cast(str, self.get_property("FillBitmapName"))
        except BaseException:
            return ""

    @FillBitmapName.setter
    def FillBitmapName(self, value: str) -> None:
        self.set_property("FillBitmapName", value)

    @property
    def FillBitmapMode(self) -> int:
        try:
            return int(self.get_property("FillBitmapMode"))
        except BaseException:
            return 0

    @FillBitmapMode.setter
    def FillBitmapMode(self, value: int) -> None:
        self.set_property("FillBitmapMode", int(value))

    @property
    def FillBitmapOffsetX(self) -> int:
        try:
            return int(self.get_property("FillBitmapOffsetX"))
        except BaseException:
            return 0

    @FillBitmapOffsetX.setter
    def FillBitmapOffsetX(self, value: int) -> None:
        self.set_property("FillBitmapOffsetX", int(value))

    @property
    def FillBitmapOffsetY(self) -> int:
        try:
            return int(self.get_property("FillBitmapOffsetY"))
        except BaseException:
            return 0

    @FillBitmapOffsetY.setter
    def FillBitmapOffsetY(self, value: int) -> None:
        self.set_property("FillBitmapOffsetY", int(value))

    @property
    def FillBitmapPositionX(self) -> int:
        try:
            return int(self.get_property("FillBitmapPositionX"))
        except BaseException:
            return 0

    @FillBitmapPositionX.setter
    def FillBitmapPositionX(self, value: int) -> None:
        self.set_property("FillBitmapPositionX", int(value))

    @property
    def FillBitmapPositionY(self) -> int:
        try:
            return int(self.get_property("FillBitmapPositionY"))
        except BaseException:
            return 0

    @FillBitmapPositionY.setter
    def FillBitmapPositionY(self, value: int) -> None:
        self.set_property("FillBitmapPositionY", int(value))

    @property
    def FillBitmapSizeX(self) -> int:
        try:
            return int(self.get_property("FillBitmapSizeX"))
        except BaseException:
            return 0

    @FillBitmapSizeX.setter
    def FillBitmapSizeX(self, value: int) -> None:
        self.set_property("FillBitmapSizeX", int(value))

    @property
    def FillBitmapSizeY(self) -> int:
        try:
            return int(self.get_property("FillBitmapSizeY"))
        except BaseException:
            return 0

    @FillBitmapSizeY.setter
    def FillBitmapSizeY(self, value: int) -> None:
        self.set_property("FillBitmapSizeY", int(value))

    @property
    def FillBackground(self) -> bool:
        try:
            return bool(self.get_property("FillBackground"))
        except BaseException:
            return False

    @FillBackground.setter
    def FillBackground(self, value: bool) -> None:
        self.set_property("FillBackground", bool(value))

    # UNO-style aliases
    def getPropertyValue(self, name: str) -> Any:  # noqa: N802 - UNO naming
        return self.get_property(name)

    def setPropertyValue(self, name: str, value: Any) -> None:  # noqa: N802 - UNO naming
        self.set_property(name, value)

    def __getattr__(self, name: str) -> Any:
        try:
            return self.get_property(name)
        except Exception as exc:  # pragma: no cover - UNO failures bubble up
            raise AttributeError(f"Unknown fill property: {name}") from exc

    def __setattr__(self, name: str, value: Any) -> None:
        if name.startswith("_"):
            return object.__setattr__(self, name, value)
        try:
            self.set_property(name, value)
        except Exception as exc:  # pragma: no cover - UNO failures bubble up
            raise AttributeError(f"Cannot set fill property: {name}") from exc
