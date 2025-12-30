from __future__ import annotations

from typing import Any, cast

from ..core import UnoObject
from ..typing import XPropertySet


class FillProperties(UnoObject):
    """Attribute-style wrapper over drawing FillProperties."""

    def _props(self) -> XPropertySet:
        return cast(XPropertySet, self.raw)

    def get_property(self, name: str) -> Any:
        return self._props().getPropertyValue(name)

    def set_property(self, name: str, value: Any) -> None:
        self._props().setPropertyValue(name, value)

    # Common FillProperties for convenience
    @property
    def FillColor(self) -> int:
        return int(self.get_property("FillColor"))

    @FillColor.setter
    def FillColor(self, value: int) -> None:
        self.set_property("FillColor", int(value))

    @property
    def FillStyle(self) -> Any:
        return self.get_property("FillStyle")

    @FillStyle.setter
    def FillStyle(self, value: Any) -> None:
        self.set_property("FillStyle", value)

    @property
    def FillTransparence(self) -> int:
        return int(self.get_property("FillTransparence"))

    @FillTransparence.setter
    def FillTransparence(self, value: int) -> None:
        self.set_property("FillTransparence", int(value))

    @property
    def FillGradientName(self) -> str:
        return cast(str, self.get_property("FillGradientName"))

    @FillGradientName.setter
    def FillGradientName(self, value: str) -> None:
        self.set_property("FillGradientName", value)

    @property
    def FillHatchName(self) -> str:
        return cast(str, self.get_property("FillHatchName"))

    @FillHatchName.setter
    def FillHatchName(self, value: str) -> None:
        self.set_property("FillHatchName", value)

    @property
    def FillBitmapName(self) -> str:
        return cast(str, self.get_property("FillBitmapName"))

    @FillBitmapName.setter
    def FillBitmapName(self, value: str) -> None:
        self.set_property("FillBitmapName", value)

    @property
    def FillBitmapMode(self) -> int:
        return int(self.get_property("FillBitmapMode"))

    @FillBitmapMode.setter
    def FillBitmapMode(self, value: int) -> None:
        self.set_property("FillBitmapMode", int(value))

    @property
    def FillBitmapOffsetX(self) -> int:
        return int(self.get_property("FillBitmapOffsetX"))

    @FillBitmapOffsetX.setter
    def FillBitmapOffsetX(self, value: int) -> None:
        self.set_property("FillBitmapOffsetX", int(value))

    @property
    def FillBitmapOffsetY(self) -> int:
        return int(self.get_property("FillBitmapOffsetY"))

    @FillBitmapOffsetY.setter
    def FillBitmapOffsetY(self, value: int) -> None:
        self.set_property("FillBitmapOffsetY", int(value))

    @property
    def FillBitmapPositionX(self) -> int:
        return int(self.get_property("FillBitmapPositionX"))

    @FillBitmapPositionX.setter
    def FillBitmapPositionX(self, value: int) -> None:
        self.set_property("FillBitmapPositionX", int(value))

    @property
    def FillBitmapPositionY(self) -> int:
        return int(self.get_property("FillBitmapPositionY"))

    @FillBitmapPositionY.setter
    def FillBitmapPositionY(self, value: int) -> None:
        self.set_property("FillBitmapPositionY", int(value))

    @property
    def FillBitmapSizeX(self) -> int:
        return int(self.get_property("FillBitmapSizeX"))

    @FillBitmapSizeX.setter
    def FillBitmapSizeX(self, value: int) -> None:
        self.set_property("FillBitmapSizeX", int(value))

    @property
    def FillBitmapSizeY(self) -> int:
        return int(self.get_property("FillBitmapSizeY"))

    @FillBitmapSizeY.setter
    def FillBitmapSizeY(self, value: int) -> None:
        self.set_property("FillBitmapSizeY", int(value))

    @property
    def FillBackground(self) -> bool:
        return bool(self.get_property("FillBackground"))

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
