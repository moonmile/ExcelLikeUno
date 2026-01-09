from __future__ import annotations

from typing import Any, cast

from ..core import UnoObject
from ..typing import XPropertySet


class TextProperties(UnoObject):
    """Attribute-style wrapper over drawing TextProperties (Text* only)."""

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
            # Best-effort; swallow when the text interface is missing
            pass

    # Text* convenience (Char* は CharacterProperties 側に委譲する)
    @property
    def TextAutoGrowHeight(self) -> bool:
        try:
            return bool(self.get_property("TextAutoGrowHeight"))
        except BaseException:
            return False

    @TextAutoGrowHeight.setter
    def TextAutoGrowHeight(self, value: bool) -> None:
        self.set_property("TextAutoGrowHeight", bool(value))

    @property
    def TextAutoGrowWidth(self) -> bool:
        try:
            return bool(self.get_property("TextAutoGrowWidth"))
        except BaseException:
            return False

    @TextAutoGrowWidth.setter
    def TextAutoGrowWidth(self, value: bool) -> None:
        self.set_property("TextAutoGrowWidth", bool(value))

    @property
    def TextLeftDistance(self) -> int:
        try:
            return int(self.get_property("TextLeftDistance"))
        except BaseException:
            return 0

    @TextLeftDistance.setter
    def TextLeftDistance(self, value: int) -> None:
        self.set_property("TextLeftDistance", int(value))

    @property
    def TextRightDistance(self) -> int:
        try:
            return int(self.get_property("TextRightDistance"))
        except BaseException:
            return 0

    @TextRightDistance.setter
    def TextRightDistance(self, value: int) -> None:
        self.set_property("TextRightDistance", int(value))

    @property
    def TextUpperDistance(self) -> int:
        try:
            return int(self.get_property("TextUpperDistance"))
        except BaseException:
            return 0

    @TextUpperDistance.setter
    def TextUpperDistance(self, value: int) -> None:
        self.set_property("TextUpperDistance", int(value))

    @property
    def TextLowerDistance(self) -> int:
        try:
            return int(self.get_property("TextLowerDistance"))
        except BaseException:
            return 0

    @TextLowerDistance.setter
    def TextLowerDistance(self, value: int) -> None:
        self.set_property("TextLowerDistance", int(value))

    def getPropertyValue(self, name: str) -> Any:  # noqa: N802 - UNO naming
        return self.get_property(name)

    def setPropertyValue(self, name: str, value: Any) -> None:  # noqa: N802 - UNO naming
        self.set_property(name, value)

    def __getattr__(self, name: str) -> Any:
        try:
            return self.get_property(name)
        except Exception as exc:  # pragma: no cover - UNO failures bubble up
            raise AttributeError(f"Unknown text property: {name}") from exc

    def __setattr__(self, name: str, value: Any) -> None:
        if name.startswith("_"):
            return object.__setattr__(self, name, value)
        try:
            self.set_property(name, value)
        except Exception as exc:  # pragma: no cover - UNO failures bubble up
            raise AttributeError(f"Cannot set text property: {name}") from exc
