from __future__ import annotations

from typing import Any, cast

from ..core import UnoObject
from ..typing import ShadowLocation, XPropertySet


class ShadowProperties(UnoObject):
    """Attribute-style wrapper over drawing ShadowProperties."""

    def _props(self) -> XPropertySet:
        return cast(XPropertySet, self.raw)

    def get_property(self, name: str) -> Any:
        return self._props().getPropertyValue(name)

    def set_property(self, name: str, value: Any) -> None:
        self._props().setPropertyValue(name, value)

    # Common ShadowProperties for convenience
    @property
    def Shadow(self) -> bool:
        return bool(self.get_property("Shadow"))

    @Shadow.setter
    def Shadow(self, value: bool) -> None:
        self.set_property("Shadow", bool(value))

    @property
    def ShadowColor(self) -> int:
        return int(self.get_property("ShadowColor"))

    @ShadowColor.setter
    def ShadowColor(self, value: int) -> None:
        self.set_property("ShadowColor", int(value))

    @property
    def ShadowTransparence(self) -> int:
        return int(self.get_property("ShadowTransparence"))

    @ShadowTransparence.setter
    def ShadowTransparence(self, value: int) -> None:
        self.set_property("ShadowTransparence", int(value))

    @property
    def ShadowXDistance(self) -> int:
        return int(self.get_property("ShadowXDistance"))

    @ShadowXDistance.setter
    def ShadowXDistance(self, value: int) -> None:
        self.set_property("ShadowXDistance", int(value))

    @property
    def ShadowYDistance(self) -> int:
        return int(self.get_property("ShadowYDistance"))

    @ShadowYDistance.setter
    def ShadowYDistance(self, value: int) -> None:
        self.set_property("ShadowYDistance", int(value))

    @property
    def ShadowLocation(self) -> ShadowLocation:
        return ShadowLocation(int(self.get_property("ShadowLocation")))

    @ShadowLocation.setter
    def ShadowLocation(self, value: ShadowLocation | int) -> None:
        self.set_property("ShadowLocation", int(value))

    def getPropertyValue(self, name: str) -> Any:  # noqa: N802 - UNO naming
        return self.get_property(name)

    def setPropertyValue(self, name: str, value: Any) -> None:  # noqa: N802 - UNO naming
        self.set_property(name, value)

    def __getattr__(self, name: str) -> Any:
        try:
            return self.get_property(name)
        except Exception as exc:  # pragma: no cover - UNO failures bubble up
            raise AttributeError(f"Unknown shadow property: {name}") from exc

    def __setattr__(self, name: str, value: Any) -> None:
        if name.startswith("_"):
            return object.__setattr__(self, name, value)
        try:
            self.set_property(name, value)
        except Exception as exc:  # pragma: no cover - UNO failures bubble up
            raise AttributeError(f"Cannot set shadow property: {name}") from exc
