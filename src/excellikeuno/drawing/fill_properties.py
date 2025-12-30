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
