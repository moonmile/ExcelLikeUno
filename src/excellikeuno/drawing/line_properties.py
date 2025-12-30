from __future__ import annotations

from typing import Any, cast

from ..core import UnoObject
from ..typing import LineDash, LineStyle, XPropertySet


class LineProperties(UnoObject):
    """Attribute-style wrapper over drawing LineProperties."""

    def _props(self) -> XPropertySet:
        return cast(XPropertySet, self.raw)

    def get_property(self, name: str) -> Any:
        return self._props().getPropertyValue(name)

    def set_property(self, name: str, value: Any) -> None:
        self._props().setPropertyValue(name, value)

    # Common LineProperties for convenience
    @property
    def LineColor(self) -> int:
        return int(self.get_property("LineColor"))

    @LineColor.setter
    def LineColor(self, value: int) -> None:
        self.set_property("LineColor", int(value))

    @property
    def LineStyle(self) -> LineStyle:
        return LineStyle(int(self.get_property("LineStyle")))

    @LineStyle.setter
    def LineStyle(self, value: LineStyle | int) -> None:
        self.set_property("LineStyle", int(value))

    @property
    def LineWidth(self) -> int:
        return int(self.get_property("LineWidth"))

    @LineWidth.setter
    def LineWidth(self, value: int) -> None:
        self.set_property("LineWidth", int(value))

    @property
    def LineTransparence(self) -> int:
        return int(self.get_property("LineTransparence"))

    @LineTransparence.setter
    def LineTransparence(self, value: int) -> None:
        self.set_property("LineTransparence", int(value))

    @property
    def LineDashName(self) -> str:
        return cast(str, self.get_property("LineDashName"))

    @LineDashName.setter
    def LineDashName(self, value: str) -> None:
        self.set_property("LineDashName", value)

    @property
    def LineDash(self) -> LineDash:
        return cast(LineDash, self.get_property("LineDash"))

    @LineDash.setter
    def LineDash(self, value: LineDash) -> None:
        self.set_property("LineDash", value)

    def getPropertyValue(self, name: str) -> Any:  # noqa: N802 - UNO naming
        return self.get_property(name)

    def setPropertyValue(self, name: str, value: Any) -> None:  # noqa: N802 - UNO naming
        self.set_property(name, value)

    def __getattr__(self, name: str) -> Any:
        try:
            return self.get_property(name)
        except Exception as exc:  # pragma: no cover - UNO failures bubble up
            raise AttributeError(f"Unknown line property: {name}") from exc

    def __setattr__(self, name: str, value: Any) -> None:
        if name.startswith("_"):
            return object.__setattr__(self, name, value)
        try:
            self.set_property(name, value)
        except Exception as exc:  # pragma: no cover - UNO failures bubble up
            raise AttributeError(f"Cannot set line property: {name}") from exc
