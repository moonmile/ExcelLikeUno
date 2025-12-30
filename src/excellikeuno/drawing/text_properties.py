from __future__ import annotations

from typing import Any, cast

from ..core import UnoObject
from ..typing import FontSlant, FontStrikeout, FontUnderline, XPropertySet


class TextProperties(UnoObject):
    """Attribute-style wrapper over drawing TextProperties."""

    def _props(self) -> XPropertySet:
        return cast(XPropertySet, self.raw)

    def get_property(self, name: str) -> Any:
        return self._props().getPropertyValue(name)

    def set_property(self, name: str, value: Any) -> None:
        self._props().setPropertyValue(name, value)

    # Common TextProperties for convenience
    @property
    def CharColor(self) -> int:
        return int(self.get_property("CharColor"))

    @CharColor.setter
    def CharColor(self, value: int) -> None:
        self.set_property("CharColor", int(value))

    @property
    def CharHeight(self) -> float:
        return float(self.get_property("CharHeight"))

    @CharHeight.setter
    def CharHeight(self, value: float) -> None:
        self.set_property("CharHeight", float(value))

    @property
    def CharFontName(self) -> str:
        return cast(str, self.get_property("CharFontName"))

    @CharFontName.setter
    def CharFontName(self, value: str) -> None:
        self.set_property("CharFontName", value)

    @property
    def CharFontStyleName(self) -> str:
        return cast(str, self.get_property("CharFontStyleName"))

    @CharFontStyleName.setter
    def CharFontStyleName(self, value: str) -> None:
        self.set_property("CharFontStyleName", value)

    @property
    def CharFontPitch(self) -> int:
        return int(self.get_property("CharFontPitch"))

    @CharFontPitch.setter
    def CharFontPitch(self, value: int) -> None:
        self.set_property("CharFontPitch", int(value))

    @property
    def CharWeight(self) -> float:
        return float(self.get_property("CharWeight"))

    @CharWeight.setter
    def CharWeight(self, value: float) -> None:
        self.set_property("CharWeight", float(value))

    @property
    def CharPosture(self) -> FontSlant:
        return FontSlant(int(self.get_property("CharPosture")))

    @CharPosture.setter
    def CharPosture(self, value: FontSlant | int) -> None:
        self.set_property("CharPosture", int(value))

    @property
    def CharUnderline(self) -> FontUnderline:
        return FontUnderline(int(self.get_property("CharUnderline")))

    @CharUnderline.setter
    def CharUnderline(self, value: FontUnderline | int) -> None:
        self.set_property("CharUnderline", int(value))

    @property
    def CharStrikeout(self) -> FontStrikeout:
        return FontStrikeout(int(self.get_property("CharStrikeout")))

    @CharStrikeout.setter
    def CharStrikeout(self, value: FontStrikeout | int) -> None:
        self.set_property("CharStrikeout", int(value))

    @property
    def CharLocale(self) -> Any:
        return self.get_property("CharLocale")

    @CharLocale.setter
    def CharLocale(self, value: Any) -> None:
        self.set_property("CharLocale", value)

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
