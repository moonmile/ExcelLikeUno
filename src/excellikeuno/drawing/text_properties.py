from __future__ import annotations

from typing import Any, cast

from ..core import UnoObject
from ..typing import FontSlant, FontStrikeout, FontUnderline, XPropertySet


class TextProperties(UnoObject):
    """Attribute-style wrapper over drawing TextProperties."""

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

    # Common TextProperties for convenience
    @property
    def CharColor(self) -> int:
        try:
            return int(self.get_property("CharColor"))
        except BaseException:
            return 0

    @CharColor.setter
    def CharColor(self, value: int) -> None:
        self.set_property("CharColor", int(value))

    @property
    def CharHeight(self) -> float:
        try:
            return float(self.get_property("CharHeight"))
        except BaseException:
            return 0.0

    @CharHeight.setter
    def CharHeight(self, value: float) -> None:
        self.set_property("CharHeight", float(value))

    @property
    def CharFontName(self) -> str:
        try:
            return cast(str, self.get_property("CharFontName"))
        except BaseException:
            return ""

    @CharFontName.setter
    def CharFontName(self, value: str) -> None:
        self.set_property("CharFontName", value)

    @property
    def CharFontStyleName(self) -> str:
        try:
            return cast(str, self.get_property("CharFontStyleName"))
        except BaseException:
            return ""

    @CharFontStyleName.setter
    def CharFontStyleName(self, value: str) -> None:
        self.set_property("CharFontStyleName", value)

    @property
    def CharFontPitch(self) -> int:
        try:
            return int(self.get_property("CharFontPitch"))
        except BaseException:
            return 0

    @CharFontPitch.setter
    def CharFontPitch(self, value: int) -> None:
        self.set_property("CharFontPitch", int(value))

    @property
    def CharWeight(self) -> float:
        try:
            return float(self.get_property("CharWeight"))
        except BaseException:
            return 0.0

    @CharWeight.setter
    def CharWeight(self, value: float) -> None:
        self.set_property("CharWeight", float(value))

    @property
    def CharPosture(self) -> FontSlant:
        try:
            return FontSlant(int(self.get_property("CharPosture")))
        except BaseException:
            return FontSlant(0)

    @CharPosture.setter
    def CharPosture(self, value: FontSlant | int) -> None:
        self.set_property("CharPosture", int(value))

    @property
    def CharUnderline(self) -> FontUnderline:
        try:
            return FontUnderline(int(self.get_property("CharUnderline")))
        except BaseException:
            return FontUnderline(0)

    @CharUnderline.setter
    def CharUnderline(self, value: FontUnderline | int) -> None:
        self.set_property("CharUnderline", int(value))

    @property
    def CharStrikeout(self) -> FontStrikeout:
        try:
            return FontStrikeout(int(self.get_property("CharStrikeout")))
        except BaseException:
            return FontStrikeout(0)

    @CharStrikeout.setter
    def CharStrikeout(self, value: FontStrikeout | int) -> None:
        self.set_property("CharStrikeout", int(value))

    @property
    def CharLocale(self) -> Any:
        try:
            return self.get_property("CharLocale")
        except BaseException:
            return None

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
