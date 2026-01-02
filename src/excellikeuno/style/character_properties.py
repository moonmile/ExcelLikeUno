from __future__ import annotations

from typing import Any, cast

from ..core import UnoObject
from ..typing import Color, FontSlant, FontStrikeout, FontUnderline, XPropertySet


class CharacterProperties(UnoObject):
    """Wraps `com.sun.star.style.CharacterProperties` for attribute-style access."""

    def _props(self) -> XPropertySet:
        # Wrapped object should already be an XPropertySet; avoid extra queries.
        return cast(XPropertySet, self.raw)

    def get_property(self, name: str) -> Any:
        return self._props().getPropertyValue(name)

    def set_property(self, name: str, value: Any) -> None:
        self._props().setPropertyValue(name, value)

    # Common character styling props

    @property
    def CharFontName(self) -> str:
        return cast(str, self.get_property("CharFontName"))

    @CharFontName.setter
    def CharFontName(self, value: str) -> None:
        self.set_property("CharFontName", value)

    @property
    def CharHeight(self) -> float:
        return float(self.get_property("CharHeight"))

    @CharHeight.setter
    def CharHeight(self, value: float) -> None:
        # Set height for all script variants when available (Western/CJK/CTL)
        height = float(value)
        self.set_property("CharHeight", height)
        for alt in ("CharHeightAsian", "CharHeightComplex"):
            try:
                self.set_property(alt, height)
            except Exception:
                # Ignore when the property is not supported by the current object
                pass

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
    def CharColor(self) -> Color:
        return cast(Color, self.get_property("CharColor"))

    @CharColor.setter
    def CharColor(self, value: Color) -> None:
        self.set_property("CharColor", value)

    @property
    def CharUnderlineHasColor(self) -> bool:
        return bool(self.get_property("CharUnderlineHasColor"))

    @CharUnderlineHasColor.setter
    def CharUnderlineHasColor(self, value: bool) -> None:
        self.set_property("CharUnderlineHasColor", bool(value))

    @property
    def CharUnderlineColor(self) -> Color:
        return cast(Color, self.get_property("CharUnderlineColor"))

    @CharUnderlineColor.setter
    def CharUnderlineColor(self, value: Color) -> None:
        self.set_property("CharUnderlineColor", value)

    @property
    def CharShadowed(self) -> bool:
        return bool(self.get_property("CharShadowed"))

    @CharShadowed.setter
    def CharShadowed(self, value: bool) -> None:
        self.set_property("CharShadowed", bool(value))

    @property
    def CharContoured(self) -> bool:
        return bool(self.get_property("CharContoured"))

    @CharContoured.setter
    def CharContoured(self, value: bool) -> None:
        self.set_property("CharContoured", bool(value))

    @property
    def CharCaseMap(self) -> int:
        return int(self.get_property("CharCaseMap"))

    @CharCaseMap.setter
    def CharCaseMap(self, value: int) -> None:
        self.set_property("CharCaseMap", int(value))

    @property
    def CharKerning(self) -> int:
        return int(self.get_property("CharKerning"))

    @CharKerning.setter
    def CharKerning(self, value: int) -> None:
        self.set_property("CharKerning", int(value))

    @property
    def CharAutoKerning(self) -> bool:
        return bool(self.get_property("CharAutoKerning"))

    @CharAutoKerning.setter
    def CharAutoKerning(self, value: bool) -> None:
        self.set_property("CharAutoKerning", bool(value))

    @property
    def CharWordMode(self) -> bool:
        return bool(self.get_property("CharWordMode"))

    @CharWordMode.setter
    def CharWordMode(self, value: bool) -> None:
        self.set_property("CharWordMode", bool(value))

    @property
    def CharRotation(self) -> int:
        return int(self.get_property("CharRotation"))

    @CharRotation.setter
    def CharRotation(self, value: int) -> None:
        self.set_property("CharRotation", int(value))

    @property
    def CharScaleWidth(self) -> int:
        return int(self.get_property("CharScaleWidth"))

    @CharScaleWidth.setter
    def CharScaleWidth(self, value: int) -> None:
        self.set_property("CharScaleWidth", int(value))

    @property
    def CharRelief(self) -> int:
        return int(self.get_property("CharRelief"))

    @CharRelief.setter
    def CharRelief(self, value: int) -> None:
        self.set_property("CharRelief", int(value))

    @property
    def CharEscapement(self) -> int:
        return int(self.get_property("CharEscapement"))

    @CharEscapement.setter
    def CharEscapement(self, value: int) -> None:
        self.set_property("CharEscapement", int(value))

    @property
    def CharEscapementHeight(self) -> int:
        return int(self.get_property("CharEscapementHeight"))

    @CharEscapementHeight.setter
    def CharEscapementHeight(self, value: int) -> None:
        self.set_property("CharEscapementHeight", int(value))

    @property
    def CharLocale(self) -> Any:
        return self.get_property("CharLocale")

    @CharLocale.setter
    def CharLocale(self, value: Any) -> None:
        self.set_property("CharLocale", value)

    @property
    def CharLocaleAsian(self) -> Any:
        return self.get_property("CharLocaleAsian")

    @CharLocaleAsian.setter
    def CharLocaleAsian(self, value: Any) -> None:
        self.set_property("CharLocaleAsian", value)

    @property
    def CharLocaleComplex(self) -> Any:
        return self.get_property("CharLocaleComplex")

    @CharLocaleComplex.setter
    def CharLocaleComplex(self, value: Any) -> None:
        self.set_property("CharLocaleComplex", value)

    @property
    def CharFontFamily(self) -> int:
        return int(self.get_property("CharFontFamily"))

    @CharFontFamily.setter
    def CharFontFamily(self, value: int) -> None:
        self.set_property("CharFontFamily", int(value))

    @property
    def CharFontCharSet(self) -> int:
        return int(self.get_property("CharFontCharSet"))

    @CharFontCharSet.setter
    def CharFontCharSet(self, value: int) -> None:
        self.set_property("CharFontCharSet", int(value))

    @property
    def CharFontPitch(self) -> int:
        return int(self.get_property("CharFontPitch"))

    @CharFontPitch.setter
    def CharFontPitch(self, value: int) -> None:
        self.set_property("CharFontPitch", int(value))


    # Passthrough for any other character property
    def __getattr__(self, name: str) -> Any:
        if name.startswith("_"):
            raise AttributeError(name)
        try:
            return self.get_property(name)
        except Exception as exc:  # pragma: no cover
            raise AttributeError(f"Unknown character property: {name}") from exc

    def __setattr__(self, name: str, value: Any) -> None:
        if name.startswith("_"):
            return object.__setattr__(self, name, value)
        cls_attr = getattr(type(self), name, None)
        if isinstance(cls_attr, property):
            setter = cls_attr.fset
            if setter is None:
                raise AttributeError(f"can't set attribute {name}")
            setter(self, value)
            return
        try:
            self.set_property(name, value)
        except Exception:
            object.__setattr__(self, name, value)
