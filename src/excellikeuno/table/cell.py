from __future__ import annotations

from typing import Any, cast
from numbers import Number

from ..core import UnoObject
from ..typing import (
    Color,
    BorderLine,
    BorderLine2,
    CellHoriJustify,
    CellOrientation,
    CellVertJustify,
    CellProtection,
    InterfaceNames,
    ShadowFormat,
    TableBorder,
    TableBorder2,
    XCell,
    XPropertySet,
)
from ..typing.structs import Point
from ..style.font import Font
from .cell_properties import CellProperties
from ..style.character_properties import CharacterProperties
from .rows import TableRows


class Cell(UnoObject):
    def _get_prop(self, name: str) -> Any:
        props = cast(XPropertySet, self.iface(InterfaceNames.X_PROPERTY_SET))
        return props.getPropertyValue(name)

    def _set_prop(self, name: str, value: Any) -> None:
        props = cast(XPropertySet, self.iface(InterfaceNames.X_PROPERTY_SET))
        props.setPropertyValue(name, value)

    @property
    def properties(self) -> CellProperties:
        existing = self.__dict__.get("_properties")
        if existing is None:
            existing = CellProperties(self.iface(InterfaceNames.X_PROPERTY_SET))
            object.__setattr__(self, "_properties", existing)
        return cast(CellProperties, existing)
    
    @property
    def character_properties(self) -> CharacterProperties:
        existing = self.__dict__.get("_character_properties")
        if existing is None:
            existing = CharacterProperties(self.iface(InterfaceNames.X_PROPERTY_SET))
            object.__setattr__(self, "_character_properties", existing)
        return cast(CharacterProperties, existing)

    # CellProperties (public attributes) shortcuts for IDE completion
    @property
    def CellStyle(self) -> str:
        return self.properties.CellStyle

    @CellStyle.setter
    def CellStyle(self, value: str) -> None:
        self.properties.CellStyle = value

    @property
    def CellBackColor(self) -> Color:
        return self.properties.CellBackColor

    @CellBackColor.setter
    def CellBackColor(self, value: Color) -> None:
        self.properties.CellBackColor = value

    @property
    def IsCellBackgroundTransparent(self) -> bool:
        return self.properties.IsCellBackgroundTransparent

    @IsCellBackgroundTransparent.setter
    def IsCellBackgroundTransparent(self, value: bool) -> None:
        self.properties.IsCellBackgroundTransparent = value

    @property
    def HoriJustify(self) -> CellHoriJustify:
        return self.properties.HoriJustify

    @HoriJustify.setter
    def HoriJustify(self, value: CellHoriJustify | int) -> None:
        self.properties.HoriJustify = value

    @property
    def VertJustify(self) -> CellVertJustify:
        return self.properties.VertJustify

    @VertJustify.setter
    def VertJustify(self, value: CellVertJustify | int) -> None:
        self.properties.VertJustify = value

    @property
    def IsTextWrapped(self) -> bool:
        return self.properties.IsTextWrapped

    @IsTextWrapped.setter
    def IsTextWrapped(self, value: bool) -> None:
        self.properties.IsTextWrapped = value

    @property
    def ParaIndent(self) -> int:
        return self.properties.ParaIndent

    @ParaIndent.setter
    def ParaIndent(self, value: int) -> None:
        self.properties.ParaIndent = value

    @property
    def Orientation(self) -> CellOrientation:
        return self.properties.Orientation

    @Orientation.setter
    def Orientation(self, value: CellOrientation) -> None:
        self.properties.Orientation = value

    @property
    def RotateAngle(self) -> int:
        return self.properties.RotateAngle

    @RotateAngle.setter
    def RotateAngle(self, value: int) -> None:
        self.properties.RotateAngle = value

    @property
    def RotateReference(self) -> int:
        return self.properties.RotateReference

    @RotateReference.setter
    def RotateReference(self, value: int) -> None:
        self.properties.RotateReference = value

    @property
    def AsianVerticalMode(self) -> bool:
        return self.properties.AsianVerticalMode

    @AsianVerticalMode.setter
    def AsianVerticalMode(self, value: bool) -> None:
        self.properties.AsianVerticalMode = value

    @property
    def TableBorder(self) -> TableBorder:
        return self.properties.TableBorder

    @TableBorder.setter
    def TableBorder(self, value: TableBorder) -> None:
        self.properties.TableBorder = value

    @property
    def TopBorder(self) -> BorderLine:
        return self.properties.TopBorder

    @TopBorder.setter
    def TopBorder(self, value: BorderLine) -> None:
        self.properties.TopBorder = value

    @property
    def BottomBorder(self) -> BorderLine:
        return self.properties.BottomBorder

    @BottomBorder.setter
    def BottomBorder(self, value: BorderLine) -> None:
        self.properties.BottomBorder = value

    @property
    def LeftBorder(self) -> BorderLine:
        return self.properties.LeftBorder

    @LeftBorder.setter
    def LeftBorder(self, value: BorderLine) -> None:
        self.properties.LeftBorder = value

    @property
    def RightBorder(self) -> BorderLine:
        return self.properties.RightBorder

    @RightBorder.setter
    def RightBorder(self, value: BorderLine) -> None:
        self.properties.RightBorder = value

    @property
    def NumberFormat(self) -> int:
        return self.properties.NumberFormat

    @NumberFormat.setter
    def NumberFormat(self, value: int) -> None:
        self.properties.NumberFormat = value

    @property
    def ShadowFormat(self) -> ShadowFormat:
        return self.properties.ShadowFormat

    @ShadowFormat.setter
    def ShadowFormat(self, value: ShadowFormat) -> None:
        self.properties.ShadowFormat = value

    @property
    def CellProtection(self) -> CellProtection:
        return self.properties.CellProtection

    @CellProtection.setter
    def CellProtection(self, value: CellProtection) -> None:
        self.properties.CellProtection = value

    @property
    def UserDefinedAttributes(self) -> Any:
        return self.properties.UserDefinedAttributes

    @UserDefinedAttributes.setter
    def UserDefinedAttributes(self, value: Any) -> None:
        self.properties.UserDefinedAttributes = value

    @property
    def DiagonalTLBR(self) -> BorderLine:
        return self.properties.DiagonalTLBR

    @DiagonalTLBR.setter
    def DiagonalTLBR(self, value: BorderLine) -> None:
        self.properties.DiagonalTLBR = value

    @property
    def DiagonalBLTR(self) -> BorderLine:
        return self.properties.DiagonalBLTR

    @DiagonalBLTR.setter
    def DiagonalBLTR(self, value: BorderLine) -> None:
        self.properties.DiagonalBLTR = value

    @property
    def ShrinkToFit(self) -> bool:
        return self.properties.ShrinkToFit

    @ShrinkToFit.setter
    def ShrinkToFit(self, value: bool) -> None:
        self.properties.ShrinkToFit = value

    @property
    def TableBorder2(self) -> TableBorder2:
        return self.properties.TableBorder2

    @TableBorder2.setter
    def TableBorder2(self, value: TableBorder2) -> None:
        self.properties.TableBorder2 = value

    @property
    def TopBorder2(self) -> BorderLine2:
        return self.properties.TopBorder2

    @TopBorder2.setter
    def TopBorder2(self, value: BorderLine2) -> None:
        self.properties.TopBorder2 = value

    @property
    def BottomBorder2(self) -> BorderLine2:
        return self.properties.BottomBorder2

    @BottomBorder2.setter
    def BottomBorder2(self, value: BorderLine2) -> None:
        self.properties.BottomBorder2 = value

    @property
    def LeftBorder2(self) -> BorderLine2:
        return self.properties.LeftBorder2

    @LeftBorder2.setter
    def LeftBorder2(self, value: BorderLine2) -> None:
        self.properties.LeftBorder2 = value

    @property
    def RightBorder2(self) -> BorderLine2:
        return self.properties.RightBorder2

    @RightBorder2.setter
    def RightBorder2(self, value: BorderLine2) -> None:
        self.properties.RightBorder2 = value

    @property
    def DiagonalTLBR2(self) -> RightBorder2:
        return self.properties.DiagonalTLBR2

    @DiagonalTLBR2.setter
    def DiagonalTLBR2(self, value: RightBorder2) -> None:
        self.properties.DiagonalTLBR2 = value

    @property
    def DiagonalBLTR2(self) -> RightBorder2:
        return self.properties.DiagonalBLTR2

    @DiagonalBLTR2.setter
    def DiagonalBLTR2(self, value: RightBorder2) -> None:
        self.properties.DiagonalBLTR2 = value

    @property
    def CellInteropGrabBag(self) -> Any:
        return self.properties.CellInteropGrabBag

    @CellInteropGrabBag.setter
    def CellInteropGrabBag(self, value: Any) -> None:
        self.properties.CellInteropGrabBag = value

    # CharacterProperties (public attributes) shortcuts for IDE completion
    @property
    def CharFontName(self) -> str:
        return self.character_properties.CharFontName

    @CharFontName.setter
    def CharFontName(self, value: str) -> None:
        self.character_properties.CharFontName = value

    @property
    def CharHeight(self) -> float:
        return self.character_properties.CharHeight

    @CharHeight.setter
    def CharHeight(self, value: float) -> None:
        self.character_properties.CharHeight = value

    @property
    def CharWeight(self) -> float:
        return self.character_properties.CharWeight

    @CharWeight.setter
    def CharWeight(self, value: float) -> None:
        self.character_properties.CharWeight = value

    @property
    def CharPosture(self) -> int:
        return int(self.character_properties.CharPosture)

    @CharPosture.setter
    def CharPosture(self, value: int) -> None:
        self.character_properties.CharPosture = value

    @property
    def CharUnderline(self) -> int:
        return int(self.character_properties.CharUnderline)

    @CharUnderline.setter
    def CharUnderline(self, value: int) -> None:
        self.character_properties.CharUnderline = value

    @property
    def CharStrikeout(self) -> int:
        return int(self.character_properties.CharStrikeout)

    @CharStrikeout.setter
    def CharStrikeout(self, value: int) -> None:
        self.character_properties.CharStrikeout = value

    @property
    def CharColor(self) -> Color:
        return self.character_properties.CharColor

    @CharColor.setter
    def CharColor(self, value: Color) -> None:
        self.character_properties.CharColor = value

    @property
    def CharUnderlineHasColor(self) -> bool:
        return self.character_properties.CharUnderlineHasColor

    @CharUnderlineHasColor.setter
    def CharUnderlineHasColor(self, value: bool) -> None:
        self.character_properties.CharUnderlineHasColor = value

    @property
    def CharUnderlineColor(self) -> Color:
        return self.character_properties.CharUnderlineColor

    @CharUnderlineColor.setter
    def CharUnderlineColor(self, value: Color) -> None:
        self.character_properties.CharUnderlineColor = value

    @property
    def CharShadowed(self) -> bool:
        return self.character_properties.CharShadowed

    @CharShadowed.setter
    def CharShadowed(self, value: bool) -> None:
        self.character_properties.CharShadowed = value

    @property
    def CharContoured(self) -> bool:
        return self.character_properties.CharContoured

    @CharContoured.setter
    def CharContoured(self, value: bool) -> None:
        self.character_properties.CharContoured = value

    @property
    def CharCaseMap(self) -> int:
        return self.character_properties.CharCaseMap

    @CharCaseMap.setter
    def CharCaseMap(self, value: int) -> None:
        self.character_properties.CharCaseMap = value

    @property
    def CharKerning(self) -> int:
        return self.character_properties.CharKerning

    @CharKerning.setter
    def CharKerning(self, value: int) -> None:
        self.character_properties.CharKerning = value

    @property
    def CharAutoKerning(self) -> bool:
        return self.character_properties.CharAutoKerning

    @CharAutoKerning.setter
    def CharAutoKerning(self, value: bool) -> None:
        self.character_properties.CharAutoKerning = value

    @property
    def CharWordMode(self) -> bool:
        return self.character_properties.CharWordMode

    @CharWordMode.setter
    def CharWordMode(self, value: bool) -> None:
        self.character_properties.CharWordMode = value

    @property
    def CharRotation(self) -> int:
        # Prefer character property when available; fallback to cell RotateAngle (1/100 deg)
        try:
            return int(self.character_properties.CharRotation)
        except Exception:
            try:
                return int(self.properties.RotateAngle)
            except Exception:
                raise

    @CharRotation.setter
    def CharRotation(self, value: int) -> None:
        # Accept degrees; convert to 1/100 degrees when value looks like plain degrees.
        degrees_hundredths = int(round(value * 100)) if -360 <= value <= 360 else int(value)
        try:
            self.character_properties.CharRotation = degrees_hundredths
            return
        except Exception:
            pass
        # Fallback to cell rotation (RotateAngle) when CharRotation is not supported.
        self.properties.RotateAngle = degrees_hundredths

    @property
    def CharScaleWidth(self) -> int:
        return self.character_properties.CharScaleWidth

    @CharScaleWidth.setter
    def CharScaleWidth(self, value: int) -> None:
        self.character_properties.CharScaleWidth = value

    @property
    def CharRelief(self) -> int:
        return self.character_properties.CharRelief

    @CharRelief.setter
    def CharRelief(self, value: int) -> None:
        self.character_properties.CharRelief = value

    @property
    def CharEscapement(self) -> int:
        return self.character_properties.CharEscapement

    @CharEscapement.setter
    def CharEscapement(self, value: int) -> None:
        self.character_properties.CharEscapement = value

    @property
    def CharEscapementHeight(self) -> int:
        return self.character_properties.CharEscapementHeight

    @CharEscapementHeight.setter
    def CharEscapementHeight(self, value: int) -> None:
        self.character_properties.CharEscapementHeight = value

    @property
    def CharLocale(self) -> Any:
        return self.character_properties.CharLocale

    @CharLocale.setter
    def CharLocale(self, value: Any) -> None:
        self.character_properties.CharLocale = value

    @property
    def CharLocaleAsian(self) -> Any:
        return self.character_properties.CharLocaleAsian

    @CharLocaleAsian.setter
    def CharLocaleAsian(self, value: Any) -> None:
        self.character_properties.CharLocaleAsian = value

    @property
    def CharLocaleComplex(self) -> Any:
        return self.character_properties.CharLocaleComplex

    @CharLocaleComplex.setter
    def CharLocaleComplex(self, value: Any) -> None:
        self.character_properties.CharLocaleComplex = value

    @property
    def CharFontFamily(self) -> int:
        return self.character_properties.CharFontFamily

    @CharFontFamily.setter
    def CharFontFamily(self, value: int) -> None:
        self.character_properties.CharFontFamily = value

    @property
    def CharFontCharSet(self) -> int:
        return self.character_properties.CharFontCharSet

    @CharFontCharSet.setter
    def CharFontCharSet(self, value: int) -> None:
        self.character_properties.CharFontCharSet = value

    @property
    def CharFontPitch(self) -> int:
        return self.character_properties.CharFontPitch

    @CharFontPitch.setter
    def CharFontPitch(self, value: int) -> None:
        self.character_properties.CharFontPitch = value

    def _font_getter(self) -> dict[str, Any]:
        try:
            esc = int(self.CharEscapement)
        except Exception:
            esc = 0
        def _safe_int(val: Any, default: int = 0) -> int:
            try:
                return int(val)
            except Exception:
                return default
        def _safe_float(val: Any, default: float = 0.0) -> float:
            try:
                return float(val)
            except Exception:
                return default
        backcolor: Any = None
        try:
            backcolor = self.CellBackColor
        except Exception:
            backcolor = None
        try:
            prop_back = self.character_properties.CharBackColor
            if prop_back is not None:
                backcolor = prop_back
        except Exception:
            pass
        cache_backcolor = getattr(self, "_font_backcolor_cache", None)
        if cache_backcolor is not None:
            backcolor = cache_backcolor
        font_state = {
            "name": self.CharFontName,
            "size": _safe_float(self.CharHeight),
            "bold": _safe_float(self.CharWeight) >= 150.0,
            "italic": bool(_safe_int(self.CharPosture)),
            "underline": _safe_int(self.CharUnderline),
            "strikeout": _safe_int(self.CharStrikeout),
            "color": self.CharColor,
            "subscript": esc < 0,
            "superscript": esc > 0,
            "backcolor": backcolor,
            "font_style": _safe_int(self.CharPosture),
            "strikthrough": _safe_int(self.CharStrikeout) != 0,
        }
        cache = getattr(self, "_font_cache", None)
        if cache:
            font_state.update(cache)
        return font_state

    def _font_setter(self, **updates: Any) -> None:
        cache = self.__dict__.get("_font_cache")
        if cache is None:
            cache = {}
            object.__setattr__(self, "_font_cache", cache)
        if "name" in updates:
            self.CharFontName = updates["name"]
            cache["name"] = updates["name"]
        if "size" in updates:
            self.CharHeight = float(updates["size"])
            cache["size"] = updates["size"]
        if "bold" in updates:
            self.CharWeight = 150.0 if updates["bold"] else 100.0
            cache["bold"] = updates["bold"]
        if "italic" in updates:
            self.CharPosture = 2 if updates["italic"] else 0
            cache["italic"] = updates["italic"]
        if "font_style" in updates:
            try:
                self.CharPosture = int(updates["font_style"])
            except Exception:
                pass
            cache["font_style"] = updates["font_style"]
        if "underline" in updates:
            self.CharUnderline = int(updates["underline"])
            cache["underline"] = updates["underline"]
        if "strikeout" in updates:
            self.CharStrikeout = int(updates["strikeout"])
            cache["strikeout"] = updates["strikeout"]
        if "color" in updates:
            self.CharColor = updates["color"]
            cache["color"] = updates["color"]
        if "backcolor" in updates:
            value = updates["backcolor"]
            cache["backcolor"] = value
            object.__setattr__(self, "_font_backcolor_cache", value)
            try:
                self.character_properties.CharBackTransparent = False
                self.character_properties.CharBackColor = value
            except Exception:
                pass
            try:
                self.CellBackColor = value
            except Exception:
                pass
        if "subscript" in updates or "superscript" in updates:
            if updates.get("superscript"):
                self.CharEscapement = 58
            elif updates.get("subscript"):
                self.CharEscapement = -25
            else:
                self.CharEscapement = 0
            cache["subscript"] = updates.get("subscript")
            cache["superscript"] = updates.get("superscript")
        if "strikthrough" in updates:
            try:
                self.CharStrikeout = 1 if updates["strikthrough"] else 0
            except Exception:
                pass
            cache["strikthrough"] = updates["strikthrough"]

    @property
    def font(self) -> Font:
        return Font(getter=self._font_getter, setter=self._font_setter)

    @font.setter
    def font(self, value: Font) -> None:
        # Accept a Font proxy or plain Font config
        try:
            current = value._current()  # type: ignore[attr-defined]
        except Exception:
            current = {}
        if not current:
            return
        self._font_setter(**current)

    # 行の高さ
    @property
    def row_height(self) -> int:
        # Resolve the parent sheet via XSheetCellRange and read the row's Height property (1/100 mm)
        sheet_range = self.iface(InterfaceNames.X_SHEET_CELL_RANGE)
        if sheet_range is None:
            raise AttributeError("XSheetCellRange not available for this cell")

        sheet = sheet_range.getSpreadsheet()
        row_idx = self.raw.getCellAddress().Row
        rows = TableRows(sheet.getRows())
        row = rows.getByIndex(row_idx)
        try:
            return int(row.raw.getPropertyValue("Height"))
        except Exception:
            # Some UNO implementations expose Height as a direct attribute
            return int(getattr(row.raw, "Height"))
    
    @row_height.setter
    def row_height(self, height: int) -> None:
        sheet_range = self.iface(InterfaceNames.X_SHEET_CELL_RANGE)
        if sheet_range is None:
            raise AttributeError("XSheetCellRange not available for this cell")

        sheet = sheet_range.getSpreadsheet()
        row_idx = self.raw.getCellAddress().Row
        rows = TableRows(sheet.getRows())
        row = rows.getByIndex(row_idx)

        # Disable optimal height before setting an explicit height when available
        try:
            row.raw.setPropertyValue("OptimalHeight", False)
        except Exception:
            pass

        try:
            row.raw.setPropertyValue("Height", int(height))
        except Exception:
            setattr(row.raw, "Height", int(height))

    @property
    def column_width(self) -> int:
        # Resolve the parent sheet via XSheetCellRange and read the column's Width property (1/100 mm)
        sheet_range = self.iface(InterfaceNames.X_SHEET_CELL_RANGE)
        if sheet_range is None:
            raise AttributeError("XSheetCellRange not available for this cell")

        sheet = sheet_range.getSpreadsheet()
        col_idx = self.raw.getCellAddress().Column
        columns = TableRows(sheet.getColumns())
        column = columns.getByIndex(col_idx)
        try:
            return int(column.raw.getPropertyValue("Width"))
        except Exception:
            # Some UNO implementations expose Width as a direct attribute
            return int(getattr(column.raw, "Width"))
        
    @column_width.setter
    def column_width(self, width: int) -> None:
        sheet_range = self.iface(InterfaceNames.X_SHEET_CELL_RANGE)
        if sheet_range is None:
            raise AttributeError("XSheetCellRange not available for this cell")

        sheet = sheet_range.getSpreadsheet()
        col_idx = self.raw.getCellAddress().Column
        columns = TableRows(sheet.getColumns())
        column = columns.getByIndex(col_idx)

        # Disable optimal width before setting an explicit width when available
        try:
            column.raw.setPropertyValue("OptimalWidth", False)
        except Exception:
            pass

        try:
            column.raw.setPropertyValue("Width", int(width))
        except Exception:
            setattr(column.raw, "Width", int(width))

    @property 
    def position(self) -> Point:
        pos = self.properties.getPropertyValue("Position")
        return Point(pos.X, pos.Y)

    @property
    def value(self) -> float:
        cell = cast(XCell, self.iface(InterfaceNames.X_CELL))
        return cell.getValue()

    @value.setter
    def value(self, value: Any) -> None:
        cell_iface = cast(XCell, self.iface(InterfaceNames.X_CELL))
        if isinstance(value, Number):
            cell_iface.setValue(float(value))
        else:
            # Fallback to string/formula set
            cell_iface.setFormula(str(value))

    @property
    def formula(self) -> str:
        cell = cast(XCell, self.iface(InterfaceNames.X_CELL))
        return cell.getFormula()

    @formula.setter
    def formula(self, formula: str) -> None:
        cell = cast(XCell, self.iface(InterfaceNames.X_CELL))
        cell.setFormula(formula)

    @property
    def text(self) -> str:
        cell = cast(XCell, self.iface(InterfaceNames.X_CELL))
        return cell.getFormula()

    @text.setter
    def text(self, text: str) -> None:
        cell = cast(XCell, self.iface(InterfaceNames.X_CELL))
        cell.setFormula(text)

    @property
    def props(self) -> CellProperties:
        return self.properties

    def __getattr__(self, name: str) -> Any:
        # Avoid recursion on internal attributes
        if name in {"_properties", "properties", "props"}:
            raise AttributeError(name)
        return getattr(self.properties, name)

    def __setattr__(self, name: str, value: Any) -> None:
        cls_attr = getattr(type(self), name, None)
        if name.startswith("_"):
            return object.__setattr__(self, name, value)
        if isinstance(cls_attr, property):
            setter = cls_attr.fset
            if setter is None:
                raise AttributeError(f"can't set attribute {name}")
            setter(self, value)
            return
        try:
            setattr(self.properties, name, value)
        except AttributeError:
            object.__setattr__(self, name, value)
