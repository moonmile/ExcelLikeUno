from __future__ import annotations

from typing import Any, cast

from excellikeuno.typing.structs import Point, Size

from ..core import UnoObject
from ..typing import Color, InterfaceNames, LineDash, LineStyle, XPropertySet, XShape
from .fill_properties import FillProperties
from .line_properties import LineProperties
from .shadow_properties import ShadowProperties
from .text_properties import TextProperties
from ..style.font import Font


class Shape(UnoObject):
    """Wraps a drawing Shape from Calc draw page."""

    def _get_prop(self, name: str) -> Any:
        props = cast(XPropertySet, self.iface(InterfaceNames.X_PROPERTY_SET))
        return props.getPropertyValue(name)

    def _set_prop(self, name: str, value: Any) -> None:
        props = cast(XPropertySet, self.iface(InterfaceNames.X_PROPERTY_SET))
        props.setPropertyValue(name, value)

    @property
    def Position(self) -> Point:
        shape = cast(XShape, self.iface(InterfaceNames.X_SHAPE))
        pos = shape.getPosition()
        try:
            return Point(pos.X, pos.Y)
        except Exception:
            return pos

    @Position.setter
    def Position(self, value: Point) -> None:
        shape = cast(XShape, self.iface(InterfaceNames.X_SHAPE))
        target = value
        if hasattr(value, "to_raw"):
            try:
                target = value.to_raw()
            except Exception:
                target = value
        shape.setPosition(target)

    @property
    def Size(self) -> Size:
        shape = cast(XShape, self.iface(InterfaceNames.X_SHAPE))
        size = shape.getSize()
        try:
            return Size(size.Width, size.Height)
        except Exception:
            return size

    @Size.setter
    def Size(self, value: Size) -> None:
        shape = cast(XShape, self.iface(InterfaceNames.X_SHAPE))
        target = value
        if hasattr(value, "to_raw"):
            try:
                target = value.to_raw()
            except Exception:
                target = value
        shape.setSize(target)

    @property
    def props(self) -> XPropertySet:
        return cast(XPropertySet, self.iface(InterfaceNames.X_PROPERTY_SET))

    @property
    def fill_properties(self) -> FillProperties:
        existing = self.__dict__.get("_fill_properties")
        if existing is None:
            try:
                fp_raw = self.iface(InterfaceNames.FILL_PROPERTIES)
            except BaseException:
                fp_raw = self.raw
            existing = FillProperties(fp_raw)
            object.__setattr__(self, "_fill_properties", existing)
        return cast(FillProperties, existing)

    @property
    def line_properties(self) -> LineProperties:
        existing = self.__dict__.get("_line_properties")
        if existing is None:
            try:
                lp_raw = self.iface(InterfaceNames.LINE_PROPERTIES)
            except BaseException:
                lp_raw = self.raw
            existing = LineProperties(lp_raw)
            object.__setattr__(self, "_line_properties", existing)
        return cast(LineProperties, existing)

    @property
    def shadow_properties(self) -> ShadowProperties:
        existing = self.__dict__.get("_shadow_properties")
        if existing is None:
            existing = ShadowProperties(self.iface(InterfaceNames.SHADOW_PROPERTIES))
            object.__setattr__(self, "_shadow_properties", existing)
        return cast(ShadowProperties, existing)

    # TextProperties implementation
    @property
    def text_properties(self) -> TextProperties:
        existing = self.__dict__.get("_text_properties")
        if existing is None:
            try:
                tp_raw = self.iface(InterfaceNames.TEXT_PROPERTIES)
            except BaseException:
                tp_raw = self.raw
            existing = TextProperties(tp_raw)
            object.__setattr__(self, "_text_properties", existing)
        return cast(TextProperties, existing)

    @property
    def IsNumbering(self) -> bool:
        return bool(self.text_properties.IsNumbering)

    @IsNumbering.setter
    def IsNumbering(self, value: bool) -> None:
        self.text_properties.IsNumbering = bool(value)

    @property
    def TextAutoGrowHeight(self) -> bool:
        return bool(self.text_properties.TextAutoGrowHeight)

    @TextAutoGrowHeight.setter
    def TextAutoGrowHeight(self, value: bool) -> None:
        self.text_properties.TextAutoGrowHeight = bool(value)

    @property
    def TextAutoGrowWidth(self) -> bool:
        return bool(self.text_properties.TextAutoGrowWidth)

    @TextAutoGrowWidth.setter
    def TextAutoGrowWidth(self, value: bool) -> None:
        self.text_properties.TextAutoGrowWidth = bool(value)

    @property
    def TextLeftDistance(self) -> int:
        return int(self.text_properties.TextLeftDistance)

    @TextLeftDistance.setter
    def TextLeftDistance(self, value: int) -> None:
        self.text_properties.TextLeftDistance = int(value)

    @property
    def TextRightDistance(self) -> int:
        return int(self.text_properties.TextRightDistance)

    @TextRightDistance.setter
    def TextRightDistance(self, value: int) -> None:
        self.text_properties.TextRightDistance = int(value)

    @property
    def TextUpperDistance(self) -> int:
        return int(self.text_properties.TextUpperDistance)

    @TextUpperDistance.setter
    def TextUpperDistance(self, value: int) -> None:
        self.text_properties.TextUpperDistance = int(value)

    @property
    def TextLowerDistance(self) -> int:
        return int(self.text_properties.TextLowerDistance)

    @TextLowerDistance.setter
    def TextLowerDistance(self, value: int) -> None:
        self.text_properties.TextLowerDistance = int(value)

    @property
    def CharHeight(self) -> float:
        return float(self.text_properties.CharHeight)

    @CharHeight.setter
    def CharHeight(self, value: float) -> None:
        self.text_properties.CharHeight = float(value)

    @property
    def CharFontName(self) -> str:
        return cast(str, self.text_properties.CharFontName)

    @CharFontName.setter
    def CharFontName(self, value: str) -> None:
        self.text_properties.CharFontName = value

    @property
    def HoriJustify(self) -> int:
        for name in ("TextHorizontalAdjust", "HoriJustify", "ParaAdjust"):
            try:
                return int(self.text_properties.get_property(name))
            except BaseException:
                continue
        return 0

    @HoriJustify.setter
    def HoriJustify(self, value: int) -> None:
        val = int(value)
        for name in ("TextHorizontalAdjust", "HoriJustify", "ParaAdjust"):
            try:
                self.text_properties.set_property(name, val)
                return
            except BaseException:
                continue

    @property
    def VertJustify(self) -> int:
        for name in ("TextVerticalAdjust", "VertJustify"):
            try:
                return int(self.text_properties.get_property(name))
            except BaseException:
                continue
        return 0

    @VertJustify.setter
    def VertJustify(self, value: int) -> None:
        val = int(value)
        for name in ("TextVerticalAdjust", "VertJustify"):
            try:
                self.text_properties.set_property(name, val)
                return
            except BaseException:
                continue

    def _font_getter(self) -> dict[str, Any]:
        esc = int(self.text_properties.CharEscapement)
        try:
            backcolor = self.text_properties.CharBackColor
        except Exception:
            backcolor = 0
        return {
            "name": self.text_properties.CharFontName,
            "size": float(self.text_properties.CharHeight),
            "bold": float(self.text_properties.CharWeight) >= 150.0,
            "italic": bool(int(self.text_properties.CharPosture)),
            "underline": int(self.text_properties.CharUnderline),
            "strikeout": int(self.text_properties.CharStrikeout),
            "color": self.text_properties.CharColor,
            "subscript": esc < 0,
            "superscript": esc > 0,
            "backcolor": backcolor,
            "font_style": int(self.text_properties.CharPosture),
            "strikthrough": int(self.text_properties.CharStrikeout) != 0,
        }

    def _font_setter(self, **updates: Any) -> None:
        tp = self.text_properties
        if "name" in updates:
            tp.CharFontName = updates["name"]
        if "size" in updates:
            tp.CharHeight = float(updates["size"])
        if "bold" in updates:
            tp.CharWeight = 150.0 if updates["bold"] else 100.0
        if "italic" in updates:
            tp.CharPosture = 2 if updates["italic"] else 0
        if "font_style" in updates:
            try:
                tp.CharPosture = int(updates["font_style"])
            except Exception:
                pass
        if "underline" in updates:
            tp.CharUnderline = int(updates["underline"])
        if "strikeout" in updates:
            tp.CharStrikeout = int(updates["strikeout"])
        if "color" in updates:
            tp.CharColor = updates["color"]
        if "backcolor" in updates:
            try:
                tp.set_property("CharBackTransparent", False)
            except Exception:
                pass
            try:
                tp.CharBackColor = updates["backcolor"]
            except Exception:
                pass
        if "subscript" in updates or "superscript" in updates:
            if updates.get("superscript"):
                tp.CharEscapement = 58
            elif updates.get("subscript"):
                tp.CharEscapement = -25
            else:
                tp.CharEscapement = 0
        if "strikthrough" in updates:
            try:
                tp.CharStrikeout = 1 if updates["strikthrough"] else 0
            except Exception:
                pass

    @property
    def font(self) -> Font:
        return Font(getter=self._font_getter, setter=self._font_setter)

    @font.setter
    def font(self, value: Font) -> None:
        # Accept a Font proxy or plain Font config dict-like
        current = {}
        try:
            current = value._current()  # type: ignore[attr-defined]
        except Exception:
            # fallback: if value is mapping-like
            try:
                current = dict(value)  # type: ignore[arg-type]
            except Exception:
                current = {}
        if not current:
            return
        self._font_setter(**current)

    # LineProperties implementation
    @property
    def LineColor(self) -> int:
        return int(self.line_properties.LineColor)

    @LineColor.setter
    def LineColor(self, value: int) -> None:
        self.line_properties.LineColor = int(value)

    @property
    def LineStyle(self) -> LineStyle:
        return LineStyle(int(self.line_properties.LineStyle))

    @LineStyle.setter
    def LineStyle(self, value: LineStyle) -> None:
        self.line_properties.LineStyle = int(value)

    @property
    def LineDashName(self) -> str:
        return cast(str, self.line_properties.LineDashName)

    @LineDashName.setter
    def LineDashName(self, value: str) -> None:
        self.line_properties.LineDashName = value

    @property
    def LineDash(self) -> LineDash:
        return cast(LineDash, self.line_properties.LineDash)

    @LineDash.setter
    def LineDash(self, value: LineDash) -> None:
        self.line_properties.LineDash = value

    @property
    def LineWidth(self) -> int:
        return int(self.line_properties.LineWidth)

    @LineWidth.setter
    def LineWidth(self, value: int) -> None:
        self.line_properties.LineWidth = int(value)

    # ShadowProperties implementation
    @property
    def Shadow(self) -> bool:
        return bool(self.shadow_properties.Shadow)

    @Shadow.setter
    def Shadow(self, value: bool) -> None:
        self.shadow_properties.Shadow = bool(value)

    @property
    def ShadowColor(self) -> int:
        return int(self.shadow_properties.ShadowColor)

    @ShadowColor.setter
    def ShadowColor(self, value: int) -> None:
        self.shadow_properties.ShadowColor = int(value)

    @property
    def ShadowTransparence(self) -> int:
        return int(self.shadow_properties.ShadowTransparence)

    @ShadowTransparence.setter
    def ShadowTransparence(self, value: int) -> None:
        self.shadow_properties.ShadowTransparence = int(value)

    @property
    def ShadowXDistance(self) -> int:
        return int(self.shadow_properties.ShadowXDistance)

    @ShadowXDistance.setter
    def ShadowXDistance(self, value: int) -> None:
        self.shadow_properties.ShadowXDistance = int(value)

    @property
    def ShadowYDistance(self) -> int:
        return int(self.shadow_properties.ShadowYDistance)

    @ShadowYDistance.setter
    def ShadowYDistance(self, value: int) -> None:
        self.shadow_properties.ShadowYDistance = int(value)

    # FillProperties implementation
    @property
    def FillColor(self) -> Color:
        return cast(Color, self.fill_properties.FillColor)

    @FillColor.setter
    def FillColor(self, value: Color) -> None:
        val = Color(value)
        solid = 1  # com.sun.star.drawing.FillStyle.SOLID
        try:
            self.fill_properties.FillStyle = solid
        except BaseException:
            pass
        self.fill_properties.FillColor = val

    @property
    def FillStyle(self) -> Any:
        return self.fill_properties.FillStyle

    @FillStyle.setter
    def FillStyle(self, value: Any) -> None:
        self.fill_properties.FillStyle = value

    @property
    def FillTransparence(self) -> int:
        return int(self.fill_properties.FillTransparence)

    @FillTransparence.setter
    def FillTransparence(self, value: int) -> None:
        self.fill_properties.FillTransparence = int(value)

    @property
    def FillGradientName(self) -> str:
        return cast(str, self.fill_properties.FillGradientName)

    @FillGradientName.setter
    def FillGradientName(self, value: str) -> None:
        self.fill_properties.FillGradientName = value

    @property
    def FillHatchName(self) -> str:
        return cast(str, self.fill_properties.FillHatchName)

    @FillHatchName.setter
    def FillHatchName(self, value: str) -> None:
        self.fill_properties.FillHatchName = value

    @property
    def FillBitmapName(self) -> str:
        return cast(str, self.fill_properties.FillBitmapName)

    @FillBitmapName.setter
    def FillBitmapName(self, value: str) -> None:
        self.fill_properties.FillBitmapName = value

    @property
    def FillBitmapMode(self) -> int:
        return int(self.fill_properties.FillBitmapMode)

    @FillBitmapMode.setter
    def FillBitmapMode(self, value: int) -> None:
        self.fill_properties.FillBitmapMode = int(value)

    @property
    def FillBitmapOffsetX(self) -> int:
        return int(self.fill_properties.FillBitmapOffsetX)

    @FillBitmapOffsetX.setter
    def FillBitmapOffsetX(self, value: int) -> None:
        self.fill_properties.FillBitmapOffsetX = int(value)

    @property
    def FillBitmapOffsetY(self) -> int:
        return int(self.fill_properties.FillBitmapOffsetY)

    @FillBitmapOffsetY.setter
    def FillBitmapOffsetY(self, value: int) -> None:
        self.fill_properties.FillBitmapOffsetY = int(value)

    @property
    def FillBitmapPositionX(self) -> int:
        return int(self.fill_properties.FillBitmapPositionX)

    @FillBitmapPositionX.setter
    def FillBitmapPositionX(self, value: int) -> None:
        self.fill_properties.FillBitmapPositionX = int(value)

    @property
    def FillBitmapPositionY(self) -> int:
        return int(self.fill_properties.FillBitmapPositionY)

    @FillBitmapPositionY.setter
    def FillBitmapPositionY(self, value: int) -> None:
        self.fill_properties.FillBitmapPositionY = int(value)

    @property
    def FillBitmapSizeX(self) -> int:
        return int(self.fill_properties.FillBitmapSizeX)

    @FillBitmapSizeX.setter
    def FillBitmapSizeX(self, value: int) -> None:
        self.fill_properties.FillBitmapSizeX = int(value)

    @property
    def FillBitmapSizeY(self) -> int:
        return int(self.fill_properties.FillBitmapSizeY)

    @FillBitmapSizeY.setter
    def FillBitmapSizeY(self, value: int) -> None:
        self.fill_properties.FillBitmapSizeY = int(value)

    @property
    def FillBackground(self) -> bool:
        return bool(self.fill_properties.FillBackground)

    @FillBackground.setter
    def FillBackground(self, value: bool) -> None:
        self.fill_properties.FillBackground = bool(value)

    # Common Shape properties
    @property
    def Name(self) -> str:
        return cast(str, self._get_prop("Name"))

    @Name.setter
    def Name(self, value: str) -> None:
        self._set_prop("Name", value)

    @property
    def Title(self) -> str:
        return cast(str, self._get_prop("Title"))

    @Title.setter
    def Title(self, value: str) -> None:
        self._set_prop("Title", value)

    @property
    def Description(self) -> str:
        return cast(str, self._get_prop("Description"))

    @Description.setter
    def Description(self, value: str) -> None:
        self._set_prop("Description", value)

    @property
    def Visible(self) -> bool:
        return bool(self._get_prop("Visible"))

    @Visible.setter
    def Visible(self, value: bool) -> None:
        self._set_prop("Visible", bool(value))

    @property
    def Printable(self) -> bool:
        return bool(self._get_prop("Printable"))

    @Printable.setter
    def Printable(self, value: bool) -> None:
        self._set_prop("Printable", bool(value))

    @property
    def MoveProtect(self) -> bool:
        return bool(self._get_prop("MoveProtect"))

    @MoveProtect.setter
    def MoveProtect(self, value: bool) -> None:
        self._set_prop("MoveProtect", bool(value))

    @property
    def SizeProtect(self) -> bool:
        return bool(self._get_prop("SizeProtect"))

    @SizeProtect.setter
    def SizeProtect(self, value: bool) -> None:
        self._set_prop("SizeProtect", bool(value))

    @property
    def ZOrder(self) -> int:
        return int(self._get_prop("ZOrder"))

    @ZOrder.setter
    def ZOrder(self, value: int) -> None:
        self._set_prop("ZOrder", int(value))

    @property
    def LayerID(self) -> int:
        return int(self._get_prop("LayerID"))

    @LayerID.setter
    def LayerID(self, value: int) -> None:
        self._set_prop("LayerID", int(value))

    @property
    def LayerName(self) -> str:
        return cast(str, self._get_prop("LayerName"))

    @LayerName.setter
    def LayerName(self, value: str) -> None:
        self._set_prop("LayerName", value)

    @property
    def Hyperlink(self) -> str:
        return cast(str, self._get_prop("Hyperlink"))

    @Hyperlink.setter
    def Hyperlink(self, value: str) -> None:
        self._set_prop("Hyperlink", value)

    @property
    def NavigationOrder(self) -> int:
        return int(self._get_prop("NavigationOrder"))

    @NavigationOrder.setter
    def NavigationOrder(self, value: int) -> None:
        self._set_prop("NavigationOrder", int(value))

    @property
    def Style(self) -> Any:
        return self._get_prop("Style")

    @Style.setter
    def Style(self, value: Any) -> None:
        self._set_prop("Style", value)

    @property
    def Transformation(self) -> Any:
        return self._get_prop("Transformation")

    @Transformation.setter
    def Transformation(self, value: Any) -> None:
        self._set_prop("Transformation", value)

    @property
    def ShapeUserDefinedAttributes(self) -> Any:
        return self._get_prop("ShapeUserDefinedAttributes")

    @ShapeUserDefinedAttributes.setter
    def ShapeUserDefinedAttributes(self, value: Any) -> None:
        self._set_prop("ShapeUserDefinedAttributes", value)

    @property
    def RelativeHeight(self) -> int:
        return int(self._get_prop("RelativeHeight"))

    @RelativeHeight.setter
    def RelativeHeight(self, value: int) -> None:
        self._set_prop("RelativeHeight", int(value))

    @property
    def RelativeWidth(self) -> int:
        return int(self._get_prop("RelativeWidth"))

    @RelativeWidth.setter
    def RelativeWidth(self, value: int) -> None:
        self._set_prop("RelativeWidth", int(value))

    @property
    def RelativeHeightRelation(self) -> int:
        return int(self._get_prop("RelativeHeightRelation"))

    @RelativeHeightRelation.setter
    def RelativeHeightRelation(self, value: int) -> None:
        self._set_prop("RelativeHeightRelation", int(value))

    @property
    def RelativeWidthRelation(self) -> int:
        return int(self._get_prop("RelativeWidthRelation"))

    @RelativeWidthRelation.setter
    def RelativeWidthRelation(self, value: int) -> None:
        self._set_prop("RelativeWidthRelation", int(value))

    @property
    def Decorative(self) -> bool:
        return bool(self._get_prop("Decorative"))

    @Decorative.setter
    def Decorative(self, value: bool) -> None:
        self._set_prop("Decorative", bool(value))

    @property
    def InteropGrabBag(self) -> Any:
        return self._get_prop("InteropGrabBag")

    @InteropGrabBag.setter
    def InteropGrabBag(self, value: Any) -> None:
        self._set_prop("InteropGrabBag", value)

    # Text anchoring and wrap related properties
    @property
    def AnchorPageNo(self) -> int:
        return int(self._get_prop("AnchorPageNo"))

    @AnchorPageNo.setter
    def AnchorPageNo(self, value: int) -> None:
        self._set_prop("AnchorPageNo", int(value))

    @property
    def AnchorType(self) -> Any:
        return self._get_prop("AnchorType")

    @AnchorType.setter
    def AnchorType(self, value: Any) -> None:
        self._set_prop("AnchorType", value)

    @property
    def AnchorFrame(self) -> Any:
        return self._get_prop("AnchorFrame")

    @AnchorFrame.setter
    def AnchorFrame(self, value: Any) -> None:
        self._set_prop("AnchorFrame", value)

    @property
    def TextRange(self) -> Any:
        return self._get_prop("TextRange")

    @TextRange.setter
    def TextRange(self, value: Any) -> None:
        self._set_prop("TextRange", value)

    @property
    def Surround(self) -> Any:
        return self._get_prop("Surround")

    @Surround.setter
    def Surround(self, value: Any) -> None:
        self._set_prop("Surround", value)

    @property
    def SurroundAnchorOnly(self) -> bool:
        return bool(self._get_prop("SurroundAnchorOnly"))

    @SurroundAnchorOnly.setter
    def SurroundAnchorOnly(self, value: bool) -> None:
        self._set_prop("SurroundAnchorOnly", bool(value))

    @property
    def SurroundContour(self) -> bool:
        return bool(self._get_prop("SurroundContour"))

    @SurroundContour.setter
    def SurroundContour(self, value: bool) -> None:
        self._set_prop("SurroundContour", bool(value))

    @property
    def ContourOutside(self) -> bool:
        return bool(self._get_prop("ContourOutside"))

    @ContourOutside.setter
    def ContourOutside(self, value: bool) -> None:
        self._set_prop("ContourOutside", bool(value))

    @property
    def Opaque(self) -> bool:
        return bool(self._get_prop("Opaque"))

    @Opaque.setter
    def Opaque(self, value: bool) -> None:
        self._set_prop("Opaque", bool(value))

    @property
    def WrapInfluenceOnPosition(self) -> int:
        return int(self._get_prop("WrapInfluenceOnPosition"))

    @WrapInfluenceOnPosition.setter
    def WrapInfluenceOnPosition(self, value: int) -> None:
        self._set_prop("WrapInfluenceOnPosition", int(value))

    # Orientation and positioning
    @property
    def HoriOrient(self) -> int:
        return int(self._get_prop("HoriOrient"))

    @HoriOrient.setter
    def HoriOrient(self, value: int) -> None:
        self._set_prop("HoriOrient", int(value))

    @property
    def HoriOrientPosition(self) -> int:
        return int(self._get_prop("HoriOrientPosition"))

    @HoriOrientPosition.setter
    def HoriOrientPosition(self, value: int) -> None:
        self._set_prop("HoriOrientPosition", int(value))

    @property
    def HoriOrientRelation(self) -> int:
        return int(self._get_prop("HoriOrientRelation"))

    @HoriOrientRelation.setter
    def HoriOrientRelation(self, value: int) -> None:
        self._set_prop("HoriOrientRelation", int(value))

    @property
    def VertOrient(self) -> int:
        return int(self._get_prop("VertOrient"))

    @VertOrient.setter
    def VertOrient(self, value: int) -> None:
        self._set_prop("VertOrient", int(value))

    @property
    def VertOrientPosition(self) -> int:
        return int(self._get_prop("VertOrientPosition"))

    @VertOrientPosition.setter
    def VertOrientPosition(self, value: int) -> None:
        self._set_prop("VertOrientPosition", int(value))

    @property
    def VertOrientRelation(self) -> int:
        return int(self._get_prop("VertOrientRelation"))

    @VertOrientRelation.setter
    def VertOrientRelation(self, value: int) -> None:
        self._set_prop("VertOrientRelation", int(value))

    @property
    def LeftMargin(self) -> int:
        return int(self._get_prop("LeftMargin"))

    @LeftMargin.setter
    def LeftMargin(self, value: int) -> None:
        self._set_prop("LeftMargin", int(value))

    @property
    def RightMargin(self) -> int:
        return int(self._get_prop("RightMargin"))

    @RightMargin.setter
    def RightMargin(self, value: int) -> None:
        self._set_prop("RightMargin", int(value))

    @property
    def TopMargin(self) -> int:
        return int(self._get_prop("TopMargin"))

    @TopMargin.setter
    def TopMargin(self, value: int) -> None:
        self._set_prop("TopMargin", int(value))

    @property
    def BottomMargin(self) -> int:
        return int(self._get_prop("BottomMargin"))

    @BottomMargin.setter
    def BottomMargin(self, value: int) -> None:
        self._set_prop("BottomMargin", int(value))

    # Layout direction helpers
    @property
    def TransformationInHoriL2R(self) -> Any:
        return self._get_prop("TransformationInHoriL2R")

    @property
    def PositionLayoutDir(self) -> int:
        return int(self._get_prop("PositionLayoutDir"))

    @PositionLayoutDir.setter
    def PositionLayoutDir(self, value: int) -> None:
        self._set_prop("PositionLayoutDir", int(value))

    @property
    def StartPositionInHoriL2R(self) -> Any:
        return self._get_prop("StartPositionInHoriL2R")

    @property
    def EndPositionInHoriL2R(self) -> Any:
        return self._get_prop("EndPositionInHoriL2R")

    # Overlap handling
    @property
    def AllowOverlap(self) -> bool:
        return bool(self._get_prop("AllowOverlap"))

    @AllowOverlap.setter
    def AllowOverlap(self, value: bool) -> None:
        self._set_prop("AllowOverlap", bool(value))

    @property
    def string(self) -> str:
        try:
            value = self._get_prop("String")
            return str(value) if value is not None else ""
        except Exception:
            try:
                return str(self.raw.getText().getString())  # type: ignore[attr-defined]
            except Exception:
                return ""

    @string.setter
    def string(self, value: str) -> None:
        try:
            self._set_prop("String", value)
            return
        except Exception:
            pass
        try:
            self.raw.setString(value)  # type: ignore[attr-defined]
            return
        except Exception:
            pass
        try:
            self.raw.getText().setString(value)  # type: ignore[attr-defined]
        except Exception:
            # Best effort; swallow if unsupported
            pass

    @property
    def String(self) -> str:
        return self.string

    @String.setter
    def String(self, value: str) -> None:
        self.string = value

