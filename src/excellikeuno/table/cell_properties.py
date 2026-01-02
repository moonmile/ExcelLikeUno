from __future__ import annotations

from typing import Any, cast

from ..core import UnoObject
from ..typing import (
    BorderLine,
    BorderLineStruct,
    BorderLine2,
    CellHoriJustify,
    CellOrientation,
    CellProtectionStruct,
    CellVertJustify,
    Color,
    ShadowFormatStruct,
    TableBorder,
    TableBorder2,
    XPropertySet,
)
from ..utils.structs import make_border_line


class CellProperties(UnoObject):
    """Wraps `com.sun.star.table.CellProperties` to offer attribute-style access."""

    def _props(self) -> XPropertySet:
        # The wrapped object is already an XPropertySet; avoid re-querying.
        return cast(XPropertySet, self.raw)

    def get_property(self, name: str) -> Any:
        return self._props().getPropertyValue(name)

    def set_property(self, name: str, value: Any) -> None:
        self._props().setPropertyValue(name, value)

    # Commonly used properties for convenience
    @property
    def CellStyle(self) -> str:
        return cast(str, self.get_property("CellStyle"))

    @CellStyle.setter
    def CellStyle(self, value: str) -> None:
        self.set_property("CellStyle", value)

    @property
    def CellBackColor(self) -> Color:
        return cast(Color, self.get_property("CellBackColor"))

    @CellBackColor.setter
    def CellBackColor(self, value: Color) -> None:
        self.set_property("CellBackColor", value)

    @property
    def IsCellBackgroundTransparent(self) -> bool:
        return bool(self.get_property("IsCellBackgroundTransparent"))

    @IsCellBackgroundTransparent.setter
    def IsCellBackgroundTransparent(self, value: bool) -> None:
        self.set_property("IsCellBackgroundTransparent", bool(value))

    @property
    def HoriJustify(self) -> CellHoriJustify:
        return CellHoriJustify(int(self.get_property("HoriJustify")))

    @HoriJustify.setter
    def HoriJustify(self, value: CellHoriJustify | int) -> None:
        self.set_property("HoriJustify", int(value))

    @property
    def VertJustify(self) -> CellVertJustify:
        return CellVertJustify(int(self.get_property("VertJustify")))

    @VertJustify.setter
    def VertJustify(self, value: CellVertJustify | int) -> None:
        self.set_property("VertJustify", int(value))

    @property
    def IsTextWrapped(self) -> bool:
        return bool(self.get_property("IsTextWrapped"))

    @IsTextWrapped.setter
    def IsTextWrapped(self, value: bool) -> None:
        self.set_property("IsTextWrapped", bool(value))

    @property
    def ParaIndent(self) -> int:
        return int(self.get_property("ParaIndent"))

    @ParaIndent.setter
    def ParaIndent(self, value: int) -> None:
        self.set_property("ParaIndent", int(value))

    @property
    def Orientation(self) -> CellOrientation:
        return CellOrientation(int(self.get_property("Orientation")))

    @Orientation.setter
    def Orientation(self, value: CellOrientation | int) -> None:
        self.set_property("Orientation", int(value))

    @property
    def RotateAngle(self) -> int:
        return int(self.get_property("RotateAngle"))

    @RotateAngle.setter
    def RotateAngle(self, value: int) -> None:
        self.set_property("RotateAngle", int(value))

    @property
    def RotateReference(self) -> int:
        return int(self.get_property("RotateReference"))

    @RotateReference.setter
    def RotateReference(self, value: int) -> None:
        self.set_property("RotateReference", int(value))

    @property
    def AsianVerticalMode(self) -> bool:
        return bool(self.get_property("AsianVerticalMode"))

    @AsianVerticalMode.setter
    def AsianVerticalMode(self, value: bool) -> None:
        self.set_property("AsianVerticalMode", bool(value))

    @property
    def TableBorder(self) -> TableBorder:
        return cast(TableBorder, self.get_property("TableBorder"))

    @TableBorder.setter
    def TableBorder(self, value: TableBorder) -> None:
        self.set_property("TableBorder", value)

    @property
    def TopBorder(self) -> BorderLine:
        return cast(BorderLine, self.get_property("TopBorder"))

    @TopBorder.setter
    def TopBorder(self, value: BorderLine | BorderLineStruct) -> None:
        if isinstance(value, BorderLineStruct):
            value = make_border_line(
                color=value.Color,
                inner_line_width=value.InnerLineWidth,
                outer_line_width=value.OuterLineWidth,
                line_distance=value.LineDistance,
            )
        self.set_property("TopBorder", value)

    @property
    def BottomBorder(self) -> BorderLine:
        return cast(BorderLine, self.get_property("BottomBorder"))

    @BottomBorder.setter
    def BottomBorder(self, value: BorderLine | BorderLineStruct) -> None:
        if isinstance(value, BorderLineStruct):
            value = make_border_line(
                color=value.Color,
                inner_line_width=value.InnerLineWidth,
                outer_line_width=value.OuterLineWidth,
                line_distance=value.LineDistance,
            )
        self.set_property("BottomBorder", value)

    @property
    def LeftBorder(self) -> BorderLine:
        return cast(BorderLine, self.get_property("LeftBorder"))

    @LeftBorder.setter
    def LeftBorder(self, value: BorderLine | BorderLineStruct) -> None:
        if isinstance(value, BorderLineStruct):
            value = make_border_line(
                color=value.Color,
                inner_line_width=value.InnerLineWidth,
                outer_line_width=value.OuterLineWidth,
                line_distance=value.LineDistance,
            )
        self.set_property("LeftBorder", value)

    @property
    def RightBorder(self) -> BorderLine:
        return cast(BorderLine, self.get_property("RightBorder"))

    @RightBorder.setter
    def RightBorder(self, value: BorderLine | BorderLineStruct) -> None:
        if isinstance(value, BorderLineStruct):
            value = make_border_line(
                color=value.Color,
                inner_line_width=value.InnerLineWidth,
                outer_line_width=value.OuterLineWidth,
                line_distance=value.LineDistance,
            )
        self.set_property("RightBorder", value)

    @property
    def NumberFormat(self) -> int:
        return int(self.get_property("NumberFormat"))

    @NumberFormat.setter
    def NumberFormat(self, value: int) -> None:
        self.set_property("NumberFormat", int(value))

    @property
    def ShadowFormat(self) -> ShadowFormatStruct:
        return cast(ShadowFormatStruct, self.get_property("ShadowFormat"))

    @ShadowFormat.setter
    def ShadowFormat(self, value: ShadowFormatStruct) -> None:
        self.set_property("ShadowFormat", value)

    @property
    def CellProtection(self) -> CellProtectionStruct:
        return cast(CellProtectionStruct, self.get_property("CellProtection"))

    @CellProtection.setter
    def CellProtection(self, value: CellProtectionStruct) -> None:
        self.set_property("CellProtection", value)

    @property
    def UserDefinedAttributes(self) -> Any:
        return self.get_property("UserDefinedAttributes")

    @UserDefinedAttributes.setter
    def UserDefinedAttributes(self, value: Any) -> None:
        self.set_property("UserDefinedAttributes", value)

    @property
    def DiagonalTLBR(self) -> BorderLine:
        return cast(BorderLine, self.get_property("DiagonalTLBR"))

    @DiagonalTLBR.setter
    def DiagonalTLBR(self, value: BorderLine) -> None:
        self.set_property("DiagonalTLBR", value)

    @property
    def DiagonalBLTR(self) -> BorderLine:
        return cast(BorderLine, self.get_property("DiagonalBLTR"))

    @DiagonalBLTR.setter
    def DiagonalBLTR(self, value: BorderLine) -> None:
        self.set_property("DiagonalBLTR", value)

    @property
    def ShrinkToFit(self) -> bool:
        return bool(self.get_property("ShrinkToFit"))

    @ShrinkToFit.setter
    def ShrinkToFit(self, value: bool) -> None:
        self.set_property("ShrinkToFit", bool(value))

    @property
    def TableBorder2(self) -> TableBorder2:
        return cast(TableBorder2, self.get_property("TableBorder2"))

    @TableBorder2.setter
    def TableBorder2(self, value: TableBorder2) -> None:
        self.set_property("TableBorder2", value)

    @property
    def TopBorder2(self) -> BorderLine2:
        return cast(BorderLine2, self.get_property("TopBorder2"))

    @TopBorder2.setter
    def TopBorder2(self, value: BorderLine2) -> None:
        self.set_property("TopBorder2", value)

    @property
    def BottomBorder2(self) -> BorderLine2:
        return cast(BorderLine2, self.get_property("BottomBorder2"))

    @BottomBorder2.setter
    def BottomBorder2(self, value: BorderLine2) -> None:
        self.set_property("BottomBorder2", value)

    @property
    def LeftBorder2(self) -> BorderLine2:
        return cast(BorderLine2, self.get_property("LeftBorder2"))

    @LeftBorder2.setter
    def LeftBorder2(self, value: BorderLine2) -> None:
        self.set_property("LeftBorder2", value)

    @property
    def RightBorder2(self) -> BorderLine2:
        return cast(BorderLine2, self.get_property("RightBorder2"))

    @RightBorder2.setter
    def RightBorder2(self, value: BorderLine2) -> None:
        self.set_property("RightBorder2", value)

    @property
    def DiagonalTLBR2(self) -> BorderLine2:
        return cast(BorderLine2, self.get_property("DiagonalTLBR2"))

    @DiagonalTLBR2.setter
    def DiagonalTLBR2(self, value: BorderLine2) -> None:
        self.set_property("DiagonalTLBR2", value)

    @property
    def DiagonalBLTR2(self) -> BorderLine2:
        return cast(BorderLine2, self.get_property("DiagonalBLTR2"))

    @DiagonalBLTR2.setter
    def DiagonalBLTR2(self, value: BorderLine2) -> None:
        self.set_property("DiagonalBLTR2", value)

    @property
    def CellInteropGrabBag(self) -> Any:
        return self.get_property("CellInteropGrabBag")

    @CellInteropGrabBag.setter
    def CellInteropGrabBag(self, value: Any) -> None:
        self.set_property("CellInteropGrabBag", value)

    # keep UNO-style methods for compatibility
    def getPropertyValue(self, name: str) -> Any:  # noqa: N802 - UNO naming
        return self.get_property(name)

    def setPropertyValue(self, name: str, value: Any) -> None:  # noqa: N802 - UNO naming
        self.set_property(name, value)

    def __getattr__(self, name: str) -> Any:
        try:
            return self.get_property(name)
        except Exception as exc:  # pragma: no cover - UNO failures bubble up
            raise AttributeError(f"Unknown cell property: {name}") from exc

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
        except Exception as exc:  # pragma: no cover - UNO failures bubble up
            raise AttributeError(f"Cannot set cell property: {name}") from exc
