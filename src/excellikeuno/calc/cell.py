from __future__ import annotations

from typing import Any, cast

from ..core import InterfaceNames, UnoObject
from ..typing import BorderLine, TableBorder, XCell, XPropertySet


class CellProperties(UnoObject):
    """Wraps `com.sun.star.table.CellProperties` to offer attribute-style access."""

    def _props(self) -> XPropertySet:
        # The wrapped object is already an XPropertySet; avoid re-querying.
        return cast(XPropertySet, self.raw)

    def get_property(self, name: str) -> Any:
        return self._props().getPropertyValue(name)

    def set_property(self, name: str, value: Any) -> None:
        self._props().setPropertyValue(name, value)

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
        try:
            self.set_property(name, value)
        except Exception as exc:  # pragma: no cover - UNO failures bubble up
            raise AttributeError(f"Cannot set cell property: {name}") from exc


class Cell(UnoObject):
    @property
    def properties(self) -> CellProperties:
        existing = self.__dict__.get("_properties")
        if existing is None:
            existing = CellProperties(self.iface(InterfaceNames.X_PROPERTY_SET))
            object.__setattr__(self, "_properties", existing)
        return cast(CellProperties, existing)

    # CellProperties (public attributes) shortcuts for IDE completion
    @property
    def CellStyle(self) -> Any:
        return self.properties.get_property("CellStyle")

    @CellStyle.setter
    def CellStyle(self, value: Any) -> None:
        self.properties.set_property("CellStyle", value)

    @property
    def CellBackColor(self) -> Any:
        return self.properties.get_property("CellBackColor")

    @CellBackColor.setter
    def CellBackColor(self, value: Any) -> None:
        self.properties.set_property("CellBackColor", value)

    @property
    def IsCellBackgroundTransparent(self) -> bool:
        return bool(self.properties.get_property("IsCellBackgroundTransparent"))

    @IsCellBackgroundTransparent.setter
    def IsCellBackgroundTransparent(self, value: bool) -> None:
        self.properties.set_property("IsCellBackgroundTransparent", bool(value))

    @property
    def HoriJustify(self) -> Any:
        return self.properties.get_property("HoriJustify")

    @HoriJustify.setter
    def HoriJustify(self, value: Any) -> None:
        self.properties.set_property("HoriJustify", value)

    @property
    def VertJustify(self) -> Any:
        return self.properties.get_property("VertJustify")

    @VertJustify.setter
    def VertJustify(self, value: Any) -> None:
        self.properties.set_property("VertJustify", value)

    @property
    def IsTextWrapped(self) -> bool:
        return bool(self.properties.get_property("IsTextWrapped"))

    @IsTextWrapped.setter
    def IsTextWrapped(self, value: bool) -> None:
        self.properties.set_property("IsTextWrapped", bool(value))

    @property
    def ParaIndent(self) -> Any:
        return self.properties.get_property("ParaIndent")

    @ParaIndent.setter
    def ParaIndent(self, value: Any) -> None:
        self.properties.set_property("ParaIndent", value)

    @property
    def Orientation(self) -> Any:
        return self.properties.get_property("Orientation")

    @Orientation.setter
    def Orientation(self, value: Any) -> None:
        self.properties.set_property("Orientation", value)

    @property
    def RotateAngle(self) -> Any:
        return self.properties.get_property("RotateAngle")

    @RotateAngle.setter
    def RotateAngle(self, value: Any) -> None:
        self.properties.set_property("RotateAngle", value)

    @property
    def RotateReference(self) -> Any:
        return self.properties.get_property("RotateReference")

    @RotateReference.setter
    def RotateReference(self, value: Any) -> None:
        self.properties.set_property("RotateReference", value)

    @property
    def AsianVerticalMode(self) -> bool:
        return bool(self.properties.get_property("AsianVerticalMode"))

    @AsianVerticalMode.setter
    def AsianVerticalMode(self, value: bool) -> None:
        self.properties.set_property("AsianVerticalMode", bool(value))

    @property
    def TableBorder(self) -> TableBorder:
        return cast(TableBorder, self.properties.get_property("TableBorder"))

    @TableBorder.setter
    def TableBorder(self, value: TableBorder) -> None:
        self.properties.set_property("TableBorder", value)

    @property
    def TopBorder(self) -> BorderLine:
        return cast(BorderLine, self.properties.get_property("TopBorder"))

    @TopBorder.setter
    def TopBorder(self, value: BorderLine) -> None:
        self.properties.set_property("TopBorder", value)

    @property
    def BottomBorder(self) -> BorderLine:
        return cast(BorderLine, self.properties.get_property("BottomBorder"))

    @BottomBorder.setter
    def BottomBorder(self, value: BorderLine) -> None:
        self.properties.set_property("BottomBorder", value)

    @property
    def LeftBorder(self) -> BorderLine:
        return cast(BorderLine, self.properties.get_property("LeftBorder"))

    @LeftBorder.setter
    def LeftBorder(self, value: BorderLine) -> None:
        self.properties.set_property("LeftBorder", value)

    @property
    def RightBorder(self) -> BorderLine:
        return cast(BorderLine, self.properties.get_property("RightBorder"))

    @RightBorder.setter
    def RightBorder(self, value: BorderLine) -> None:
        self.properties.set_property("RightBorder", value)

    @property
    def NumberFormat(self) -> Any:
        return self.properties.get_property("NumberFormat")

    @NumberFormat.setter
    def NumberFormat(self, value: Any) -> None:
        self.properties.set_property("NumberFormat", value)

    @property
    def ShadowFormat(self) -> Any:
        return self.properties.get_property("ShadowFormat")

    @ShadowFormat.setter
    def ShadowFormat(self, value: Any) -> None:
        self.properties.set_property("ShadowFormat", value)

    @property
    def CellProtection(self) -> Any:
        return self.properties.get_property("CellProtection")

    @CellProtection.setter
    def CellProtection(self, value: Any) -> None:
        self.properties.set_property("CellProtection", value)

    @property
    def UserDefinedAttributes(self) -> Any:
        return self.properties.get_property("UserDefinedAttributes")

    @UserDefinedAttributes.setter
    def UserDefinedAttributes(self, value: Any) -> None:
        self.properties.set_property("UserDefinedAttributes", value)

    @property
    def DiagonalTLBR(self) -> Any:
        return self.properties.get_property("DiagonalTLBR")

    @DiagonalTLBR.setter
    def DiagonalTLBR(self, value: Any) -> None:
        self.properties.set_property("DiagonalTLBR", value)

    @property
    def DiagonalBLTR(self) -> Any:
        return self.properties.get_property("DiagonalBLTR")

    @DiagonalBLTR.setter
    def DiagonalBLTR(self, value: Any) -> None:
        self.properties.set_property("DiagonalBLTR", value)

    @property
    def ShrinkToFit(self) -> bool:
        return bool(self.properties.get_property("ShrinkToFit"))

    @ShrinkToFit.setter
    def ShrinkToFit(self, value: bool) -> None:
        self.properties.set_property("ShrinkToFit", bool(value))

    @property
    def TableBorder2(self) -> Any:
        return self.properties.get_property("TableBorder2")

    @TableBorder2.setter
    def TableBorder2(self, value: Any) -> None:
        self.properties.set_property("TableBorder2", value)

    @property
    def TopBorder2(self) -> Any:
        return self.properties.get_property("TopBorder2")

    @TopBorder2.setter
    def TopBorder2(self, value: Any) -> None:
        self.properties.set_property("TopBorder2", value)

    @property
    def BottomBorder2(self) -> Any:
        return self.properties.get_property("BottomBorder2")

    @BottomBorder2.setter
    def BottomBorder2(self, value: Any) -> None:
        self.properties.set_property("BottomBorder2", value)

    @property
    def LeftBorder2(self) -> Any:
        return self.properties.get_property("LeftBorder2")

    @LeftBorder2.setter
    def LeftBorder2(self, value: Any) -> None:
        self.properties.set_property("LeftBorder2", value)

    @property
    def RightBorder2(self) -> Any:
        return self.properties.get_property("RightBorder2")

    @RightBorder2.setter
    def RightBorder2(self, value: Any) -> None:
        self.properties.set_property("RightBorder2", value)

    @property
    def DiagonalTLBR2(self) -> Any:
        return self.properties.get_property("DiagonalTLBR2")

    @DiagonalTLBR2.setter
    def DiagonalTLBR2(self, value: Any) -> None:
        self.properties.set_property("DiagonalTLBR2", value)

    @property
    def DiagonalBLTR2(self) -> Any:
        return self.properties.get_property("DiagonalBLTR2")

    @DiagonalBLTR2.setter
    def DiagonalBLTR2(self, value: Any) -> None:
        self.properties.set_property("DiagonalBLTR2", value)

    @property
    def CellInteropGrabBag(self) -> Any:
        return self.properties.get_property("CellInteropGrabBag")

    @CellInteropGrabBag.setter
    def CellInteropGrabBag(self, value: Any) -> None:
        self.properties.set_property("CellInteropGrabBag", value)

    @property
    def value(self) -> float:
        cell = cast(XCell, self.iface(InterfaceNames.X_CELL))
        return cell.getValue()

    @value.setter
    def value(self, value: float) -> None:
        cell = cast(XCell, self.iface(InterfaceNames.X_CELL))
        cell.setValue(value)

    @property
    def formula(self) -> str:
        cell = cast(XCell, self.iface(InterfaceNames.X_CELL))
        return cell.getFormula()

    @formula.setter
    def formula(self, formula: str) -> None:
        cell = cast(XCell, self.iface(InterfaceNames.X_CELL))
        cell.setFormula(formula)

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
