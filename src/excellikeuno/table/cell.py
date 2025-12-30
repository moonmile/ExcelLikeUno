from __future__ import annotations

from typing import Any, cast

from ..core import UnoObject
from ..typing import (
    BorderLine,
    BorderLine2,
    CellHoriJustify,
    CellVertJustify,
    InterfaceNames,
    TableBorder,
    TableBorder2,
    XCell,
    XPropertySet,
)


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

    # CellProperties (public attributes) shortcuts for IDE completion
    @property
    def CellStyle(self) -> Any:
        return self._get_prop("CellStyle")

    @CellStyle.setter
    def CellStyle(self, value: Any) -> None:
        self._set_prop("CellStyle", value)

    @property
    def CellBackColor(self) -> Any:
        return self._get_prop("CellBackColor")

    @CellBackColor.setter
    def CellBackColor(self, value: Any) -> None:
        self._set_prop("CellBackColor", value)

    @property
    def IsCellBackgroundTransparent(self) -> bool:
        return bool(self._get_prop("IsCellBackgroundTransparent"))

    @IsCellBackgroundTransparent.setter
    def IsCellBackgroundTransparent(self, value: bool) -> None:
        self._set_prop("IsCellBackgroundTransparent", bool(value))

    @property
    def HoriJustify(self) -> CellHoriJustify:
        raw = self._get_prop("HoriJustify")
        if isinstance(raw, CellHoriJustify):
            return raw
        if hasattr(raw, "value"):
            val = getattr(raw, "value")
        else:
            val = raw
        try:
            return CellHoriJustify(int(val))
        except Exception:
            return CellHoriJustify[val] if isinstance(val, str) else CellHoriJustify(val)

    @HoriJustify.setter
    def HoriJustify(self, value: CellHoriJustify | int) -> None:
        self._set_prop("HoriJustify", int(value))

    @property
    def VertJustify(self) -> CellVertJustify:
        raw = self._get_prop("VertJustify")
        if isinstance(raw, CellVertJustify):
            return raw
        if hasattr(raw, "value"):
            val = getattr(raw, "value")
        else:
            val = raw
        try:
            return CellVertJustify(int(val))
        except Exception:
            return CellVertJustify[val] if isinstance(val, str) else CellVertJustify(val)

    @VertJustify.setter
    def VertJustify(self, value: CellVertJustify | int) -> None:
        self._set_prop("VertJustify", int(value))

    @property
    def IsTextWrapped(self) -> bool:
        return bool(self._get_prop("IsTextWrapped"))

    @IsTextWrapped.setter
    def IsTextWrapped(self, value: bool) -> None:
        self._set_prop("IsTextWrapped", bool(value))

    @property
    def ParaIndent(self) -> Any:
        return self._get_prop("ParaIndent")

    @ParaIndent.setter
    def ParaIndent(self, value: Any) -> None:
        self._set_prop("ParaIndent", value)

    @property
    def Orientation(self) -> Any:
        return self._get_prop("Orientation")

    @Orientation.setter
    def Orientation(self, value: Any) -> None:
        self._set_prop("Orientation", value)

    @property
    def RotateAngle(self) -> Any:
        return self._get_prop("RotateAngle")

    @RotateAngle.setter
    def RotateAngle(self, value: Any) -> None:
        self._set_prop("RotateAngle", value)

    @property
    def RotateReference(self) -> Any:
        return self._get_prop("RotateReference")

    @RotateReference.setter
    def RotateReference(self, value: Any) -> None:
        self._set_prop("RotateReference", value)

    @property
    def AsianVerticalMode(self) -> bool:
        return bool(self._get_prop("AsianVerticalMode"))

    @AsianVerticalMode.setter
    def AsianVerticalMode(self, value: bool) -> None:
        self._set_prop("AsianVerticalMode", bool(value))

    @property
    def TableBorder(self) -> TableBorder:
        return cast(TableBorder, self._get_prop("TableBorder"))

    @TableBorder.setter
    def TableBorder(self, value: TableBorder) -> None:
        self._set_prop("TableBorder", value)

    @property
    def TopBorder(self) -> BorderLine:
        return cast(BorderLine, self._get_prop("TopBorder"))

    @TopBorder.setter
    def TopBorder(self, value: BorderLine) -> None:
        self._set_prop("TopBorder", value)

    @property
    def BottomBorder(self) -> BorderLine:
        return cast(BorderLine, self._get_prop("BottomBorder"))

    @BottomBorder.setter
    def BottomBorder(self, value: BorderLine) -> None:
        self._set_prop("BottomBorder", value)

    @property
    def LeftBorder(self) -> BorderLine:
        return cast(BorderLine, self._get_prop("LeftBorder"))

    @LeftBorder.setter
    def LeftBorder(self, value: BorderLine) -> None:
        self._set_prop("LeftBorder", value)

    @property
    def RightBorder(self) -> BorderLine:
        return cast(BorderLine, self._get_prop("RightBorder"))

    @RightBorder.setter
    def RightBorder(self, value: BorderLine) -> None:
        self._set_prop("RightBorder", value)

    @property
    def NumberFormat(self) -> Any:
        return self._get_prop("NumberFormat")

    @NumberFormat.setter
    def NumberFormat(self, value: Any) -> None:
        self._set_prop("NumberFormat", value)

    @property
    def ShadowFormat(self) -> Any:
        return self._get_prop("ShadowFormat")

    @ShadowFormat.setter
    def ShadowFormat(self, value: Any) -> None:
        self._set_prop("ShadowFormat", value)

    @property
    def CellProtection(self) -> Any:
        return self._get_prop("CellProtection")

    @CellProtection.setter
    def CellProtection(self, value: Any) -> None:
        self._set_prop("CellProtection", value)

    @property
    def UserDefinedAttributes(self) -> Any:
        return self._get_prop("UserDefinedAttributes")

    @UserDefinedAttributes.setter
    def UserDefinedAttributes(self, value: Any) -> None:
        self._set_prop("UserDefinedAttributes", value)

    @property
    def DiagonalTLBR(self) -> Any:
        return self._get_prop("DiagonalTLBR")

    @DiagonalTLBR.setter
    def DiagonalTLBR(self, value: Any) -> None:
        self._set_prop("DiagonalTLBR", value)

    @property
    def DiagonalBLTR(self) -> Any:
        return self._get_prop("DiagonalBLTR")

    @DiagonalBLTR.setter
    def DiagonalBLTR(self, value: Any) -> None:
        self._set_prop("DiagonalBLTR", value)

    @property
    def ShrinkToFit(self) -> bool:
        return bool(self._get_prop("ShrinkToFit"))

    @ShrinkToFit.setter
    def ShrinkToFit(self, value: bool) -> None:
        self._set_prop("ShrinkToFit", bool(value))

    @property
    def TableBorder2(self) -> TableBorder2:
        return cast(TableBorder2, self._get_prop("TableBorder2"))

    @TableBorder2.setter
    def TableBorder2(self, value: TableBorder2) -> None:
        self._set_prop("TableBorder2", value)

    @property
    def TopBorder2(self) -> BorderLine2:
        return cast(BorderLine2, self._get_prop("TopBorder2"))

    @TopBorder2.setter
    def TopBorder2(self, value: BorderLine2) -> None:
        self._set_prop("TopBorder2", value)

    @property
    def BottomBorder2(self) -> BorderLine2:
        return cast(BorderLine2, self._get_prop("BottomBorder2"))

    @BottomBorder2.setter
    def BottomBorder2(self, value: BorderLine2) -> None:
        self._set_prop("BottomBorder2", value)

    @property
    def LeftBorder2(self) -> BorderLine2:
        return cast(BorderLine2, self._get_prop("LeftBorder2"))

    @LeftBorder2.setter
    def LeftBorder2(self, value: BorderLine2) -> None:
        self._set_prop("LeftBorder2", value)

    @property
    def RightBorder2(self) -> BorderLine2:
        return cast(BorderLine2, self._get_prop("RightBorder2"))

    @RightBorder2.setter
    def RightBorder2(self, value: BorderLine2) -> None:
        self._set_prop("RightBorder2", value)

    @property
    def DiagonalTLBR2(self) -> Any:
        return self._get_prop("DiagonalTLBR2")

    @DiagonalTLBR2.setter
    def DiagonalTLBR2(self, value: Any) -> None:
        self._set_prop("DiagonalTLBR2", value)

    @property
    def DiagonalBLTR2(self) -> Any:
        return self._get_prop("DiagonalBLTR2")

    @DiagonalBLTR2.setter
    def DiagonalBLTR2(self, value: Any) -> None:
        self._set_prop("DiagonalBLTR2", value)

    @property
    def CellInteropGrabBag(self) -> Any:
        return self._get_prop("CellInteropGrabBag")

    @CellInteropGrabBag.setter
    def CellInteropGrabBag(self, value: Any) -> None:
        self._set_prop("CellInteropGrabBag", value)

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
