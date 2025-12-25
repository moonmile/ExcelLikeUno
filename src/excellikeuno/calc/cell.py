from __future__ import annotations

from typing import cast

from ..core import InterfaceNames, UnoObject
from ..typing import XCell, XPropertySet


class Cell(UnoObject):
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
    def props(self) -> XPropertySet:
        return cast(XPropertySet, self.iface(InterfaceNames.X_PROPERTY_SET))
