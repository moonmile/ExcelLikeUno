from __future__ import annotations

from typing import Any, TYPE_CHECKING

from ..core import UnoObject
from ..typing import InterfaceNames
from ..style.font import Font
from ..style.border import Borders
from .cell import Cell

if TYPE_CHECKING:  # pragma: no cover - type hints only
    from com.sun.star.beans import XPropertySet
    from com.sun.star.table import XCell


class RawProps:
    """XPropertySet への透過ラッパー（CellProperties を使わない軽量版）。"""

    def __init__(self, owner: UnoObject) -> None:
        object.__setattr__(self, "_owner", owner)

    def _props(self) -> "XPropertySet":
        return self._owner.iface(InterfaceNames.X_PROPERTY_SET)  # type: ignore[return-value]

    def get_property(self, name: str) -> Any:
        return self._props().getPropertyValue(name)

    def set_property(self, name: str, value: Any) -> None:
        self._props().setPropertyValue(name, value)

    def __getattr__(self, name: str) -> Any:
        try:
            return self.get_property(name)
        except Exception as exc:  # pragma: no cover - UNO errors bubble up
            raise AttributeError(f"Unknown property: {name}") from exc

    def __setattr__(self, name: str, value: Any) -> None:
        if name.startswith("_"):
            object.__setattr__(self, name, value)
            return
        try:
            self.set_property(name, value)
        except Exception as exc:  # pragma: no cover - UNO errors bubble up
            raise AttributeError(f"Cannot set property: {name}") from exc


class SheetCell(Cell):
    """VBAライクな Cell ラッパー。CellProperties ではなく RawProps を利用する。"""

    @property
    def raw(self) -> "XCell":
        return super().raw  # type: ignore[return-value]

    @property
    def props(self) -> RawProps:
        existing = self.__dict__.get("_raw_props")
        if existing is None:
            existing = RawProps(self)
            object.__setattr__(self, "_raw_props", existing)
        return existing  # type: ignore[return-value]

    @property
    def borders(self) -> Borders:
        existing = self.__dict__.get("_borders")
        if existing is None:
            existing = Borders(owner=self)
            object.__setattr__(self, "_borders", existing)
        return existing

    @borders.setter
    def borders(self, value: Borders) -> None:
        current_proxy = self.__dict__.get("_borders") or Borders(owner=self)
        object.__setattr__(self, "_borders", current_proxy)
        try:
            current = value._current()  # type: ignore[attr-defined]
        except Exception:
            try:
                current = dict(value)  # type: ignore[arg-type]
            except Exception:
                current = {}
        if not current:
            return
        current_proxy.apply(**current)

    @property
    def font(self) -> Font:
        return Font(owner=self)

    @font.setter
    def font(self, value: Font) -> None:
        try:
            current = value._current()  # type: ignore[attr-defined]
        except Exception:
            try:
                current = dict(value)  # type: ignore[arg-type]
            except Exception:
                current = {}
        if not current:
            return
        Font(owner=self).apply(**current)

    @property
    def value(self) -> float:
        cell_iface = self.iface(InterfaceNames.X_CELL)
        return cell_iface.getValue() if cell_iface is not None else 0.0

    @value.setter
    def value(self, value: Any) -> None:
        cell_iface = self.iface(InterfaceNames.X_CELL)
        if cell_iface is None:
            return
        if isinstance(value, (int, float)):
            cell_iface.setValue(float(value))
        else:
            cell_iface.setFormula(str(value))

    @property
    def formula(self) -> str:
        cell_iface = self.iface(InterfaceNames.X_CELL)
        return cell_iface.getFormula() if cell_iface is not None else ""

    @formula.setter
    def formula(self, formula: str) -> None:
        cell_iface = self.iface(InterfaceNames.X_CELL)
        if cell_iface is not None:
            cell_iface.setFormula(formula)

    @property
    def text(self) -> str:
        return self.formula

    @text.setter
    def text(self, text: str) -> None:
        cell_iface = self.iface(InterfaceNames.X_CELL)
        if cell_iface is not None:
            cell_iface.setFormula(text)
