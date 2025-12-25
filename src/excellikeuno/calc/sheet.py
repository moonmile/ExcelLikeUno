from __future__ import annotations

from typing import Any, List, cast

from ..core import InterfaceNames, UnoObject
from ..typing import XDrawPageSupplier, XNamed, XPropertySet, XSpreadsheet
from .cell import Cell
from .shape import Shape


class Sheet(UnoObject):
    def cell(self, column: int, row: int) -> Cell:
        sheet = cast(XSpreadsheet, self.iface(InterfaceNames.X_SPREADSHEET))
        return Cell(sheet.getCellByPosition(column, row))

    def _draw_page(self):
        supplier = cast(XDrawPageSupplier, self.iface(InterfaceNames.X_DRAW_PAGE_SUPPLIER))
        return supplier.getDrawPage()

    def shape(self, index: int) -> Shape:
        draw_page = self._draw_page()
        return Shape(draw_page.getByIndex(index))

    def shapes(self) -> List[Shape]:
        draw_page = self._draw_page()
        return [Shape(draw_page.getByIndex(i)) for i in range(draw_page.getCount())]
    
    @property
    def name(self) -> str:
        named = cast(XNamed, self.iface(InterfaceNames.X_NAMED))
        return named.getName()

    @name.setter
    def name(self, value: str) -> None:
        named = cast(XNamed, self.iface(InterfaceNames.X_NAMED))
        named.setName(value)

    @property
    def is_visible(self) -> bool:
        props = cast(XPropertySet, self.iface(InterfaceNames.X_PROPERTY_SET))
        return bool(props.getPropertyValue("IsVisible"))

    @is_visible.setter
    def is_visible(self, visible: bool) -> None:
        props = cast(XPropertySet, self.iface(InterfaceNames.X_PROPERTY_SET))
        props.setPropertyValue("IsVisible", bool(visible))

    @property
    def page_style(self) -> str:
        props = cast(XPropertySet, self.iface(InterfaceNames.X_PROPERTY_SET))
        return cast(str, props.getPropertyValue("PageStyle"))

    @page_style.setter
    def page_style(self, style: str) -> None:
        props = cast(XPropertySet, self.iface(InterfaceNames.X_PROPERTY_SET))
        props.setPropertyValue("PageStyle", style)

    @property
    def tab_color(self) -> Any:
        props = cast(XPropertySet, self.iface(InterfaceNames.X_PROPERTY_SET))
        return props.getPropertyValue("TabColor")

    @tab_color.setter
    def tab_color(self, color: Any) -> None:
        props = cast(XPropertySet, self.iface(InterfaceNames.X_PROPERTY_SET))
        props.setPropertyValue("TabColor", color)

    @property
    def table_layout(self) -> int:
        props = cast(XPropertySet, self.iface(InterfaceNames.X_PROPERTY_SET))
        return int(props.getPropertyValue("TableLayout"))

    @table_layout.setter
    def table_layout(self, layout: int) -> None:
        props = cast(XPropertySet, self.iface(InterfaceNames.X_PROPERTY_SET))
        props.setPropertyValue("TableLayout", int(layout))

    @property
    def automatic_print_area(self) -> bool:
        props = cast(XPropertySet, self.iface(InterfaceNames.X_PROPERTY_SET))
        return bool(props.getPropertyValue("AutomaticPrintArea"))

    @automatic_print_area.setter
    def automatic_print_area(self, enabled: bool) -> None:
        props = cast(XPropertySet, self.iface(InterfaceNames.X_PROPERTY_SET))
        props.setPropertyValue("AutomaticPrintArea", bool(enabled))

    @property
    def conditional_formats(self) -> Any:
        props = cast(XPropertySet, self.iface(InterfaceNames.X_PROPERTY_SET))
        return props.getPropertyValue("ConditionalFormats")

    @conditional_formats.setter
    def conditional_formats(self, value: Any) -> None:
        props = cast(XPropertySet, self.iface(InterfaceNames.X_PROPERTY_SET))
        props.setPropertyValue("ConditionalFormats", value)
