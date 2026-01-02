from __future__ import annotations

import re
from typing import Any, List, cast, TYPE_CHECKING

from excellikeuno.typing.structs import Point, Size

from ..core import UnoObject
from ..drawing import Shape, EllipseShape
from ..typing import InterfaceNames, XDrawPageSupplier, XNamed, XPropertySet, XSpreadsheet, XTableRows, XTableColumns
from .cell import Cell
from .range import Range, TableRow, TableColumn
from .rows import TableRows
from .columns import TableColumns

if TYPE_CHECKING:  # pragma: no cover - only for type hints
    from ..core.calc_document import CalcDocument

class Sheet(UnoObject):
    def __init__(self, sheet_obj: Any, document: "CalcDocument | None" = None) -> None:
        super().__init__(sheet_obj)
        self._document = document

    def _a1_to_pos(self, ref: str) -> tuple[int, int]:
        match = re.fullmatch(r"\$?([A-Za-z]+)\$?([1-9][0-9]*)", ref)
        if not match:
            raise ValueError(f"Invalid A1 reference: {ref}")
        col_label, row_str = match.groups()
        col = 0
        for ch in col_label.upper():
            col = col * 26 + (ord(ch) - ord("A") + 1)
        return col - 1, int(row_str) - 1

    def _normalize_range_args(
        self,
        start_column: int | str,
        start_row: int | None,
        end_column: int | str | None,
        end_row: int | None,
    ) -> tuple[int, int, int, int]:
        # Support forms: (col,row,end_col,end_row) as ints, ("A1","B3"), or ("A1:B3", None, None, None)
        if isinstance(start_column, str) and start_row is None and end_column is None and end_row is None:
            if ":" not in start_column:
                raise ValueError("Single A1 reference must include ':' for range, or pass end separately")
            left, right = start_column.split(":", 1)
            sc, sr = self._a1_to_pos(left)
            ec, er = self._a1_to_pos(right)
            return sc, sr, ec, er

        if isinstance(start_column, str):
            if start_row is not None:
                raise ValueError("When start is A1 notation, omit start_row")
            sc, sr = self._a1_to_pos(start_column)
        else:
            if start_row is None:
                raise ValueError("Row is required when start column is numeric")
            sc, sr = int(start_column), int(start_row)

        if isinstance(end_column, str):
            if end_row is not None:
                raise ValueError("When end is A1 notation, omit end_row")
            ec, er = self._a1_to_pos(end_column)
        else:
            if end_column is None or end_row is None:
                raise ValueError("End column/row are required for numeric range")
            ec, er = int(end_column), int(end_row)

        return sc, sr, ec, er

    def cell(self, column: int | str, row: int | None = None) -> Cell:
        sheet = cast(XSpreadsheet, self.iface(InterfaceNames.X_SPREADSHEET))
        if isinstance(column, str):
            if row is not None:
                raise ValueError("When using A1 notation, do not pass row separately")
            column, row = self._a1_to_pos(column)
        if row is None:
            raise ValueError("Row is required when column is numeric")
        return Cell(sheet.getCellByPosition(int(column), int(row)))
    
    def range(
        self,
        start_column: int | str,
        start_row: int | None = None,
        end_column: int | str | None = None,
        end_row: int | None = None,
    ) -> Range:
        sc, sr, ec, er = self._normalize_range_args(start_column, start_row, end_column, end_row)
        sheet = cast(XSpreadsheet, self.iface(InterfaceNames.X_SPREADSHEET))
        return Range(sheet.getCellRangeByPosition(sc, sr, ec, er))

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
    def shapes(self) -> "Shapes":
        return Shapes(self)

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

    @property
    def rows(self) -> XTableRows:
        sheet = cast(XSpreadsheet, self.iface(InterfaceNames.X_SPREADSHEET))
        return TableRows(sheet.getRows())

    def row(self, index: int) -> TableRow:
        return self.rows.getByIndex(index)

    @property
    def columns(self) -> XTableColumns:
        sheet = cast(XSpreadsheet, self.iface(InterfaceNames.X_SPREADSHEET))
        return TableColumns(sheet.getColumns())

    def column(self, index: int) -> TableColumn:
        return self.columns.getByIndex(index)

    @property
    def document(self) -> "CalcDocument":
        if self._document is not None:
            return self._document
        raise AttributeError("Sheet has no associated document; construct Sheet with document reference")


class Shapes:
    """Helper for creating and managing shapes on a sheet's draw page."""

    def __init__(self, sheet: Sheet) -> None:
        self.sheet = sheet

    def __call__(self) -> List[Shape]:
        """Return all shapes on the sheet as wrapped Shape objects."""
        draw_page = self.sheet._draw_page()
        return [Shape(draw_page.getByIndex(i)) for i in range(draw_page.getCount())]

    def add_ellipse_shape(
        self,
        x: int,
        y: int,
        width: int,
        height: int,
        fill_color: int | None = None,
        line_color: int | None = None,
    ) -> EllipseShape:

        draw_page = self.sheet._draw_page()
        doc = self.sheet.document
        ellipse_raw = doc.createInstance("com.sun.star.drawing.EllipseShape")
        ellipse = EllipseShape(ellipse_raw)

        # Position and size (1/100 mm)
        point_struct = Point(x, y)
        size_struct = Size(width, height)
        point = point_struct.to_raw() if hasattr(point_struct, "to_raw") else point_struct
        size = size_struct.to_raw() if hasattr(size_struct, "to_raw") else size_struct
        ellipse.Position = point
        ellipse.Size = size
        if fill_color is not None:
            ellipse.FillColor = int(fill_color)
        if line_color is not None:
            ellipse.LineColor = int(line_color)

        # Must add to draw page before some properties become available
        draw_page.add(ellipse_raw)

        return ellipse
    