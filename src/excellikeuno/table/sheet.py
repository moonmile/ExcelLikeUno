from __future__ import annotations

import re
from typing import Any, List, cast, TYPE_CHECKING

from excellikeuno.drawing.closed_bezier_shape import ClosedBezierShape
from excellikeuno.drawing.connector_shape import ConnectorShape
from excellikeuno.drawing.control_shape import ControlShape
from excellikeuno.drawing.custom_shape import CustomShape
from excellikeuno.drawing.group_shape import GroupShape
from excellikeuno.drawing.line_shape import LineShape
from excellikeuno.drawing.polyline_shape import PolyLineShape
from excellikeuno.drawing.polypolygon_bezier_shape import PolyPolygonBezierShape
from excellikeuno.drawing.polypolygon_shape import PolyPolygonShape
from excellikeuno.drawing.rectangle_shape import RectangleShape
from excellikeuno.drawing.text_shape import TextShape
from excellikeuno.typing.calc import Color, ConnectionType, XConnectorShape
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

        # Support sheet.range("A1", "B2")
        if isinstance(start_column, str) and isinstance(start_row, str) and end_column is None and end_row is None:
            sc, sr = self._a1_to_pos(start_column)
            ec, er = self._a1_to_pos(start_row)
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

    def _wrap_shape(self, raw: Any) -> Shape:
        supports = getattr(raw, "supportsService", None)
        if callable(supports):
            try:
                if supports(InterfaceNames.TEXT_SHAPE):
                    return TextShape(raw)
                if supports(InterfaceNames.LINE_SHAPE):
                    return LineShape(raw)
                if supports(InterfaceNames.RECTANGLE_SHAPE):
                    return RectangleShape(raw)
                if supports(InterfaceNames.ELLIPSE_SHAPE):
                    return EllipseShape(raw)
                if supports(InterfaceNames.POLYLINE_SHAPE):
                    return PolyLineShape(raw)
                if supports(InterfaceNames.POLYPOLYGON_SHAPE):
                    return PolyPolygonShape(raw)
                if supports(InterfaceNames.POLYPOLYGON_BEZIER_SHAPE):
                    return PolyPolygonBezierShape(raw)
                if supports(InterfaceNames.CLOSED_BEZIER_SHAPE):
                    return ClosedBezierShape(raw)
                if supports(InterfaceNames.CONNECTOR_SHAPE):
                    return ConnectorShape(raw)
                if supports(InterfaceNames.CONTROL_SHAPE):
                    return ControlShape(raw)
                if supports(InterfaceNames.CUSTOM_SHAPE):
                    return CustomShape(raw)  
                if supports(InterfaceNames.GROUP_SHAPE):
                    return GroupShape(raw)
                
            except Exception:
                pass
        return Shape(raw)

    def __call__(self) -> List[Shape]:
        """Return all shapes on the sheet as wrapped Shape objects."""
        draw_page = self.sheet._draw_page()
        return [self._wrap_shape(draw_page.getByIndex(i)) for i in range(draw_page.getCount())]

    def __len__(self) -> int:
        draw_page = self.sheet._draw_page()
        return draw_page.getCount()

    def __iter__(self):
        return iter(self())

    def __getitem__(self, index: int) -> Shape:
        draw_page = self.sheet._draw_page()
        count = draw_page.getCount()
        if index < 0:
            index += count
        if index < 0 or index >= count:
            raise IndexError("shape index out of range")
        return self._wrap_shape(draw_page.getByIndex(index))

    def remove(self, shape: Any) -> None:
        """Remove a shape from the sheet draw page.

        Accepts a Shape wrapper, a raw UNO shape, or an index.
        """
        draw_page = self.sheet._draw_page()

        target = shape
        if isinstance(shape, int):
            count = draw_page.getCount()
            idx = shape if shape >= 0 else count + shape
            if idx < 0 or idx >= count:
                raise IndexError("shape index out of range")
            target = draw_page.getByIndex(idx)
        elif hasattr(shape, "raw"):
            target = getattr(shape, "raw")

        try:
            draw_page.remove(target)
        except Exception as exc:
            raise RuntimeError(f"failed to remove shape: {exc}")

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
        ellipse_raw = doc.createInstance(InterfaceNames.ELLIPSE_SHAPE)
        ellipse = EllipseShape(ellipse_raw)

        # Position and size (1/100 mm)
        ellipse.Position = Point(x, y)
        ellipse.Size = Size(width, height)
        if fill_color is not None:
            ellipse.FillColor = int(fill_color)
        if line_color is not None:
            ellipse.LineColor = int(line_color)

        # Must add to draw page before some properties become available
        draw_page.add(ellipse_raw)
        return ellipse
    
    def add_circle_shape(
        self,
        x: int,
        y: int,
        diameter: int,
        fill_color: int | None = None,
        line_color: int | None = None,
    ) -> EllipseShape:
        """Add a circle shape to the sheet's draw page.

        Args:
            x: The X position (1/100 mm).
            y: The Y position (1/100 mm).
            diameter: The diameter (1/100 mm).
            fill_color: Optional fill color as integer RGB.
            line_color: Optional line color as integer RGB.

        Returns:
            The created EllipseShape representing the circle.
        """
        return self.add_ellipse_shape(
            x=x,
            y=y,
            width=diameter,
            height=diameter,
            fill_color=fill_color,
            line_color=line_color,
        )
    
    def add_line_shape(
        self,
        x1: int,
        y1: int,
        x2: int,
        y2: int,
        line_color: int | None = None,
        line_width: int | None = None,
    ) -> LineShape:
        """Add a line shape to the sheet's draw page.

        Args:
            x1: The starting X position (1/100 mm).
            y1: The starting Y position (1/100 mm).
            x2: The ending X position (1/100 mm).
            y2: The ending Y position (1/100 mm).
            line_color: Optional line color as integer RGB.
            line_width: Optional line width (1/100 mm).

        Returns:
            The created Shape representing the line.
        """
        draw_page = self.sheet._draw_page()
        doc = self.sheet.document
        line_raw = doc.createInstance(InterfaceNames.LINE_SHAPE)
        line = LineShape(line_raw)

        # Position and size (1/100 mm)
        line.Position = Point(min(x1, x2), min(y1, y2))
        line.Size = Size(abs(x2 - x1), abs(y2 - y1))
        if line_color is not None:
            line.LineColor = int(line_color)
        if line_width is not None:
            line.LineWidth = int(line_width)

        # Must add to draw page before some properties become available
        draw_page.add(line_raw)
        return line
    
    def add_rectangle_shape(
        self,
        x: int,
        y: int,
        width: int,
        height: int,
        fill_color: int | None = None,
        line_color: int | None = None,
    ) -> RectangleShape:
        """Add a rectangle shape to the sheet's draw page.
        
        Args:
            x: The X position (1/100 mm).
            y: The Y position (1/100 mm).
            width: The width (1/100 mm).
            height: The height (1/100 mm).
            fill_color: Optional fill color as integer RGB.
            line_color: Optional line color as integer RGB.
        Returns:
            The created Shape representing the rectangle.
        """

        draw_page = self.sheet._draw_page()
        doc = self.sheet.document
        rect_raw = doc.createInstance(InterfaceNames.RECTANGLE_SHAPE)
        rect = RectangleShape(rect_raw)

        # Position and size (1/100 mm)
        rect.Position = Point(x, y)
        rect.Size = Size(width, height)
        if fill_color is not None:
            rect.FillColor = int(fill_color)
        if line_color is not None:
            rect.LineColor = int(line_color)

        # Must add to draw page before some properties become available
        draw_page.add(rect_raw)
        return rect
    
    def add_polyline_shape(
        self,
        points: List[Point],
        line_color: int | None = None,
        line_width: int | None = None,
    ) -> PolyLineShape:
        """Add a polyline shape to the sheet's draw page.

        Args:
            points: A list of Point objects defining the polyline vertices.
            line_color: Optional line color as integer RGB.
            line_width: Optional line width (1/100 mm).

        Returns:
            The created Shape representing the polyline.
        """
        draw_page = self.sheet._draw_page()
        doc = self.sheet.document
        polyline_raw = doc.createInstance(InterfaceNames.POLYLINE_SHAPE)
        polyline = PolyLineShape(polyline_raw)

        # Set the points
        raw_points = [p.to_raw() for p in points]
        polyline_raw.setPoints(tuple(raw_points))

        if line_color is not None:
            polyline.LineColor = int(line_color)
        if line_width is not None:
            polyline.LineWidth = int(line_width)

        # Must add to draw page before some properties become available
        draw_page.add(polyline_raw)
        return polyline
    
    def add_polypolygon_shape(
        self,
        polygons: List[List[Point]],
        line_color: int | None = None,
        line_width: int | None = None,
        fill_color: int | None = None,
    ) -> PolyPolygonShape:
        """Add a polypolygon shape to the sheet's draw page.

        Args:
            polygons: A list of polygons, each defined as a list of Point objects.
            line_color: Optional line color as integer RGB.
            line_width: Optional line width (1/100 mm).
            fill_color: Optional fill color as integer RGB.

        Returns:
            The created Shape representing the polypolygon.
        """
        draw_page = self.sheet._draw_page()
        doc = self.sheet.document
        polypolygon_raw = doc.createInstance(InterfaceNames.POLYPOLYGON_SHAPE)
        polypolygon = PolyPolygonShape(polypolygon_raw)

        # Set the polygons
        raw_polygons = [tuple(p.to_raw() for p in polygon) for polygon in polygons]
        polypolygon_raw.setPolygons(tuple(raw_polygons))

        if line_color is not None:
            polypolygon.LineColor = int(line_color)
        if line_width is not None:
            polypolygon.LineWidth = int(line_width)
        if fill_color is not None:
            polypolygon.FillColor = int(fill_color)

        # Must add to draw page before some properties become available
        draw_page.add(polypolygon_raw)
        return polypolygon
    
    def add_ploypolygon_bezier_shape(
        self,
        polygons: List[List[Point]],
        line_color: int | None = None,
        line_width: int | None = None,
        fill_color: int | None = None,
    ) -> PolyPolygonBezierShape:
        """Add a polypolygon bezier shape to the sheet's draw page.

        Args:
            polygons: A list of polygons, each defined as a list of Point objects.
            line_color: Optional line color as integer RGB.
            line_width: Optional line width (1/100 mm).
            fill_color: Optional fill color as integer RGB.

        Returns:
            The created Shape representing the polypolygon bezier.
        """
        draw_page = self.sheet._draw_page()
        doc = self.sheet.document
        polypolygon_bezier_raw = doc.createInstance(InterfaceNames.POLYPOLYGON_BEZIER_SHAPE)
        polypolygon_bezier = PolyPolygonBezierShape(polypolygon_bezier_raw)

        # Set the polygons
        raw_polygons = [tuple(p.to_raw() for p in polygon) for polygon in polygons]
        polypolygon_bezier_raw.setPolygons(tuple(raw_polygons))

        if line_color is not None:
            polypolygon_bezier.LineColor = int(line_color)
        if line_width is not None:
            polypolygon_bezier.LineWidth = int(line_width)
        if fill_color is not None:
            polypolygon_bezier.FillColor = int(fill_color)

        # Must add to draw page before some properties become available
        draw_page.add(polypolygon_bezier_raw)
        return polypolygon_bezier

    def add_text_shape(
        self,
        x: int,
        y: int,
        width: int,
        height: int,
        text: str = "",
        fill_color: int | None = None,
        line_color: int | None = None,
    ) -> TextShape:
        """Add a text shape to the sheet's draw page.

        Args:
            x: The X position (1/100 mm).
            y: The Y position (1/100 mm).
            width: The width (1/100 mm).
            height: The height (1/100 mm).
            text: The text content of the shape.
            fill_color: Optional fill color as integer RGB.
            line_color: Optional line color as integer RGB.

        Returns:
            The created Shape representing the text shape.
        """
        draw_page = self.sheet._draw_page()
        doc = self.sheet.document
        textshape_raw = doc.createInstance(InterfaceNames.TEXT_SHAPE)
        textshape = TextShape(textshape_raw)

        # Position and size (1/100 mm)
        textshape.Position = Point(x, y)
        textshape.Size = Size(width, height)
        # Must add to draw page before some properties become available
        draw_page.add(textshape_raw)
        textshape.String = text
        if fill_color is not None:
            textshape.FillColor = Color(fill_color)
        if line_color is not None:
            textshape.LineColor = Color(line_color)
        return textshape
    
    def add_closed_bezier_shape(
        self,
        points: List[Point],
        line_color: int | None = None,
        line_width: int | None = None,
        fill_color: int | None = None,
    ) -> ClosedBezierShape:
        """Add a closed bezier shape to the sheet's draw page.

        Args:
            points: A list of Point objects defining the bezier vertices.
            line_color: Optional line color as integer RGB.
            line_width: Optional line width (1/100 mm).
            fill_color: Optional fill color as integer RGB.

        Returns:
            The created Shape representing the closed bezier.
        """
        draw_page = self.sheet._draw_page()
        doc = self.sheet.document
        closed_bezier_raw = doc.createInstance(InterfaceNames.CLOSED_BEZIER_SHAPE)
        closed_bezier = ClosedBezierShape(closed_bezier_raw)

        # Set the points
        raw_points = [p.to_raw() for p in points]
        closed_bezier_raw.setControlPoints(tuple(raw_points))

        if line_color is not None:
            closed_bezier.LineColor = int(line_color)
        if line_width is not None:
            closed_bezier.LineWidth = int(line_width)
        if fill_color is not None:
            closed_bezier.FillColor = int(fill_color)

        # Must add to draw page before some properties become available
        draw_page.add(closed_bezier_raw)
        return closed_bezier
    
    def add_control_shape(
        self,
        x: int,
        y: int,
        width: int,
        height: int,
        control_type: str,
        fill_color: int | None = None,
        line_color: int | None = None,
    ) -> ControlShape:
        """Add a control shape to the sheet's draw page.

        Args:
            x: The X position (1/100 mm).
            y: The Y position (1/100 mm).
            width: The width (1/100 mm).
            height: The height (1/100 mm).
            control_type: The control type identifier.
            fill_color: Optional fill color as integer RGB.
            line_color: Optional line color as integer RGB.

        Returns:
            The created Shape representing the control shape.
        """
        draw_page = self.sheet._draw_page()
        doc = self.sheet.document
        control_shape_raw = doc.createInstance(InterfaceNames.CONTROL_SHAPE)
        control_shape = ControlShape(control_shape_raw)

        # Position and size (1/100 mm)
        control_shape.Position = Point(x, y)
        control_shape.Size = Size(width, height)
        control_shape.ControlType = control_type
        if fill_color is not None:
            control_shape.FillColor = int(fill_color)
        if line_color is not None:
            control_shape.LineColor = int(line_color)

        # Must add to draw page before some properties become available
        draw_page.add(control_shape_raw)
        return control_shape    
    
    def add_custom_shape(
        self,
        x: int,
        y: int,
        width: int,
        height: int,
        shape_data: bytes,
        fill_color: int | None = None,
        line_color: int | None = None,
    ) -> CustomShape:
        """Add a custom shape to the sheet's draw page.

        Args:
            x: The X position (1/100 mm).
            y: The Y position (1/100 mm).
            width: The width (1/100 mm).
            height: The height (1/100 mm).
            shape_data: The custom shape data as bytes.
            fill_color: Optional fill color as integer RGB.
            line_color: Optional line color as integer RGB.

        Returns:
            The created Shape representing the custom shape.
        """
        draw_page = self.sheet._draw_page()
        doc = self.sheet.document
        custom_shape_raw = doc.createInstance(InterfaceNames.CUSTOM_SHAPE)
        custom_shape = CustomShape(custom_shape_raw)

        # Position and size (1/100 mm)
        custom_shape.Position = Point(x, y)
        custom_shape.Size = Size(width, height)
        custom_shape.ShapeData = shape_data
        if fill_color is not None:
            custom_shape.FillColor = int(fill_color)
        if line_color is not None:
            custom_shape.LineColor = int(line_color)

        # Must add to draw page before some properties become available
        draw_page.add(custom_shape_raw)
        return custom_shape
    
    def add_group_shape(
        self,
        shapes: List[Shape],
    ) -> GroupShape:
        """Add a group shape to the sheet's draw page.

        Args:
            shapes: A list of Shape objects to group.

        Returns:
            The created Shape representing the group shape.
        """
        draw_page = self.sheet._draw_page()
        doc = self.sheet.document
        group_shape_raw = doc.createInstance(InterfaceNames.GROUP_SHAPE)
        group_shape = GroupShape(group_shape_raw)

        # Add shapes to group
        group_iface = cast(XDrawPageSupplier, group_shape.iface(InterfaceNames.X_DRAW_PAGE_SUPPLIER))
        group_draw_page = group_iface.getDrawPage()
        for shape in shapes:
            group_draw_page.add(shape.iface(InterfaceNames.X_SHAPE))

        # Must add to draw page before some properties become available
        draw_page.add(group_shape_raw)
        return group_shape
    
    def add_connector_shape(
        self,
        start_shape: Shape,
        end_shape: Shape,
        line_color: int | None = None,
        line_width: int | None = None,
    ) -> ConnectorShape:
        """Add a connector shape between two shapes on the sheet's draw page.

        Args:
            start_shape: The starting Shape object.
            end_shape: The ending Shape object.
            line_color: Optional line color as integer RGB.
            line_width: Optional line width (1/100 mm).

        Returns:
            The created Shape representing the connector shape.
        """
        draw_page = self.sheet._draw_page()
        doc = self.sheet.document
        connector_shape_raw = doc.createInstance(InterfaceNames.CONNECTOR_SHAPE)
        connector_shape = ConnectorShape(connector_shape_raw)

        # Set start and end shapes
        connector_iface = cast("XConnectorShape", connector_shape.iface(InterfaceNames.X_CONNECTOR_SHAPE))
        connector_iface.connectStart(start_shape.iface(InterfaceNames.X_SHAPE), ConnectionType.AUTO)
        connector_iface.connectEnd(end_shape.iface(InterfaceNames.X_SHAPE), ConnectionType.AUTO)

        if line_color is not None:
            connector_shape.LineColor = int(line_color)
        if line_width is not None:
            connector_shape.LineWidth = int(line_width)

        # Must add to draw page before some properties become available
        draw_page.add(connector_shape_raw)
        return connector_shape