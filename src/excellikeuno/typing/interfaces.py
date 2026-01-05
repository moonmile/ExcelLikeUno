class InterfaceNames:
    """String constants for common UNO interfaces used by Calc."""

    X_CELL = "com.sun.star.table.XCell"
    X_PROPERTY_SET = "com.sun.star.beans.XPropertySet"
    X_NAMED = "com.sun.star.container.XNamed"
    X_SPREADSHEET = "com.sun.star.sheet.XSpreadsheet"
    X_SHEET_CELL_RANGE = "com.sun.star.sheet.XSheetCellRange"
    X_SHEET_CELL_RANGES = "com.sun.star.sheet.XSheetCellRanges"
    X_CELL_RANGE_ADDRESSABLE = "com.sun.star.sheet.XCellRangeAddressable"
    X_COLUMN_ROW_RANGE = "com.sun.star.table.XColumnRowRange"
    X_SPREADSHEET_DOCUMENT = "com.sun.star.sheet.XSpreadsheetDocument"
    X_DRAW_PAGE_SUPPLIER = "com.sun.star.drawing.XDrawPageSupplier"
    X_SHAPE = "com.sun.star.drawing.XShape"
    X_SHAPES = "com.sun.star.drawing.XShapes"
    X_MERGEABLE = "com.sun.star.sheet.XMergeable"
    X_MERGEABLE_CELL = "com.sun.star.sheet.XMergeable"  # alias for backward compatibility
    X_MERGEABLE_CELL_RANGE = "com.sun.star.sheet.XMergeable"  # alias for backward compatibility
    FILL_PROPERTIES = "com.sun.star.drawing.FillProperties"
    LINE_PROPERTIES = "com.sun.star.drawing.LineProperties"
    SHADOW_PROPERTIES = "com.sun.star.drawing.ShadowProperties"
    TEXT_PROPERTIES = "com.sun.star.drawing.TextProperties"
    X_CONNECTOR_SHAPE = "com.sun.star.drawing.XConnectorShape"
    X_TEXT_DOCUMENT = "com.sun.star.text.XTextDocument"

    # Drawing shape services
    TEXT_SHAPE = "com.sun.star.drawing.TextShape"
    LINE_SHAPE = "com.sun.star.drawing.LineShape"
    RECTANGLE_SHAPE = "com.sun.star.drawing.RectangleShape"
    ELLIPSE_SHAPE = "com.sun.star.drawing.EllipseShape"
    POLYLINE_SHAPE = "com.sun.star.drawing.PolyLineShape"
    POLYPOLYGON_SHAPE = "com.sun.star.drawing.PolyPolygonShape"
    POLYPOLYGON_BEZIER_SHAPE = "com.sun.star.drawing.PolyPolygonBezierShape"
    CLOSED_BEZIER_SHAPE = "com.sun.star.drawing.ClosedBezierShape"
    CONNECTOR_SHAPE = "com.sun.star.drawing.ConnectorShape"
    CONTROL_SHAPE = "com.sun.star.drawing.ControlShape"
    CUSTOM_SHAPE = "com.sun.star.drawing.CustomShape"
    GROUP_SHAPE = "com.sun.star.drawing.GroupShape"


class StructNames:
    """String constants for UNO struct names used by Calc wrappers."""

    BORDER_LINE = "com.sun.star.table.BorderLine"
    BORDER_LINE2 = "com.sun.star.table.BorderLine2"
    POINT = "com.sun.star.awt.Point"
    SIZE = "com.sun.star.awt.Size"
    CELL_ADDRESS = "com.sun.star.table.CellAddress"
    CELL_RANGE_ADDRESS = "com.sun.star.table.CellRangeAddress"
    TABLE_SORT_FIELD = "com.sun.star.table.TableSortField"
    BAR_CODE = "com.sun.star.drawing.BarCode"
    BEZIER_POINT = "com.sun.star.drawing.BezierPoint"
    TABLE_BORDER = "com.sun.star.table.TableBorder"
    TABLE_BORDER2 = "com.sun.star.table.TableBorder2"
    SHADOW_FORMAT = "com.sun.star.drawing.ShadowFormat"
    CELL_PROTECTION = "com.sun.star.table.CellProtection"
    

__all__ = ["InterfaceNames", "StructNames"]
