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
    

__all__ = ["InterfaceNames"]
