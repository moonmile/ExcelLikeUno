from __future__ import annotations

from typing import Any

from ..typing.structs import (
    BarCodeStruct,
    BezierPointStruct,
    BorderLine2Struct,
    BorderLineStruct,
    CellAddressStruct,
    CellRangeAddressStruct,
    PointStruct,
    TableBorder2Struct,
    TableBorderStruct,
    TableSortFieldStruct,
)


def _try_uno_struct(name: str) -> Any | None:
    try:
        import uno  # type: ignore

        return uno.createUnoStruct(name)
    except Exception:
        return None


def make_border_line(
    color: int = 0,
    inner_line_width: int = 0,
    outer_line_width: int = 0,
    line_distance: int = 0,
) -> Any:
    struct = _try_uno_struct("com.sun.star.table.BorderLine")
    if struct is None:
        struct = BorderLineStruct()
    struct.Color = color
    struct.InnerLineWidth = inner_line_width
    struct.OuterLineWidth = outer_line_width
    struct.LineDistance = line_distance
    return struct


def make_border_line2(
    color: int = 0,
    inner_line_width: int = 0,
    outer_line_width: int = 0,
    line_distance: int = 0,
    line_style: int = 0,
    line_width: int = 0,
) -> Any:
    struct = _try_uno_struct("com.sun.star.table.BorderLine2")
    if struct is None:
        struct = BorderLine2Struct()
    struct.Color = color
    struct.InnerLineWidth = inner_line_width
    struct.OuterLineWidth = outer_line_width
    struct.LineDistance = line_distance
    struct.LineStyle = line_style
    struct.LineWidth = line_width
    return struct


def make_point(x: int = 0, y: int = 0) -> Any:
    struct = _try_uno_struct("com.sun.star.awt.Point")
    if struct is None:
        struct = PointStruct()
    struct.X = x
    struct.Y = y
    return struct


def _make_point_from_kwargs(kwargs: dict[str, Any] | None) -> Any:
    if kwargs is None:
        return make_point()
    return make_point(**kwargs)


def make_cell_address(sheet: int = 0, column: int = 0, row: int = 0) -> Any:
    struct = _try_uno_struct("com.sun.star.table.CellAddress")
    if struct is None:
        struct = CellAddressStruct()
    struct.Sheet = sheet
    struct.Column = column
    struct.Row = row
    return struct


def make_cell_range_address(
    sheet: int = 0,
    start_column: int = 0,
    start_row: int = 0,
    end_column: int = 0,
    end_row: int = 0,
) -> Any:
    struct = _try_uno_struct("com.sun.star.table.CellRangeAddress")
    if struct is None:
        struct = CellRangeAddressStruct()
    struct.Sheet = sheet
    struct.StartColumn = start_column
    struct.StartRow = start_row
    struct.EndColumn = end_column
    struct.EndRow = end_row
    return struct


def make_table_sort_field(
    field: int = 0,
    is_ascending: bool = True,
    field_type: int = 0,
    compare_flags: int = 0,
) -> Any:
    struct = _try_uno_struct("com.sun.star.table.TableSortField")
    if struct is None:
        struct = TableSortFieldStruct()
    struct.Field = field
    struct.IsAscending = is_ascending
    struct.FieldType = field_type
    struct.CompareFlags = compare_flags
    return struct


def make_bar_code(
    type: int = 0,
    payload: str = "",
    error_correction: int = 0,
    border: int = 0,
) -> Any:
    struct = _try_uno_struct("com.sun.star.drawing.BarCode")
    if struct is None:
        struct = BarCodeStruct()
    struct.Type = type
    struct.Payload = payload
    struct.ErrorCorrection = error_correction
    struct.Border = border
    return struct


def make_bezier_point(
    position: dict[str, Any] | None = None,
    control_point1: dict[str, Any] | None = None,
    control_point2: dict[str, Any] | None = None,
) -> Any:
    struct = _try_uno_struct("com.sun.star.drawing.BezierPoint")
    if struct is None:
        struct = BezierPointStruct()
    struct.Position = _make_point_from_kwargs(position)
    struct.ControlPoint1 = _make_point_from_kwargs(control_point1)
    struct.ControlPoint2 = _make_point_from_kwargs(control_point2)
    return struct


def _make_line_defaults(line_factory: Any, line_kwargs: dict[str, Any] | None) -> Any:
    if line_kwargs is None:
        return line_factory()
    return line_factory(**line_kwargs)


def make_table_border(
    top: dict[str, Any] | None = None,
    bottom: dict[str, Any] | None = None,
    left: dict[str, Any] | None = None,
    right: dict[str, Any] | None = None,
    horizontal: dict[str, Any] | None = None,
    vertical: dict[str, Any] | None = None,
    distance: int = 0,
    valid_flags: dict[str, bool] | None = None,
) -> Any:
    struct = _try_uno_struct("com.sun.star.table.TableBorder")
    if struct is None:
        struct = TableBorderStruct()

    struct.TopLine = _make_line_defaults(make_border_line, top)
    struct.BottomLine = _make_line_defaults(make_border_line, bottom)
    struct.LeftLine = _make_line_defaults(make_border_line, left)
    struct.RightLine = _make_line_defaults(make_border_line, right)
    struct.HorizontalLine = _make_line_defaults(make_border_line, horizontal)
    struct.VerticalLine = _make_line_defaults(make_border_line, vertical)
    struct.Distance = distance

    flags = valid_flags or {}
    struct.IsTopLineValid = flags.get("top", struct.IsTopLineValid)
    struct.IsBottomLineValid = flags.get("bottom", struct.IsBottomLineValid)
    struct.IsLeftLineValid = flags.get("left", struct.IsLeftLineValid)
    struct.IsRightLineValid = flags.get("right", struct.IsRightLineValid)
    struct.IsHorizontalLineValid = flags.get("horizontal", struct.IsHorizontalLineValid)
    struct.IsVerticalLineValid = flags.get("vertical", struct.IsVerticalLineValid)
    struct.IsDistanceValid = flags.get("distance", struct.IsDistanceValid)
    return struct


def make_table_border2(
    top: dict[str, Any] | None = None,
    bottom: dict[str, Any] | None = None,
    left: dict[str, Any] | None = None,
    right: dict[str, Any] | None = None,
    horizontal: dict[str, Any] | None = None,
    vertical: dict[str, Any] | None = None,
    distance: int = 0,
    valid_flags: dict[str, bool] | None = None,
) -> Any:
    struct = _try_uno_struct("com.sun.star.table.TableBorder2")
    if struct is None:
        struct = TableBorder2Struct()

    struct.TopLine = _make_line_defaults(make_border_line2, top)
    struct.BottomLine = _make_line_defaults(make_border_line2, bottom)
    struct.LeftLine = _make_line_defaults(make_border_line2, left)
    struct.RightLine = _make_line_defaults(make_border_line2, right)
    struct.HorizontalLine = _make_line_defaults(make_border_line2, horizontal)
    struct.VerticalLine = _make_line_defaults(make_border_line2, vertical)
    struct.Distance = distance

    flags = valid_flags or {}
    struct.IsTopLineValid = flags.get("top", struct.IsTopLineValid)
    struct.IsBottomLineValid = flags.get("bottom", struct.IsBottomLineValid)
    struct.IsLeftLineValid = flags.get("left", struct.IsLeftLineValid)
    struct.IsRightLineValid = flags.get("right", struct.IsRightLineValid)
    struct.IsHorizontalLineValid = flags.get("horizontal", struct.IsHorizontalLineValid)
    struct.IsVerticalLineValid = flags.get("vertical", struct.IsVerticalLineValid)
    struct.IsDistanceValid = flags.get("distance", struct.IsDistanceValid)
    return struct
