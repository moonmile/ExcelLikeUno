from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any
from .calc import Color, ShadowLocation
from .interfaces import StructNames

def _try_uno_struct(name: str) -> Any | None:
    try:
        import uno  # type: ignore

        return uno.createUnoStruct(name)
    except Exception:
        return None


@dataclass
class BorderLine:
    Color: int = 0
    InnerLineWidth: int = 0
    OuterLineWidth: int = 0
    LineDistance: int = 0
    def to_raw(self) -> Any:
        struct = _try_uno_struct(StructNames.BORDER_LINE)
        if struct is None:
            struct = type("BorderLine", (), {})()
        struct.Color = self.Color
        struct.InnerLineWidth = self.InnerLineWidth
        struct.OuterLineWidth = self.OuterLineWidth
        struct.LineDistance = self.LineDistance
        return struct


@dataclass
class BorderLine2(BorderLine):
    LineStyle: int = 0
    LineWidth: int = 0
    def to_raw(self) -> Any:
        struct = _try_uno_struct(StructNames.BORDER_LINE2)
        if struct is None:
            struct = type("BorderLine2", (), {})()
        struct.Color = self.Color
        struct.InnerLineWidth = self.InnerLineWidth
        struct.OuterLineWidth = self.OuterLineWidth
        struct.LineDistance = self.LineDistance
        struct.LineStyle = self.LineStyle
        struct.LineWidth = self.LineWidth
        return struct


@dataclass
class Point:
    X: int = 0
    Y: int = 0
    def to_raw(self) -> Any:
        struct = _try_uno_struct(StructNames.POINT)
        if struct is None:
            struct = type("Point", (), {})()
        struct.X = self.X
        struct.Y = self.Y
        return struct

@dataclass
class Size:
    Width: int = 0
    Height: int = 0
    def to_raw(self) -> Any:
        struct = _try_uno_struct(StructNames.SIZE)
        if struct is None:
            struct = type("Size", (), {})()
        struct.Width = self.Width
        struct.Height = self.Height
        return struct

@dataclass
class CellAddress:
    Sheet: int = 0
    Column: int = 0
    Row: int = 0
    def to_raw(self) -> Any:
        struct = _try_uno_struct(StructNames.CELL_ADDRESS)
        if struct is None:
            struct = type("CellAddress", (), {})()
        struct.Sheet = self.Sheet
        struct.Column = self.Column
        struct.Row = self.Row
        return struct


@dataclass
class CellRangeAddress:
    Sheet: int = 0
    StartColumn: int = 0
    StartRow: int = 0
    EndColumn: int = 0
    EndRow: int = 0
    def to_raw(self) -> Any:
        struct = _try_uno_struct(StructNames.CELL_RANGE_ADDRESS)
        if struct is None:
            struct = type("CellRangeAddress", (), {})()
        struct.Sheet = self.Sheet
        struct.StartColumn = self.StartColumn
        struct.StartRow = self.StartRow
        struct.EndColumn = self.EndColumn
        struct.EndRow = self.EndRow
        return struct


@dataclass
class TableSortField:
    Field: int = 0
    IsAscending: bool = True
    FieldType: int = 0
    CompareFlags: int = 0
    def to_raw(self) -> Any:
        struct = _try_uno_struct(StructNames.TABLE_SORT_FIELD)
        if struct is None:
            struct = type("TableSortField", (), {})()
        struct.Field = self.Field
        struct.IsAscending = self.IsAscending
        struct.FieldType = self.FieldType
        struct.CompareFlags = self.CompareFlags
        return struct


@dataclass
class BarCode:
    Type: int = 0
    Payload: str = ""
    ErrorCorrection: int = 0
    Border: int = 0
    def to_raw(self) -> Any:
        struct = _try_uno_struct(StructNames.BAR_CODE)
        if struct is None:
            struct = type("BarCode", (), {})()
        struct.Type = self.Type
        struct.Payload = self.Payload
        struct.ErrorCorrection = self.ErrorCorrection
        struct.Border = self.Border
        return struct


@dataclass
class BezierPoint:
    Position: Any = field(default_factory=Point)
    ControlPoint1: Any = field(default_factory=Point)
    ControlPoint2: Any = field(default_factory=Point)
    def to_raw(self) -> Any:
        struct = _try_uno_struct(StructNames.BEZIER_POINT)
        if struct is None:
            struct = type("BezierPoint", (), {})()
        struct.Position = self.Position.to_raw()
        struct.ControlPoint1 = self.ControlPoint1.to_raw()
        struct.ControlPoint2 = self.ControlPoint2.to_raw()
        return struct


@dataclass
class TableBorder:
    TopLine: BorderLine = field(default_factory=BorderLine)
    IsTopLineValid: bool = False
    BottomLine: BorderLine = field(default_factory=BorderLine)
    IsBottomLineValid: bool = False
    LeftLine: BorderLine = field(default_factory=BorderLine)
    IsLeftLineValid: bool = False
    RightLine: BorderLine = field(default_factory=BorderLine)
    IsRightLineValid: bool = False
    HorizontalLine: BorderLine = field(default_factory=BorderLine)
    IsHorizontalLineValid: bool = False
    VerticalLine: BorderLine = field(default_factory=BorderLine)
    IsVerticalLineValid: bool = False
    Distance: int = 0
    IsDistanceValid: bool = False
    def to_raw(self) -> Any:
        struct = _try_uno_struct(StructNames.TABLE_BORDER)
        if struct is None:
            struct = type("TableBorder", (), {})()
        struct.TopLine = self.TopLine.to_raw()
        struct.IsTopLineValid = self.IsTopLineValid
        struct.BottomLine = self.BottomLine.to_raw()
        struct.IsBottomLineValid = self.IsBottomLineValid
        struct.LeftLine = self.LeftLine.to_raw()
        struct.IsLeftLineValid = self.IsLeftLineValid
        struct.RightLine = self.RightLine.to_raw()
        struct.IsRightLineValid = self.IsRightLineValid
        struct.HorizontalLine = self.HorizontalLine.to_raw()
        struct.IsHorizontalLineValid = self.IsHorizontalLineValid
        struct.VerticalLine = self.VerticalLine.to_raw()
        struct.IsVerticalLineValid = self.IsVerticalLineValid
        struct.Distance = self.Distance
        struct.IsDistanceValid = self.IsDistanceValid
        return struct


@dataclass
class TableBorder2:
    TopLine: BorderLine2 = field(default_factory=BorderLine2)
    IsTopLineValid: bool = False
    BottomLine: BorderLine2 = field(default_factory=BorderLine2)
    IsBottomLineValid: bool = False
    LeftLine: BorderLine2 = field(default_factory=BorderLine2)
    IsLeftLineValid: bool = False
    RightLine: BorderLine2 = field(default_factory=BorderLine2)
    IsRightLineValid: bool = False
    HorizontalLine: BorderLine2 = field(default_factory=BorderLine2)
    IsHorizontalLineValid: bool = False
    VerticalLine: BorderLine2 = field(default_factory=BorderLine2)
    IsVerticalLineValid: bool = False
    Distance: int = 0
    IsDistanceValid: bool = False
    def to_raw(self) -> Any:
        struct = _try_uno_struct(StructNames.TABLE_BORDER2)
        if struct is None:
            struct = type("TableBorder2", (), {})()
        struct.TopLine = self.TopLine.to_raw()
        struct.IsTopLineValid = self.IsTopLineValid
        struct.BottomLine = self.BottomLine.to_raw()
        struct.IsBottomLineValid = self.IsBottomLineValid
        struct.LeftLine = self.LeftLine.to_raw()
        struct.IsLeftLineValid = self.IsLeftLineValid
        struct.RightLine = self.RightLine.to_raw()
        struct.IsRightLineValid = self.IsRightLineValid
        struct.HorizontalLine = self.HorizontalLine.to_raw()
        struct.IsHorizontalLineValid = self.IsHorizontalLineValid
        struct.VerticalLine = self.VerticalLine.to_raw()
        struct.IsVerticalLineValid = self.IsVerticalLineValid
        struct.Distance = self.Distance
        struct.IsDistanceValid = self.IsDistanceValid
        return struct


@dataclass
class ShadowFormat:
    Location: ShadowLocation = ShadowLocation.NONE
    ShadowWidth: int = 0
    ShadowColor: Color = 0
    IsTransparent: bool = False
    def to_raw(self) -> Any:
        struct = _try_uno_struct(StructNames.SHADOW_FORMAT)
        if struct is None:
            struct = type("ShadowFormat", (), {})()
        struct.Location = self.Location.value
        struct.ShadowWidth = self.ShadowWidth
        struct.ShadowColor = self.ShadowColor
        struct.IsTransparent = self.IsTransparent
        return struct


@dataclass
class CellProtection:
    IsLocked: bool = False
    IsFormulaHidden: bool = False
    IsHidden: bool = False
    IsPrintHidden: bool = False
    def to_raw(self) -> Any:
        struct = _try_uno_struct(StructNames.CELL_PROTECTION)
        if struct is None:
            struct = type("CellProtection", (), {})()
        struct.IsLocked = self.IsLocked
        struct.IsFormulaHidden = self.IsFormulaHidden
        struct.IsHidden = self.IsHidden
        struct.IsPrintHidden = self.IsPrintHidden
        return struct


__all__ = [
    "BorderLine",
    "BorderLine2",
    "Point",
    "CellAddress",
    "CellRangeAddress",
    "TableSortField",
    "BarCode",
    "BezierPoint",
    "TableBorder",
    "TableBorder2",
]
