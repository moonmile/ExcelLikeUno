from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any


@dataclass
class BorderLineStruct:
    Color: int = 0
    InnerLineWidth: int = 0
    OuterLineWidth: int = 0
    LineDistance: int = 0


@dataclass
class BorderLine2Struct(BorderLineStruct):
    LineStyle: int = 0
    LineWidth: int = 0


@dataclass
class PointStruct:
    X: int = 0
    Y: int = 0


@dataclass
class CellAddressStruct:
    Sheet: int = 0
    Column: int = 0
    Row: int = 0


@dataclass
class CellRangeAddressStruct:
    Sheet: int = 0
    StartColumn: int = 0
    StartRow: int = 0
    EndColumn: int = 0
    EndRow: int = 0


@dataclass
class TableSortFieldStruct:
    Field: int = 0
    IsAscending: bool = True
    FieldType: int = 0
    CompareFlags: int = 0


@dataclass
class BarCodeStruct:
    Type: int = 0
    Payload: str = ""
    ErrorCorrection: int = 0
    Border: int = 0


@dataclass
class BezierPointStruct:
    Position: Any = field(default_factory=PointStruct)
    ControlPoint1: Any = field(default_factory=PointStruct)
    ControlPoint2: Any = field(default_factory=PointStruct)


@dataclass
class TableBorderStruct:
    TopLine: BorderLineStruct = field(default_factory=BorderLineStruct)
    IsTopLineValid: bool = False
    BottomLine: BorderLineStruct = field(default_factory=BorderLineStruct)
    IsBottomLineValid: bool = False
    LeftLine: BorderLineStruct = field(default_factory=BorderLineStruct)
    IsLeftLineValid: bool = False
    RightLine: BorderLineStruct = field(default_factory=BorderLineStruct)
    IsRightLineValid: bool = False
    HorizontalLine: BorderLineStruct = field(default_factory=BorderLineStruct)
    IsHorizontalLineValid: bool = False
    VerticalLine: BorderLineStruct = field(default_factory=BorderLineStruct)
    IsVerticalLineValid: bool = False
    Distance: int = 0
    IsDistanceValid: bool = False


@dataclass
class TableBorder2Struct:
    TopLine: BorderLine2Struct = field(default_factory=BorderLine2Struct)
    IsTopLineValid: bool = False
    BottomLine: BorderLine2Struct = field(default_factory=BorderLine2Struct)
    IsBottomLineValid: bool = False
    LeftLine: BorderLine2Struct = field(default_factory=BorderLine2Struct)
    IsLeftLineValid: bool = False
    RightLine: BorderLine2Struct = field(default_factory=BorderLine2Struct)
    IsRightLineValid: bool = False
    HorizontalLine: BorderLine2Struct = field(default_factory=BorderLine2Struct)
    IsHorizontalLineValid: bool = False
    VerticalLine: BorderLine2Struct = field(default_factory=BorderLine2Struct)
    IsVerticalLineValid: bool = False
    Distance: int = 0
    IsDistanceValid: bool = False


__all__ = [
    "BorderLineStruct",
    "BorderLine2Struct",
    "PointStruct",
    "CellAddressStruct",
    "CellRangeAddressStruct",
    "TableSortFieldStruct",
    "BarCodeStruct",
    "BezierPointStruct",
    "TableBorderStruct",
    "TableBorder2Struct",
]
