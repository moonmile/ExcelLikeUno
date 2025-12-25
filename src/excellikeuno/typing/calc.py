from __future__ import annotations

from enum import IntEnum
from typing import Any, Protocol, runtime_checkable


@runtime_checkable
class XCell(Protocol):
    def getValue(self) -> float:
        ...

    def setValue(self, value: float) -> None:
        ...

    def getFormula(self) -> str:
        ...

    def setFormula(self, formula: str) -> None:
        ...


@runtime_checkable
class XPropertySet(Protocol):
    def getPropertyValue(self, name: str) -> Any:
        ...

    def setPropertyValue(self, name: str, value: Any) -> None:
        ...


@runtime_checkable
class XSpreadsheet(Protocol):
    def getCellByPosition(self, column: int, row: int) -> Any:
        ...


@runtime_checkable
class XDrawPage(Protocol):
    def getByIndex(self, index: int) -> Any:
        ...

    def getCount(self) -> int:
        ...


@runtime_checkable
class XDrawPageSupplier(Protocol):
    def getDrawPage(self) -> XDrawPage:
        ...


@runtime_checkable
class XShape(Protocol):
    def getPosition(self) -> Any:
        ...

    def setPosition(self, position: Any) -> None:
        ...

    def getSize(self) -> Any:
        ...

    def setSize(self, size: Any) -> None:
        ...


@runtime_checkable
class XNamed(Protocol):
    def getName(self) -> str:
        ...

    def setName(self, name: str) -> None:
        ...


@runtime_checkable
class XSpreadsheetDocument(Protocol):
    def getSheets(self) -> Any:
        ...


@runtime_checkable
class BorderLine(Protocol):
    Color: int
    InnerLineWidth: int
    OuterLineWidth: int
    LineDistance: int


@runtime_checkable
class TableBorder(Protocol):
    TopLine: BorderLine
    IsTopLineValid: bool
    BottomLine: BorderLine
    IsBottomLineValid: bool
    LeftLine: BorderLine
    IsLeftLineValid: bool
    RightLine: BorderLine
    IsRightLineValid: bool
    HorizontalLine: BorderLine
    IsHorizontalLineValid: bool
    VerticalLine: BorderLine
    IsVerticalLineValid: bool
    Distance: int
    IsDistanceValid: bool


@runtime_checkable
class BorderLine2(Protocol):
    Color: int
    InnerLineWidth: int
    OuterLineWidth: int
    LineDistance: int
    LineStyle: int
    LineWidth: int


@runtime_checkable
class TableBorder2(Protocol):
    TopLine: BorderLine2
    IsTopLineValid: bool
    BottomLine: BorderLine2
    IsBottomLineValid: bool
    LeftLine: BorderLine2
    IsLeftLineValid: bool
    RightLine: BorderLine2
    IsRightLineValid: bool
    HorizontalLine: BorderLine2
    IsHorizontalLineValid: bool
    VerticalLine: BorderLine2
    IsVerticalLineValid: bool
    Distance: int
    IsDistanceValid: bool


class CellHoriJustify(IntEnum):
    STANDARD = 0
    LEFT = 1
    CENTER = 2
    RIGHT = 3
    BLOCK = 4
    REPEAT = 5


class CellVertJustify(IntEnum):
    STANDARD = 0
    TOP = 1
    CENTER = 2
    BOTTOM = 3
    BLOCK = 4
