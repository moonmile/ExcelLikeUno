from __future__ import annotations

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
