from __future__ import annotations

from typing import TYPE_CHECKING, Any, Iterable, cast

from ..core import UnoObject
from ..typing import (
    BorderLine,
    BorderLine2,
    CellHoriJustify,
    CellOrientation,
    CellProtection,
    CellVertJustify,
    Color,
    InterfaceNames,
    ShadowFormat,
    TableBorder,
    TableBorder2,
    XCell,
    XCellRangeAddressable,
    XSheetCellRange,
    XTableColumns,
    XTableRows,
    XMergeable,
)
from ..style.font import Font
from ..style.border import Borders
from .cell import Cell
from .cell_properties import CellProperties

if TYPE_CHECKING:
    from .columns import TableColumns
    from .rows import TableRows


class Range(UnoObject):
    """Wraps a UNO cell range and exposes cell-level access."""

    # --- navigation helpers -------------------------------------------------
    def cell(self, column: int, row: int) -> Cell:
        rng = cast(XSheetCellRange, self.iface(InterfaceNames.X_SHEET_CELL_RANGE))
        return Cell(rng.getCellByPosition(int(column), int(row)))

    def subrange(self, start_column: int, start_row: int, end_column: int, end_row: int) -> "Range":
        rng = cast(XSheetCellRange, self.iface(InterfaceNames.X_SHEET_CELL_RANGE))
        sub = rng.getCellRangeByPosition(int(start_column), int(start_row), int(end_column), int(end_row))
        return Range(sub)

    def cell_properties(self, column: int, row: int) -> CellProperties:
        return self.cell(column, row).properties

    # PascalCase/UNO-style aliases
    def CellProperties(self, column: int, row: int) -> CellProperties:  # noqa: N802 - UNO alias
        return self.cell_properties(column, row)

    getCellByPosition = cell  # noqa: N815 - UNO alias
    getCellRangeByPosition = subrange  # noqa: N815 - UNO alias

    # --- row/column grouping -------------------------------------------------
    @property
    def rows(self) -> TableRows:
        from .rows import TableRows  # local import to avoid cycles

        colrow = self.iface(InterfaceNames.X_COLUMN_ROW_RANGE)
        if colrow is None:
            raise AttributeError("XColumnRowRange not available on this range")
        return TableRows(colrow.getRows())  # type: ignore[arg-type]

    def row(self, index: int) -> "TableRow":
        return self.rows.getByIndex(int(index))

    @property
    def columns(self) -> XTableColumns:
        from .columns import TableColumns  # local import to avoid cycles

        colrow = self.iface(InterfaceNames.X_COLUMN_ROW_RANGE)
        if colrow is None:
            raise AttributeError("XColumnRowRange not available on this range")
        return TableColumns(colrow.getColumns())  # type: ignore[arg-type]

    def column(self, index: int) -> "TableColumn":
        return self.columns.getByIndex(int(index))

    # --- iteration ----------------------------------------------------------
    def __iter__(self) -> Iterable[Cell]:  # pragma: no cover - helper for convenience
        rng = cast(XSheetCellRange, self.iface(InterfaceNames.X_SHEET_CELL_RANGE))
        addr = cast(XCellRangeAddressable, self.iface(InterfaceNames.X_CELL_RANGE_ADDRESSABLE)).getRangeAddress()
        row_count = addr.EndRow - addr.StartRow + 1
        col_count = addr.EndColumn - addr.StartColumn + 1
        for row in range(row_count):
            for col in range(col_count):
                yield Cell(rng.getCellByPosition(col, row))

    def _first_cell(self) -> Cell:
        """Return the top-left cell without iterating the whole range."""
        rng = cast(XSheetCellRange, self.iface(InterfaceNames.X_SHEET_CELL_RANGE))
        return Cell(rng.getCellByPosition(0, 0))

    def __repr__(self) -> str:  # pragma: no cover - debugging helper
        return f"Range({self.raw!r})"

    # --- property access ----------------------------------------------------
    @property
    def properties(self) -> CellProperties:
        existing = self.__dict__.get("_properties")
        if existing is None:
            existing = CellProperties(self.iface(InterfaceNames.X_PROPERTY_SET))
            object.__setattr__(self, "_properties", existing)
        return cast(CellProperties, existing)

    @property
    def props(self) -> CellProperties:
        return self.properties

    @property
    def font(self) -> Font:
        # Use the top-left cell as representative for getter; setter broadcasts to all cells in range.
        first_cell = self._first_cell()
        return Font(owner=first_cell, setter=self._font_broadcast)

    @font.setter
    def font(self, value: Font) -> None:
        # Accept a Font proxy or plain Font config dict-like and broadcast to the range
        current = {}
        try:
            current = value._current()  # type: ignore[attr-defined]
        except Exception:
            try:
                current = dict(value)  # type: ignore[arg-type]
            except Exception:
                current = {}
        if not current:
            return
        self._font_broadcast(**current)

    def _font_broadcast(self, **updates: Any) -> None:
        for cell in self:
            Font(owner=cell).apply(**updates)

    @property
    def borders(self) -> Borders:
        existing = self.__dict__.get("_borders")
        if existing is None:
            first_cell = self._first_cell()
            existing = Borders(owner=first_cell, setter=self._border_broadcast)
            object.__setattr__(self, "_borders", existing)
        return existing

    @borders.setter
    def borders(self, value: Borders) -> None:
        current_proxy = self.__dict__.get("_borders") or Borders(owner=self._first_cell(), setter=self._border_broadcast)
        object.__setattr__(self, "_borders", current_proxy)
        try:
            current = value._current()  # type: ignore[attr-defined]
        except Exception:
            try:
                current = dict(value)  # type: ignore[arg-type]
            except Exception:
                current = {}
        if not current:
            return
        self._border_broadcast(**current)

    def _border_broadcast(self, **updates: Any) -> None:
        for cell in self:
            Borders(owner=cell).apply(**updates)

    # CellProperties shortcuts for IDE completion
    @property
    def CellStyle(self) -> str:
        return self.properties.CellStyle

    @CellStyle.setter
    def CellStyle(self, value: str) -> None:
        self.properties.CellStyle = value

    @property
    def CellBackColor(self) -> Color:
        return self.properties.CellBackColor

    @CellBackColor.setter
    def CellBackColor(self, value: Color) -> None:
        self.properties.CellBackColor = value

    @property
    def IsCellBackgroundTransparent(self) -> bool:
        return self.properties.IsCellBackgroundTransparent

    @IsCellBackgroundTransparent.setter
    def IsCellBackgroundTransparent(self, value: bool) -> None:
        self.properties.IsCellBackgroundTransparent = value

    @property
    def HoriJustify(self) -> CellHoriJustify:
        return self.properties.HoriJustify

    @HoriJustify.setter
    def HoriJustify(self, value: CellHoriJustify | int) -> None:
        self.properties.HoriJustify = value

    @property
    def VertJustify(self) -> CellVertJustify:
        return self.properties.VertJustify

    @VertJustify.setter
    def VertJustify(self, value: CellVertJustify | int) -> None:
        self.properties.VertJustify = value

    @property
    def IsTextWrapped(self) -> bool:
        return self.properties.IsTextWrapped

    @IsTextWrapped.setter
    def IsTextWrapped(self, value: bool) -> None:
        self.properties.IsTextWrapped = value

    @property
    def ParaIndent(self) -> int:
        return self.properties.ParaIndent

    @ParaIndent.setter
    def ParaIndent(self, value: int) -> None:
        self.properties.ParaIndent = value

    @property
    def Orientation(self) -> CellOrientation:
        return self.properties.Orientation

    @Orientation.setter
    def Orientation(self, value: CellOrientation) -> None:
        self.properties.Orientation = value

    @property
    def RotateAngle(self) -> int:
        return self.properties.RotateAngle

    @RotateAngle.setter
    def RotateAngle(self, value: int) -> None:
        self.properties.RotateAngle = value

    @property
    def RotateReference(self) -> int:
        return self.properties.RotateReference

    @RotateReference.setter
    def RotateReference(self, value: int) -> None:
        self.properties.RotateReference = value

    @property
    def AsianVerticalMode(self) -> bool:
        return self.properties.AsianVerticalMode

    @AsianVerticalMode.setter
    def AsianVerticalMode(self, value: bool) -> None:
        self.properties.AsianVerticalMode = value

    @property
    def TableBorder(self) -> TableBorder:
        return self.properties.TableBorder

    @TableBorder.setter
    def TableBorder(self, value: TableBorder) -> None:
        self.properties.TableBorder = value

    @property
    def TopBorder(self) -> BorderLine:
        return self.properties.TopBorder

    @TopBorder.setter
    def TopBorder(self, value: BorderLine) -> None:
        self.properties.TopBorder = value

    @property
    def BottomBorder(self) -> BorderLine:
        return self.properties.BottomBorder

    @BottomBorder.setter
    def BottomBorder(self, value: BorderLine) -> None:
        self.properties.BottomBorder = value

    @property
    def LeftBorder(self) -> BorderLine:
        return self.properties.LeftBorder

    @LeftBorder.setter
    def LeftBorder(self, value: BorderLine) -> None:
        self.properties.LeftBorder = value

    @property
    def RightBorder(self) -> BorderLine:
        return self.properties.RightBorder

    @RightBorder.setter
    def RightBorder(self, value: BorderLine) -> None:
        self.properties.RightBorder = value

    @property
    def NumberFormat(self) -> int:
        return self.properties.NumberFormat

    @NumberFormat.setter
    def NumberFormat(self, value: int) -> None:
        self.properties.NumberFormat = value

    @property
    def ShadowFormat(self) -> ShadowFormat:
        return self.properties.ShadowFormat

    @ShadowFormat.setter
    def ShadowFormat(self, value: ShadowFormat) -> None:
        self.properties.ShadowFormat = value

    @property
    def CellProtection(self) -> CellProtection:
        return self.properties.CellProtection

    @CellProtection.setter
    def CellProtection(self, value: CellProtection) -> None:
        self.properties.CellProtection = value

    @property
    def UserDefinedAttributes(self) -> Any:
        return self.properties.UserDefinedAttributes

    @UserDefinedAttributes.setter
    def UserDefinedAttributes(self, value: Any) -> None:
        self.properties.UserDefinedAttributes = value

    @property
    def DiagonalTLBR(self) -> BorderLine:
        return self.properties.DiagonalTLBR

    @DiagonalTLBR.setter
    def DiagonalTLBR(self, value: BorderLine) -> None:
        self.properties.DiagonalTLBR = value

    @property
    def DiagonalBLTR(self) -> BorderLine:
        return self.properties.DiagonalBLTR

    @DiagonalBLTR.setter
    def DiagonalBLTR(self, value: BorderLine) -> None:
        self.properties.DiagonalBLTR = value

    @property
    def ShrinkToFit(self) -> bool:
        return self.properties.ShrinkToFit

    @ShrinkToFit.setter
    def ShrinkToFit(self, value: bool) -> None:
        self.properties.ShrinkToFit = value

    @property
    def TableBorder2(self) -> TableBorder2:
        return self.properties.TableBorder2

    @TableBorder2.setter
    def TableBorder2(self, value: TableBorder2) -> None:
        self.properties.TableBorder2 = value

    @property
    def TopBorder2(self) -> BorderLine2:
        return self.properties.TopBorder2

    @TopBorder2.setter
    def TopBorder2(self, value: BorderLine2) -> None:
        self.properties.TopBorder2 = value

    @property
    def BottomBorder2(self) -> BorderLine2:
        return self.properties.BottomBorder2

    @BottomBorder2.setter
    def BottomBorder2(self, value: BorderLine2) -> None:
        self.properties.BottomBorder2 = value

    @property
    def LeftBorder2(self) -> BorderLine2:
        return self.properties.LeftBorder2

    @LeftBorder2.setter
    def LeftBorder2(self, value: BorderLine2) -> None:
        self.properties.LeftBorder2 = value

    @property
    def RightBorder2(self) -> BorderLine2:
        return self.properties.RightBorder2

    @RightBorder2.setter
    def RightBorder2(self, value: BorderLine2) -> None:
        self.properties.RightBorder2 = value

    @property
    def DiagonalTLBR2(self) -> BorderLine2:
        return self.properties.DiagonalTLBR2

    @DiagonalTLBR2.setter
    def DiagonalTLBR2(self, value: BorderLine2) -> None:
        self.properties.DiagonalTLBR2 = value

    @property
    def DiagonalBLTR2(self) -> BorderLine2:
        return self.properties.DiagonalBLTR2

    @DiagonalBLTR2.setter
    def DiagonalBLTR2(self, value: BorderLine2) -> None:
        self.properties.DiagonalBLTR2 = value

    @property
    def CellInteropGrabBag(self) -> Any:
        return self.properties.CellInteropGrabBag

    @CellInteropGrabBag.setter
    def CellInteropGrabBag(self, value: Any) -> None:
        self.properties.CellInteropGrabBag = value

    # --- values -------------------------------------------------------------
    @property
    def value(self) -> Any:
        rng = cast(XSheetCellRange, self.iface(InterfaceNames.X_SHEET_CELL_RANGE))
        addr = cast(XCellRangeAddressable, self.iface(InterfaceNames.X_CELL_RANGE_ADDRESSABLE)).getRangeAddress()
        row_count = addr.EndRow - addr.StartRow + 1
        col_count = addr.EndColumn - addr.StartColumn + 1

        values: list[list[Any]] = []
        for row_idx in range(row_count):
            row_vals: list[Any] = []
            for col_idx in range(col_count):
                cell = cast(XCell, rng.getCellByPosition(col_idx, row_idx))
                row_vals.append(cell.getFormula())
            values.append(row_vals)
        return values

    @value.setter
    def value(self, value: Any) -> None:
        rng = cast(XSheetCellRange, self.iface(InterfaceNames.X_SHEET_CELL_RANGE))
        addr = cast(XCellRangeAddressable, self.iface(InterfaceNames.X_CELL_RANGE_ADDRESSABLE)).getRangeAddress()
        row_count = addr.EndRow - addr.StartRow + 1
        col_count = addr.EndColumn - addr.StartColumn + 1

        # Normalize incoming to 2D list matching range shape
        matrix: list[list[Any]]
        if isinstance(value, (list, tuple)) and value and isinstance(value[0], (list, tuple)):
            matrix = [list(row) for row in value]  # type: ignore[arg-type]
        elif isinstance(value, (list, tuple)):
            matrix = [list(value)]  # type: ignore[list-item]
        else:
            matrix = [[value]]

        for r_idx in range(row_count):
            for c_idx in range(col_count):
                v = matrix[r_idx][c_idx] if r_idx < len(matrix) and c_idx < len(matrix[r_idx]) else None
                cell = cast(XCell, rng.getCellByPosition(c_idx, r_idx))
                cell.setFormula("" if v is None else str(v))

    def getCells(self) -> list[list[Cell]]:
        rng = cast(XSheetCellRange, self.iface(InterfaceNames.X_SHEET_CELL_RANGE))
        addr = cast(XCellRangeAddressable, self.iface(InterfaceNames.X_CELL_RANGE_ADDRESSABLE)).getRangeAddress()
        row_count = addr.EndRow - addr.StartRow + 1
        col_count = addr.EndColumn - addr.StartColumn + 1

        cells: list[list[Cell]] = []
        for row_idx in range(row_count):
            row_cells: list[Cell] = []
            for col_idx in range(col_count):
                row_cells.append(Cell(rng.getCellByPosition(col_idx, row_idx)))
            cells.append(row_cells)
        return cells

    # セル結合/解除
    def merge(self, merged: bool = True) -> None:
        iface_name = InterfaceNames.X_MERGEABLE
        # Prefer interface if available, otherwise call raw.merge when provided.
        raw_merge = getattr(self.raw, "merge", None)
        if callable(raw_merge):
            raw_merge(bool(merged))
            return
        raise AttributeError("merge not supported on this range")

    def unmerge(self) -> None:
        self.merge(False)

    def is_merged(self) -> bool:
        raw_is_merged = getattr(self.raw, "isMerged", None)
        if callable(raw_is_merged):
            return bool(raw_is_merged())
        raise AttributeError("isMerged not supported on this range")

    # 行の高さ
    @property
    def row_height(self) -> int:
        # 先頭行の高さを返す
        rows = self.rows
        first_row = rows.getByIndex(0)
        return first_row.Height
    @row_height.setter
    def row_height(self, height: int) -> None:
        # 範囲内のすべての行の高さを設定する
        row_count = self.rows.count
        for i in range(row_count):
            row = self.rows.getByIndex(i)
            row.Height = height

    # 列の幅
    @property
    def column_width(self) -> int:
        # 先頭列の幅を返す
        columns = self.columns
        first_column = columns.getByIndex(0)
        return first_column.Width
    @column_width.setter
    def column_width(self, width: int) -> None:
        # 範囲内のすべての列の幅を設定する
        column_count = self.columns.count
        for i in range(column_count):
            column = self.columns.getByIndex(i)
            column.Width = width

    @property
    def cells(self) -> list[list[Cell]]:
        return self.getCells()

    # --- dynamic property passthrough --------------------------------------
    def get_property(self, name: str) -> Any:
        return self.properties.get_property(name)

    def set_property(self, name: str, value: Any) -> None:
        self.properties.set_property(name, value)

    def __getattr__(self, name: str) -> Any:
        if name in {"properties", "props"}:
            raise AttributeError(name)
        try:
            return self.get_property(name)
        except Exception as exc:  # pragma: no cover - bubble up UNO failures
            raise AttributeError(f"Unknown range property: {name}") from exc

    def __setattr__(self, name: str, value: Any) -> None:
        if name.startswith("_"):
            return object.__setattr__(self, name, value)
        cls_attr = getattr(type(self), name, None)
        if isinstance(cls_attr, property):
            setter = cls_attr.fset
            if setter is None:
                raise AttributeError(f"can't set attribute {name}")
            setter(self, value)
            return
        try:
            self.set_property(name, value)
        except Exception:
            object.__setattr__(self, name, value)


class TableRow(Range):
    """Lightweight wrapper for a single table row range."""


class TableColumn(Range):
    """Lightweight wrapper for a single table column range."""
