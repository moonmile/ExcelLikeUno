from __future__ import annotations

from typing import Any, Iterable, cast

from ..core import UnoObject
from ..typing import InterfaceNames, XCell, XSheetCellRange, XCellRangeAddressable
from ..style.font import Font
from ..style.border import Borders
from .sheet_cell import SheetCell
from ..table.rows import TableRows
from ..table.columns import TableColumns


class SheetCellRange(UnoObject):
    """SheetCell ベースの範囲ラッパー。Range と同等の API で SheetCell を返す。"""

    # navigation
    def cell(self, column: int, row: int) -> SheetCell:
        rng = cast(XSheetCellRange, self.iface(InterfaceNames.X_SHEET_CELL_RANGE))
        return SheetCell(rng.getCellByPosition(int(column), int(row)))

    def subrange(self, start_column: int, start_row: int, end_column: int, end_row: int) -> "SheetCellRange":
        rng = cast(XSheetCellRange, self.iface(InterfaceNames.X_SHEET_CELL_RANGE))
        sub = rng.getCellRangeByPosition(int(start_column), int(start_row), int(end_column), int(end_row))
        return SheetCellRange(sub)

    getCellByPosition = cell  # noqa: N815 - UNO alias
    getCellRangeByPosition = subrange  # noqa: N815 - UNO alias

    # iteration
    def __iter__(self) -> Iterable[SheetCell]:
        rng = cast(XSheetCellRange, self.iface(InterfaceNames.X_SHEET_CELL_RANGE))
        addr = cast(XCellRangeAddressable, self.iface(InterfaceNames.X_CELL_RANGE_ADDRESSABLE)).getRangeAddress()
        row_count = addr.EndRow - addr.StartRow + 1
        col_count = addr.EndColumn - addr.StartColumn + 1
        for row in range(row_count):
            for col in range(col_count):
                yield SheetCell(rng.getCellByPosition(col, row))

    def _first_cell(self) -> SheetCell:
        rng = cast(XSheetCellRange, self.iface(InterfaceNames.X_SHEET_CELL_RANGE))
        return SheetCell(rng.getCellByPosition(0, 0))

    # row/column grouping for sizing
    @property
    def rows(self) -> TableRows:
        colrow = self.iface(InterfaceNames.X_COLUMN_ROW_RANGE)
        if colrow is None:
            raise AttributeError("XColumnRowRange not available on this range")
        return TableRows(colrow.getRows())  # type: ignore[arg-type]

    @property
    def columns(self) -> TableColumns:
        colrow = self.iface(InterfaceNames.X_COLUMN_ROW_RANGE)
        if colrow is None:
            raise AttributeError("XColumnRowRange not available on this range")
        return TableColumns(colrow.getColumns())  # type: ignore[arg-type]

    # font
    @property
    def font(self) -> Font:
        first = self._first_cell()
        return Font(owner=first, setter=self._font_broadcast)

    @font.setter
    def font(self, value: Font) -> None:
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

    # borders
    @property
    def borders(self) -> Borders:
        existing = self.__dict__.get("_borders")
        if existing is None:
            existing = Borders(owner=self._first_cell(), setter=self._border_broadcast)
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
        around = updates.pop("around", None)
        inner = updates.pop("inner", None)
        if around is not None:
            self._apply_around(around)
        if inner is not None:
            self._apply_inner(inner)
        if updates:
            for cell in self:
                cell.borders.apply(**updates)

    def _apply_around(self, line: Any) -> None:
        rng = cast(XSheetCellRange, self.iface(InterfaceNames.X_SHEET_CELL_RANGE))
        addr = cast(XCellRangeAddressable, self.iface(InterfaceNames.X_CELL_RANGE_ADDRESSABLE)).getRangeAddress()
        row_count = addr.EndRow - addr.StartRow + 1
        col_count = addr.EndColumn - addr.StartColumn + 1

        def _to_line(val: Any) -> Any:
            try:
                line_obj = getattr(line, "_clone_line", None)
                if callable(line_obj):
                    return line_obj(val)
            except Exception:
                pass
            return line

        for r in range(row_count):
            for c in range(col_count):
                updates: dict[str, Any] = {}
                if r == 0:
                    updates["top"] = _to_line(line)
                if r == row_count - 1:
                    updates["bottom"] = _to_line(line)
                if c == 0:
                    updates["left"] = _to_line(line)
                if c == col_count - 1:
                    updates["right"] = _to_line(line)
                if not updates:
                    continue
                cell = SheetCell(rng.getCellByPosition(c, r))
                bproxy = cell.borders
                bproxy.apply(**updates)

    def _apply_inner(self, line: Any) -> None:
        rng = cast(XSheetCellRange, self.iface(InterfaceNames.X_SHEET_CELL_RANGE))
        addr = cast(XCellRangeAddressable, self.iface(InterfaceNames.X_CELL_RANGE_ADDRESSABLE)).getRangeAddress()
        row_count = addr.EndRow - addr.StartRow + 1
        col_count = addr.EndColumn - addr.StartColumn + 1

        def _to_line(val: Any) -> Any:
            try:
                line_obj = getattr(line, "_clone_line", None)
                if callable(line_obj):
                    return line_obj(val)
            except Exception:
                pass
            return line

        for r in range(row_count):
            for c in range(col_count):
                updates: dict[str, Any] = {}
                if r > 0:
                    updates["top"] = _to_line(line)
                if r < row_count - 1:
                    updates["bottom"] = _to_line(line)
                if c > 0:
                    updates["left"] = _to_line(line)
                if c < col_count - 1:
                    updates["right"] = _to_line(line)
                cell = SheetCell(rng.getCellByPosition(c, r))
                bproxy = cell.borders
                bproxy.apply(**updates)

    # values
    @property
    def value(self) -> Any:
        rng = cast(XSheetCellRange, self.iface(InterfaceNames.X_SHEET_CELL_RANGE))
        addr = cast(XCellRangeAddressable, self.iface(InterfaceNames.X_CELL_RANGE_ADDRESSABLE)).getRangeAddress()
        row_count = addr.EndRow - addr.StartRow + 1
        col_count = addr.EndColumn - addr.StartColumn + 1
        values: list[list[Any]] = []
        for r in range(row_count):
            row_vals: list[Any] = []
            for c in range(col_count):
                cell = cast(XCell, rng.getCellByPosition(c, r))
                row_vals.append(cell.getFormula())
            values.append(row_vals)
        return values

    @value.setter
    def value(self, value: Any) -> None:
        rng = cast(XSheetCellRange, self.iface(InterfaceNames.X_SHEET_CELL_RANGE))
        addr = cast(XCellRangeAddressable, self.iface(InterfaceNames.X_CELL_RANGE_ADDRESSABLE)).getRangeAddress()
        row_count = addr.EndRow - addr.StartRow + 1
        col_count = addr.EndColumn - addr.StartColumn + 1
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

    @property
    def text(self) -> Any:
        return self.value

    @text.setter
    def text(self, value: Any) -> None:
        self.value = value

    # background color broadcast
    @property
    def backcolor(self) -> int:
        return self._first_cell().backcolor

    @backcolor.setter
    def backcolor(self, value: int) -> None:
        for cell in self:
            cell.backcolor = value

    # row/column sizing
    @property
    def row_height(self) -> int:
        rows = self.rows
        first_row = rows.getByIndex(0)
        return first_row.Height

    @row_height.setter
    def row_height(self, height: int) -> None:
        rows = self.rows
        for i in range(rows.count):
            row = rows.getByIndex(i)
            row.Height = height

    @property
    def column_width(self) -> int:
        columns = self.columns
        first_column = columns.getByIndex(0)
        return first_column.Width

    @column_width.setter
    def column_width(self, width: int) -> None:
        columns = self.columns
        for i in range(columns.count):
            column = columns.getByIndex(i)
            column.Width = width

    def getCells(self) -> list[list[SheetCell]]:
        rng = cast(XSheetCellRange, self.iface(InterfaceNames.X_SHEET_CELL_RANGE))
        addr = cast(XCellRangeAddressable, self.iface(InterfaceNames.X_CELL_RANGE_ADDRESSABLE)).getRangeAddress()
        row_count = addr.EndRow - addr.StartRow + 1
        col_count = addr.EndColumn - addr.StartColumn + 1
        cells: list[list[SheetCell]] = []
        for r in range(row_count):
            row_cells: list[SheetCell] = []
            for c in range(col_count):
                row_cells.append(SheetCell(rng.getCellByPosition(c, r)))
            cells.append(row_cells)
        return cells

    @property
    def cells(self) -> list[list[SheetCell]]:
        return self.getCells()
