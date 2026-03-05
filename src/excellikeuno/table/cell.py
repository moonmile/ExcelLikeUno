from __future__ import annotations

from typing import Any, cast
from numbers import Number

from excellikeuno.typing.calc import BorderLineStyle

from ..core import UnoObject
from ..typing import (
    Color,
    BorderLine,
    BorderLine2,
    CellHoriJustify,
    CellOrientation,
    CellVertJustify,
    CellProtection,
    InterfaceNames,
    ShadowFormat,
    TableBorder,
    TableBorder2,
    XCell,
    XPropertySet,
)
from ..typing.structs import Point
from ..style.font import Font
from ..style.border import Borders
from .cell_properties import CellProperties
from ..style.character_properties import CharacterProperties
from .rows import TableRows


class Cell(UnoObject):
    # -- internal border helpers --
    def _as_border_line(self, value: Any, fallback: BorderLine | None = None) -> BorderLine:
        try:
            return BorderLine(
                Color=int(getattr(value, "Color", getattr(fallback, "Color", 0))),
                InnerLineWidth=int(getattr(value, "InnerLineWidth", getattr(fallback, "InnerLineWidth", 0))),
                OuterLineWidth=int(getattr(value, "OuterLineWidth", getattr(fallback, "OuterLineWidth", 0))),
                LineDistance=int(getattr(value, "LineDistance", getattr(fallback, "LineDistance", 0))),
            )
        except Exception:
            return BorderLine()

    def _as_border_line2(self, value: Any, fallback: BorderLine | None = None) -> BorderLine2:
        try:
            base = self._as_border_line(value, fallback)
            return BorderLine2(
                Color=base.Color,
                InnerLineWidth=base.InnerLineWidth,
                OuterLineWidth=base.OuterLineWidth,
                LineDistance=base.LineDistance,
                LineStyle=int(getattr(value, "LineStyle", 0)),
                LineWidth=int(getattr(value, "LineWidth", getattr(fallback, "OuterLineWidth", 0))),
            )
        except Exception:
            return BorderLine2()

    def _get_prop(self, name: str) -> Any:
        props = cast(XPropertySet, self.iface(InterfaceNames.X_PROPERTY_SET))
        return props.getPropertyValue(name)

    def _set_prop(self, name: str, value: Any) -> None:
        props = cast(XPropertySet, self.iface(InterfaceNames.X_PROPERTY_SET))
        props.setPropertyValue(name, value)

    def _border_getter(self) -> dict[str, BorderLine]:
        def prefer_line2(name: str) -> BorderLine:
            try:
                val2 = self._get_prop(f"{name}2")
                if hasattr(val2, "LineStyle"):
                    return val2
            except Exception:
                pass
            return self._as_border_line(self._get_prop(name))

        return {
            "top": prefer_line2("TopBorder"),
            "bottom": prefer_line2("BottomBorder"),
            "left": prefer_line2("LeftBorder"),
            "right": prefer_line2("RightBorder"),
        }

    def _border_setter(self, **updates: BorderLine) -> None:
        try:
            current_table = self.TableBorder
        except Exception:
            current_table = None
        table_border = TableBorder(
            TopLine=self._as_border_line(getattr(current_table, "TopLine", None)),
            IsTopLineValid=bool(getattr(current_table, "IsTopLineValid", False)),
            BottomLine=self._as_border_line(getattr(current_table, "BottomLine", None)),
            IsBottomLineValid=bool(getattr(current_table, "IsBottomLineValid", False)),
            LeftLine=self._as_border_line(getattr(current_table, "LeftLine", None)),
            IsLeftLineValid=bool(getattr(current_table, "IsLeftLineValid", False)),
            RightLine=self._as_border_line(getattr(current_table, "RightLine", None)),
            IsRightLineValid=bool(getattr(current_table, "IsRightLineValid", False)),
            HorizontalLine=self._as_border_line(getattr(current_table, "HorizontalLine", None)),
            IsHorizontalLineValid=bool(getattr(current_table, "IsHorizontalLineValid", False)),
            VerticalLine=self._as_border_line(getattr(current_table, "VerticalLine", None)),
            IsVerticalLineValid=bool(getattr(current_table, "IsVerticalLineValid", False)),
            Distance=int(getattr(current_table, "Distance", 0)) if current_table is not None else 0,
            IsDistanceValid=bool(getattr(current_table, "IsDistanceValid", False)),
        )

        try:
            current_table2 = self.TableBorder2
        except Exception:
            current_table2 = None

        table_border2 = TableBorder2(
            TopLine=self._as_border_line2(getattr(current_table2, "TopLine", None)),
            IsTopLineValid=bool(getattr(current_table2, "IsTopLineValid", False)),
            BottomLine=self._as_border_line2(getattr(current_table2, "BottomLine", None)),
            IsBottomLineValid=bool(getattr(current_table2, "IsBottomLineValid", False)),
            LeftLine=self._as_border_line2(getattr(current_table2, "LeftLine", None)),
            IsLeftLineValid=bool(getattr(current_table2, "IsLeftLineValid", False)),
            RightLine=self._as_border_line2(getattr(current_table2, "RightLine", None)),
            IsRightLineValid=bool(getattr(current_table2, "IsRightLineValid", False)),
            HorizontalLine=self._as_border_line2(getattr(current_table2, "HorizontalLine", None)),
            IsHorizontalLineValid=bool(getattr(current_table2, "IsHorizontalLineValid", False)),
            VerticalLine=self._as_border_line2(getattr(current_table2, "VerticalLine", None)),
            IsVerticalLineValid=bool(getattr(current_table2, "IsVerticalLineValid", False)),
            Distance=int(getattr(current_table2, "Distance", 0)) if current_table2 is not None else 0,
            IsDistanceValid=bool(getattr(current_table2, "IsDistanceValid", False)),
        )

        mapping = {
            "top": "TopBorder",
            "bottom": "BottomBorder",
            "left": "LeftBorder",
            "right": "RightBorder",
        }
        valid_flags = {
            "top": "IsTopLineValid",
            "bottom": "IsBottomLineValid",
            "left": "IsLeftLineValid",
            "right": "IsRightLineValid",
        }
        line_attrs = {
            "top": "TopLine",
            "bottom": "BottomLine",
            "left": "LeftLine",
            "right": "RightLine",
        }
        for side, line in updates.items():
            attr = mapping.get(side)
            if attr is None:
                continue
            line_struct = line.to_raw() if hasattr(line, "to_raw") else line
            try:
                self._set_prop(attr, line_struct)
            except Exception:
                setattr(self, attr, line)
            try:
                line2 = self._as_border_line2(line, self._as_border_line(line))
                self._set_prop(f"{attr}2", line2.to_raw())
            except Exception:
                pass
            valid_attr = valid_flags.get(side)
            line_attr = line_attrs.get(side)
            try:
                if line_attr is not None:
                    setattr(table_border, line_attr, self._as_border_line(line))
            except Exception:
                pass
            try:
                if valid_attr is not None:
                    setattr(table_border, valid_attr, True)
            except Exception:
                continue

            try:
                if line_attr is not None:
                    setattr(table_border2, line_attr, self._as_border_line2(line, self._as_border_line(line)))
            except Exception:
                pass
            try:
                if valid_attr is not None:
                    setattr(table_border2, valid_attr, True)
            except Exception:
                continue

        try:
            self.TableBorder = table_border.to_raw()
        except Exception:
            try:
                self.TableBorder = table_border
            except Exception:
                pass

        try:
            self.TableBorder2 = table_border2.to_raw()
        except Exception:
            try:
                self.TableBorder2 = table_border2
            except Exception:
                pass

    @property
    def props(self) -> CellProperties:
        existing = self.__dict__.get("_properties")
        if existing is None:
            existing = CellProperties(self.iface(InterfaceNames.X_PROPERTY_SET))
            object.__setattr__(self, "_properties", existing)
        return cast(CellProperties, existing)

    @property
    def char_props(self) -> CharacterProperties:
        existing = self.__dict__.get("_character_properties")
        if existing is None:
            existing = CharacterProperties(self.iface(InterfaceNames.X_PROPERTY_SET))
            object.__setattr__(self, "_character_properties", existing)
        return cast(CharacterProperties, existing)

    # Backward-compatible alias for font proxy (expects character_properties attribute)
    @property
    def character_properties(self) -> CharacterProperties:
        return self.char_props

    @property
    def borders(self) -> Borders:
        existing = self.__dict__.get("_borders")
        if existing is None:
            existing = Borders(owner=self)
            object.__setattr__(self, "_borders", existing)
        return existing

    @borders.setter
    def borders(self, value: Borders) -> None:
        current_proxy = self.__dict__.get("_borders") or Borders(owner=self)
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
        current_proxy.apply(**current)

    @property
    def font(self) -> Font:
        return Font(owner=self)

    @font.setter
    def font(self, value: Font) -> None:
        # Accept a Font proxy or plain Font config
        try:
            current = value._current()  # type: ignore[attr-defined]
        except Exception:
            current = {}
        if not current:
            return
        Font(owner=self).apply(**current)

    # 行の高さ
    @property
    def row_height(self) -> int:
        # Resolve the parent sheet via XSheetCellRange and read the row's Height property (1/100 mm)
        sheet_range = self.iface(InterfaceNames.X_SHEET_CELL_RANGE)
        if sheet_range is None:
            raise AttributeError("XSheetCellRange not available for this cell")

        sheet = sheet_range.getSpreadsheet()
        row_idx = self.raw.getCellAddress().Row
        rows = TableRows(sheet.getRows())
        row = rows.getByIndex(row_idx)
        try:
            return int(row.raw.getPropertyValue("Height"))
        except Exception:
            # Some UNO implementations expose Height as a direct attribute
            return int(getattr(row.raw, "Height"))
    
    @row_height.setter
    def row_height(self, height: int) -> None:
        sheet_range = self.iface(InterfaceNames.X_SHEET_CELL_RANGE)
        if sheet_range is None:
            raise AttributeError("XSheetCellRange not available for this cell")

        sheet = sheet_range.getSpreadsheet()
        row_idx = self.raw.getCellAddress().Row
        rows = TableRows(sheet.getRows())
        row = rows.getByIndex(row_idx)

        # Disable optimal height before setting an explicit height when available
        try:
            row.raw.setPropertyValue("OptimalHeight", False)
        except Exception:
            pass

        try:
            row.raw.setPropertyValue("Height", int(height))
        except Exception:
            setattr(row.raw, "Height", int(height))

    @property
    def column_width(self) -> int:
        # Resolve the parent sheet via XSheetCellRange and read the column's Width property (1/100 mm)
        sheet_range = self.iface(InterfaceNames.X_SHEET_CELL_RANGE)
        if sheet_range is None:
            raise AttributeError("XSheetCellRange not available for this cell")

        sheet = sheet_range.getSpreadsheet()
        col_idx = self.raw.getCellAddress().Column
        columns = TableRows(sheet.getColumns())
        column = columns.getByIndex(col_idx)
        try:
            return int(column.raw.getPropertyValue("Width"))
        except Exception:
            # Some UNO implementations expose Width as a direct attribute
            return int(getattr(column.raw, "Width"))
        
    @column_width.setter
    def column_width(self, width: int) -> None:
        sheet_range = self.iface(InterfaceNames.X_SHEET_CELL_RANGE)
        if sheet_range is None:
            raise AttributeError("XSheetCellRange not available for this cell")

        sheet = sheet_range.getSpreadsheet()
        col_idx = self.raw.getCellAddress().Column
        columns = TableRows(sheet.getColumns())
        column = columns.getByIndex(col_idx)

        # Disable optimal width before setting an explicit width when available
        try:
            column.raw.setPropertyValue("OptimalWidth", False)
        except Exception:
            pass

        try:
            column.raw.setPropertyValue("Width", int(width))
        except Exception:
            setattr(column.raw, "Width", int(width))

    @property 
    def position(self) -> Point:
        pos = self.props.getPropertyValue("Position")
        return Point(pos.X, pos.Y)

    @property
    def value(self) -> float:
        cell = cast(XCell, self.iface(InterfaceNames.X_CELL))
        return cell.getValue()

    @value.setter
    def value(self, value: Any) -> None:
        cell_iface = cast(XCell, self.iface(InterfaceNames.X_CELL))
        if isinstance(value, Number):
            cell_iface.setValue(float(value))
        else:
            # Fallback to string/formula set
            cell_iface.setFormula(str(value))

    @property
    def formula(self) -> str:
        cell = cast(XCell, self.iface(InterfaceNames.X_CELL))
        return cell.getFormula()

    @formula.setter
    def formula(self, formula: str) -> None:
        cell = cast(XCell, self.iface(InterfaceNames.X_CELL))
        cell.setFormula(formula)

    @property
    def text(self) -> str:
        cell = cast(XCell, self.iface(InterfaceNames.X_CELL))
        return cell.getFormula()

    @text.setter
    def text(self, text: str) -> None:
        cell = cast(XCell, self.iface(InterfaceNames.X_CELL))
        cell.setFormula(text)


    @property
    def vertical_align(self) -> CellVertJustify:
        return CellVertJustify(self.props.VertJustify)
    @vertical_align.setter
    def vertical_align(self, value: CellVertJustify) -> None:
        self.props.VertJustify = int(value)

    @property
    def horizontal_align(self) -> CellHoriJustify:
        return CellHoriJustify(self.props.HoriJustify)
    @horizontal_align.setter
    def horizontal_align(self, value: CellHoriJustify) -> None:
        self.props.HoriJustify = int(value)
    