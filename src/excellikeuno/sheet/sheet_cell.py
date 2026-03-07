from __future__ import annotations

from typing import Any, TYPE_CHECKING

from ..core import UnoObject
from ..typing import (
    InterfaceNames,
    BorderLine,
    BorderLine2,
    TableBorder,
    TableBorder2,
    CellVertJustify,
    CellHoriJustify,
)
from ..style.font import Font
from ..style.border import Borders

if TYPE_CHECKING:  # pragma: no cover - type hints only
    from com.sun.star.beans import XPropertySet
    from com.sun.star.table import XCell


class RawProps:
    """XPropertySet への透過ラッパー（CellProperties を使わない軽量版）。"""

    def __init__(self, owner: UnoObject) -> None:
        object.__setattr__(self, "_owner", owner)

    def _props(self) -> "XPropertySet":
        return self._owner.iface(InterfaceNames.X_PROPERTY_SET)  # type: ignore[return-value]

    def get_property(self, name: str) -> Any:
        return self._props().getPropertyValue(name)

    def set_property(self, name: str, value: Any) -> None:
        self._props().setPropertyValue(name, value)

    # UNO-style helpers
    def getPropertyValue(self, name: str) -> Any:  # noqa: N802 - UNO naming
        return self.get_property(name)

    def setPropertyValue(self, name: str, value: Any) -> None:  # noqa: N802 - UNO naming
        self.set_property(name, value)

    def __getattr__(self, name: str) -> Any:
        try:
            return self.get_property(name)
        except Exception as exc:  # pragma: no cover - UNO errors bubble up
            raise AttributeError(f"Unknown property: {name}") from exc

    def __setattr__(self, name: str, value: Any) -> None:
        if name.startswith("_"):
            object.__setattr__(self, name, value)
            return
        try:
            self.set_property(name, value)
        except Exception as exc:  # pragma: no cover - UNO errors bubble up
            raise AttributeError(f"Cannot set property: {name}") from exc


class SheetCell(UnoObject):
    """VBAライクな Cell ラッパー。CellProperties ではなく RawProps を利用する。"""

    def __init__(self, cell_obj: Any) -> None:
        super().__init__(cell_obj)

    def __getattr__(self, name: str) -> Any:
        try:
            return getattr(self.props, name)
        except Exception:
            pass
        try:
            return getattr(self.raw, name)
        except Exception as exc:
            raise AttributeError(name) from exc

    def __setattr__(self, name: str, value: Any) -> None:
        if name.startswith("_") or name in ("_obj", "_iface_cache"):
            object.__setattr__(self, name, value)
            return
        descriptor = getattr(self.__class__, name, None)
        if hasattr(descriptor, "__set__"):
            descriptor.__set__(self, value)  # type: ignore[misc]
            return
        try:
            setattr(self.props, name, value)
            return
        except Exception:
            pass
        try:
            setattr(self.raw, name, value)
            return
        except Exception:
            pass
        object.__setattr__(self, name, value)

    # -- internal border helpers (adapted from Cell) --
    def _get_prop(self, name: str) -> Any:
        return self.props.get_property(name)

    def _set_prop(self, name: str, value: Any) -> None:
        self.props.set_property(name, value)

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

    # geometry helpers
    @property
    def position(self) -> Any:
        try:
            return getattr(self.raw, "Position")
        except Exception:
            pass
        try:
            return self.props.get_property("Position")
        except Exception as exc:
            raise AttributeError("position") from exc

    @property
    def column_width(self) -> int:
        try:
            addr = self.raw.getCellAddress()
            sheet = getattr(self.raw, "Spreadsheet", None) or self.raw.getSpreadsheet()
            column = sheet.getColumns().getByIndex(addr.Column)
            return int(getattr(column, "Width"))
        except Exception as exc:
            raise AttributeError("column_width") from exc

    @column_width.setter
    def column_width(self, width: int) -> None:
        try:
            addr = self.raw.getCellAddress()
            sheet = getattr(self.raw, "Spreadsheet", None) or self.raw.getSpreadsheet()
            column = sheet.getColumns().getByIndex(addr.Column)
            column.Width = int(width)
        except Exception as exc:
            raise AttributeError("column_width") from exc

    @property
    def row_height(self) -> int:
        try:
            addr = self.raw.getCellAddress()
            sheet = getattr(self.raw, "Spreadsheet", None) or self.raw.getSpreadsheet()
            row = sheet.getRows().getByIndex(addr.Row)
            return int(getattr(row, "Height"))
        except Exception as exc:
            raise AttributeError("row_height") from exc

    @row_height.setter
    def row_height(self, height: int) -> None:
        try:
            addr = self.raw.getCellAddress()
            sheet = getattr(self.raw, "Spreadsheet", None) or self.raw.getSpreadsheet()
            row = sheet.getRows().getByIndex(addr.Row)
            row.Height = int(height)
        except Exception as exc:
            raise AttributeError("row_height") from exc

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

    def _coerce_hori_justify(self, value: Any) -> CellHoriJustify:
        if isinstance(value, CellHoriJustify):
            return value
        name = getattr(value, "name", None)
        if isinstance(name, str) and name in CellHoriJustify.__members__:
            return CellHoriJustify[name]
        try:
            return CellHoriJustify(int(value))
        except Exception:
            pass
        value_attr = getattr(value, "value", None)
        if isinstance(value_attr, str) and value_attr in CellHoriJustify.__members__:
            return CellHoriJustify[value_attr]
        if isinstance(value_attr, (int, float)):
            return CellHoriJustify(int(value_attr))
        if isinstance(value, str) and value in CellHoriJustify.__members__:
            return CellHoriJustify[value]
        return CellHoriJustify.STANDARD

    def _coerce_vert_justify(self, value: Any) -> CellVertJustify:
        try:
            return CellVertJustify(int(value))
        except Exception:
            val_name = getattr(value, "name", None)
            if isinstance(val_name, str) and val_name in CellVertJustify.__members__:
                return CellVertJustify[val_name]
            val_value = getattr(value, "value", None)
            if isinstance(val_value, (int, float)):
                return CellVertJustify(int(val_value))
            if isinstance(value, str) and value in CellVertJustify.__members__:
                return CellVertJustify[value]
        return CellVertJustify.STANDARD

    @property
    def props(self) -> RawProps:
        existing = self.__dict__.get("_raw_props")
        if existing is None:
            existing = RawProps(self)
            object.__setattr__(self, "_raw_props", existing)
        return existing  # type: ignore[return-value]

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
        try:
            current = value._current()  # type: ignore[attr-defined]
        except Exception:
            try:
                current = dict(value)  # type: ignore[arg-type]
            except Exception:
                current = {}
        if not current:
            return
        Font(owner=self).apply(**current)

    @property
    def value(self) -> float:
        return self.raw.getValue()  # type: ignore[union-attr]

    @value.setter
    def value(self, value: Any) -> None:
        if isinstance(value, (int, float)):
            self.raw.setValue(float(value))  # type: ignore[union-attr]
        else:
            self.raw.setFormula(str(value))  # type: ignore[union-attr]

    @property
    def formula(self) -> str:
        return self.raw.getFormula()  # type: ignore[union-attr]

    @formula.setter
    def formula(self, formula: str) -> None:
        self.raw.setFormula(formula)  # type: ignore[union-attr]

    @property
    def text(self) -> str:
        return self.formula

    @text.setter
    def text(self, text: str) -> None:
        self.raw.setFormula(text)  # type: ignore[union-attr]

    @property
    def vertical_align(self) -> CellVertJustify:
        try:
            return self._coerce_vert_justify(self.props.VertJustify)
        except Exception:
            return CellVertJustify.STANDARD

    @vertical_align.setter
    def vertical_align(self, value: CellVertJustify) -> None:
        self.props.VertJustify = int(value)

    @property
    def horizontal_align(self) -> CellHoriJustify:
        try:
            return self._coerce_hori_justify(self.props.HoriJustify)
        except Exception:
            return CellHoriJustify.STANDARD

    @horizontal_align.setter
    def horizontal_align(self, value: CellHoriJustify) -> None:
        self.props.HoriJustify = int(value)

    # font.color alias for backward compatibility
    @property
    def color(self) -> int:
        return self.font.color

    @color.setter
    def color(self, value: int) -> None:
        self.font.color = value

    # props.CellBackColor alias for backward compatibility
    @property
    def backcolor(self) -> int:
        return self.props.CellBackColor  # type: ignore[attr-defined]

    @backcolor.setter
    def backcolor(self, value: int) -> None:
        self.props.CellBackColor = value  # type: ignore[attr-defined]

    @property
    def rotate(self) -> int:
        return self.props.RotateAngle / 100  # type: ignore[attr-defined]
    @rotate.setter
    def rotate(self, value: int) -> None:
        self.props.RotateAngle = value * 100  # type: ignore[attr-defined]
