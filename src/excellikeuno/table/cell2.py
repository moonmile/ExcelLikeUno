from __future__ import annotations

from typing import TYPE_CHECKING, Any

from ..core import UnoObject
from ..style.border import Borders
from ..typing import BorderLine, BorderLine2, TableBorder, TableBorder2

# UNO runtime does not expose all service names as importable modules; keep type-only imports
if TYPE_CHECKING:
    from com.sun.star.table import CellProperties as UNOCellProperties, Cell as UNOCell, CellContentType
    from com.sun.star.style import CharacterProperties as UNOCharacterProperties, ParagraphProperties as UNOParagraphProperties
    from com.sun.star.sheet import SheetCell, XSheetConditionalEntries
    from com.sun.star.awt import Point, Size
    from com.sun.star.beans import XPropertySet
else:  # type hints only
    UNOCellProperties = UNOCell = Any  # type: ignore
    UNOCharacterProperties = UNOParagraphProperties = Any  # type: ignore
    SheetCell = XSheetConditionalEntries = Any  # type: ignore
    CellContentType = Any  # type: ignore
    Point = Size = XPropertySet = Any  # type: ignore

class Cell2(UnoObject):
    def __init__(self, obj: SheetCell) -> None:
        super().__init__(obj)

    cell_props: UNOCellProperties
    char_props: UNOCharacterProperties
    para_props: UNOParagraphProperties

    raw: SheetCell

    # -- border helpers (standalone; not inheriting Cell) --
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

    def _border_getter(self) -> dict[str, BorderLine]:
        try:
            tb2 = self.raw.TableBorder2
        except Exception:
            tb2 = None
        try:
            tb1 = self.raw.TableBorder
        except Exception:
            tb1 = None

        def pick(name: str) -> BorderLine:
            val2 = getattr(tb2, name, None)
            if val2 is not None and hasattr(val2, "LineStyle"):
                return val2
            val1 = getattr(tb1, name, None)
            if val1 is not None:
                return val1
            return BorderLine()

        return {
            "top": pick("TopLine"),
            "bottom": pick("BottomLine"),
            "left": pick("LeftLine"),
            "right": pick("RightLine"),
        }

    def _border_setter(self, **updates: BorderLine) -> None:
        try:
            current2 = self.raw.TableBorder2
        except Exception:
            current2 = None
        try:
            current1 = self.raw.TableBorder
        except Exception:
            current1 = None

        tb1 = TableBorder(
            TopLine=self._as_border_line(getattr(current1, "TopLine", None)),
            IsTopLineValid=bool(getattr(current1, "IsTopLineValid", False)),
            BottomLine=self._as_border_line(getattr(current1, "BottomLine", None)),
            IsBottomLineValid=bool(getattr(current1, "IsBottomLineValid", False)),
            LeftLine=self._as_border_line(getattr(current1, "LeftLine", None)),
            IsLeftLineValid=bool(getattr(current1, "IsLeftLineValid", False)),
            RightLine=self._as_border_line(getattr(current1, "RightLine", None)),
            IsRightLineValid=bool(getattr(current1, "IsRightLineValid", False)),
            HorizontalLine=self._as_border_line(getattr(current1, "HorizontalLine", None)),
            IsHorizontalLineValid=bool(getattr(current1, "IsHorizontalLineValid", False)),
            VerticalLine=self._as_border_line(getattr(current1, "VerticalLine", None)),
            IsVerticalLineValid=bool(getattr(current1, "IsVerticalLineValid", False)),
            Distance=int(getattr(current1, "Distance", 0)) if current1 is not None else 0,
            IsDistanceValid=bool(getattr(current1, "IsDistanceValid", False)),
        )

        tb2 = TableBorder2(
            TopLine=self._as_border_line2(getattr(current2, "TopLine", None)),
            IsTopLineValid=bool(getattr(current2, "IsTopLineValid", False)),
            BottomLine=self._as_border_line2(getattr(current2, "BottomLine", None)),
            IsBottomLineValid=bool(getattr(current2, "IsBottomLineValid", False)),
            LeftLine=self._as_border_line2(getattr(current2, "LeftLine", None)),
            IsLeftLineValid=bool(getattr(current2, "IsLeftLineValid", False)),
            RightLine=self._as_border_line2(getattr(current2, "RightLine", None)),
            IsRightLineValid=bool(getattr(current2, "IsRightLineValid", False)),
            HorizontalLine=self._as_border_line2(getattr(current2, "HorizontalLine", None)),
            IsHorizontalLineValid=bool(getattr(current2, "IsHorizontalLineValid", False)),
            VerticalLine=self._as_border_line2(getattr(current2, "VerticalLine", None)),
            IsVerticalLineValid=bool(getattr(current2, "IsVerticalLineValid", False)),
            Distance=int(getattr(current2, "Distance", 0)) if current2 is not None else 0,
            IsDistanceValid=bool(getattr(current2, "IsDistanceValid", False)),
        )

        mapping = {
            "top": ("TopLine", "IsTopLineValid"),
            "bottom": ("BottomLine", "IsBottomLineValid"),
            "left": ("LeftLine", "IsLeftLineValid"),
            "right": ("RightLine", "IsRightLineValid"),
        }
        for side, line in updates.items():
            line_attr, valid_attr = mapping.get(side, (None, None))
            if line_attr is None:
                continue
            try:
                setattr(tb1, line_attr, self._as_border_line(line))
                setattr(tb1, valid_attr, True)
            except Exception:
                pass
            try:
                setattr(tb2, line_attr, self._as_border_line2(line, self._as_border_line(line)))
                setattr(tb2, valid_attr, True)
            except Exception:
                pass

        try:
            self.raw.TableBorder = tb1.to_raw()
        except Exception:
            try:
                self.raw.TableBorder = tb1
            except Exception:
                pass
        try:
            self.raw.TableBorder2 = tb2.to_raw()
        except Exception:
            try:
                self.raw.TableBorder2 = tb2
            except Exception:
                pass

    @property
    def borders(self) -> Borders:
        existing = self.__dict__.get("_borders")
        if existing is None:
            existing = Borders(owner=self)
            object.__setattr__(self, "_borders", existing)
        return existing

    # impliment methods of XCell
    @property
    def formula(self) -> str:
        return self.raw.getFormula()
    @formula.setter
    def formula(self, value: str) -> None:
        self.raw.setFormula(value)

    @property
    def value(self) -> float:
        return self.raw.getValue()
    @value.setter
    def value(self, value: float) -> None:
        self.raw.setValue(value)

    @property
    def text(self) -> str:
        return self.formula
    @text.setter
    def text(self, value: str) -> None:
        self.formula = value

    # implement properties of SheetCell
    @property
    def position(self) -> Point:
        return self.raw.Position
    
    @property
    def size(self) -> Size:
        return self.raw.Size
    
    @property
    def formula_local(self) -> str:
        return self.raw.FormulaLocal
    
    @property
    def formula_result_type(self) -> int:
        return self.raw.FormulaResultType
    
    @property
    def conditional_format(self) -> XSheetConditionalEntries:
        return self.raw.ConditionalFormat
    
    @property
    def conditional_format_local(self) -> XSheetConditionalEntries:
        return self.raw.ConditionalFormatLocal
    
    @property
    def validation(self) -> XPropertySet:
        return self.raw.Validation
    
    @property
    def validation_local(self) -> XPropertySet:
        return self.raw.ValidationLocal
    
    @property
    def absolute_name(self) -> str:
        return self.raw.AbsoluteName
    
    @property
    def cell_content_type(self) -> CellContentType:
        return self.raw.CellContentType
    
    @property
    def formula_result_type2(self) -> int:
        return self.raw.FormulaResultType2
    