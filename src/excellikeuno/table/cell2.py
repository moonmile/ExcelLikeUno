from __future__ import annotations

from typing import TYPE_CHECKING, Any

from ..core import UnoObject

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
    cell_props: UNOCellProperties
    char_props: UNOCharacterProperties
    para_props: UNOParagraphProperties

    raw: SheetCell

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
    