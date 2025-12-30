from .connection import connect_calc, open_calc_document, wrap_sheet
from .core import UnoObject
from .core.calc_document import Document
from .typing import InterfaceNames
from .drawing import (
    ClosedBezierShape,
    ConnectorShape,
    ControlShape,
    CustomShape,
    EllipseShape,
    GroupShape,
    LineShape,
    MeasureShape,
    OpenBezierShape,
    PageShape,
    PolyLineShape,
    PolyPolygonBezierShape,
    PolyPolygonShape,
    RectangleShape,
    Shape,
    TextShape,
)
from .table import Cell, CellProperties, Sheet

__all__ = [
    "Cell",
    "CellProperties",
    "Sheet",
    "Shape",
    "ConnectorShape",
    "LineShape",
    "RectangleShape",
    "EllipseShape",
    "PolyLineShape",
    "PolyPolygonShape",
    "PolyPolygonBezierShape",
    "TextShape",
    "ClosedBezierShape",
    "ControlShape",
    "CustomShape",
    "GroupShape",
    "MeasureShape",
    "OpenBezierShape",
    "PageShape",
    "Document",
    "connect_calc",
    "open_calc_document",
    "wrap_sheet",
    "InterfaceNames",
    "UnoObject",
]

# Provide uno.connect_calc convenience when UNO runtime is available.
try:
    import uno  # type: ignore

    if not hasattr(uno, "connect_calc"):
        uno.connect_calc = connect_calc  # type: ignore[attr-defined]
except Exception:
    # Ignore when UNO runtime is absent; normal imports still work.
    pass
