from .connection import (
    active_document,
    active_sheet,
    connect_calc,
    connect_calc_script,
    connect_writer,
    get_desktop,
    new_calc_document,
    open_calc_document,
    this_document,
    this_sheet,
    wrap_sheet,
)
from .core import UnoObject
from .core.calc_document import CalcDocument
from .core.writer_document import WriterDocument
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
from .table import Cell, CellProperties, Sheet, PivotTable, PivotTables
from .chart import Chart, ChartCollection

__all__ = [
    "Cell",
    "CellProperties",
    "Chart",
    "ChartCollection",
    "PivotTable",
    "PivotTables",
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
    "CalcDocument",
    "WriterDocument",
    "connect_calc",
    "connect_writer",
    "open_calc_document",
    "new_calc_document",
    "get_desktop",
    "active_document",
    "this_document",
    "active_sheet",
    "this_sheet",
    "wrap_sheet",
    "connect_calc_script",
    "InterfaceNames",
    "UnoObject",
]

# Provide uno.connect_calc convenience when UNO runtime is available.
try:
    import uno  # type: ignore

    if not hasattr(uno, "connect_calc"):
        uno.connect_calc = connect_calc  # type: ignore[attr-defined]
    if not hasattr(uno, "connect_calc_script"):
        uno.connect_calc_script = connect_calc_script  # type: ignore[attr-defined]
    if not hasattr(uno, "connect_writer"):
        uno.connect_writer = connect_writer  # type: ignore[attr-defined]
    if not hasattr(uno, "WriterDocument"):
        uno.WriterDocument = WriterDocument  # type: ignore[attr-defined]
except Exception:
    # Ignore when UNO runtime is absent; normal imports still work.
    pass
