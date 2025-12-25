from .calc import Cell, Sheet
from .connection import connect_calc, open_calc_document, wrap_sheet
from .core import InterfaceNames, UnoObject

__all__ = [
    "Cell",
    "Sheet",
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
