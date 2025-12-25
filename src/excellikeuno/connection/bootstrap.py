from __future__ import annotations

from typing import Any, Tuple
from ..calc import Sheet
from ..core import InterfaceNames, UnoObject


def open_calc_document(path: str) -> Tuple[Any, Any, Sheet]:
    """Open a Calc document and return (desktop, document, first_sheet).

    This is a minimal example that relies on the bundled LibreOffice Python.
    It raises RuntimeError if UNO is not available.
    """

    try:
        import uno
        from com.sun.star.beans import PropertyValue
    except ImportError as exc:  # pragma: no cover - depends on LibreOffice runtime
        raise RuntimeError("UNO runtime is not available") from exc

    ctx = uno.getComponentContext()
    smgr = ctx.getServiceManager()
    desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)

    url = uno.systemPathToFileUrl(path)
    properties = tuple([PropertyValue("Hidden", 0, True, 0)])
    document = desktop.loadComponentFromURL(url, "_blank", 0, properties)

    doc_wrapper = UnoObject(document)
    spreadsheet_doc = doc_wrapper.iface(InterfaceNames.X_SPREADSHEET_DOCUMENT)
    sheets = spreadsheet_doc.getSheets()
    first_sheet = sheets.getByIndex(0)
    return desktop, document, Sheet(first_sheet)

def connect_calc() -> Tuple[Any, Any, Sheet]:

    try:
        import uno
        from com.sun.star.beans import PropertyValue
    except ImportError as exc:  # pragma: no cover - depends on LibreOffice runtime
        raise RuntimeError("UNO runtime is not available") from exc

    # UNO接続の準備
    local_ctx = uno.getComponentContext()
    resolver = local_ctx.ServiceManager.createInstanceWithContext(
        "com.sun.star.bridge.UnoUrlResolver", local_ctx)

    # LibreOfficeに接続
    ctx = resolver.resolve(
        "uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")

    smgr = ctx.ServiceManager
    desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)

    # 現在開いているCalcドキュメントを取得し、UnoObject経由でインターフェースを解決
    doc = desktop.getCurrentComponent()
    doc_wrapper = UnoObject(doc)
    spreadsheet_doc = doc_wrapper.iface(InterfaceNames.X_SPREADSHEET_DOCUMENT)
    controller = spreadsheet_doc.getCurrentController()
    sheet = controller.getActiveSheet()

    return desktop, doc, Sheet(sheet)



def wrap_sheet(sheet_obj: Any) -> Sheet:
    """Wrap an existing UNO sheet object in a Sheet helper."""
    return Sheet(sheet_obj)
