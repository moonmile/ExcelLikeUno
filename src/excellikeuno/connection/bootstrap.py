from __future__ import annotations

from typing import Any, Tuple
from ..core import UnoObject
from ..core.calc_document import CalcDocument
from ..core.writer_document import WriterDocument
from ..typing import InterfaceNames
from ..table import Sheet

current_desktop: Any = None

def open_calc_document(path: str) -> Tuple[Any, Any, Sheet]:
    """Open a Calc document and return (desktop, document, first_sheet).

    This is a minimal example that relies on the bundled LibreOffice Python.
    It raises RuntimeError if UNO is not available.
    """
    global current_desktop
    try:
        import uno
        import unohelper
        from com.sun.star.beans import PropertyValue
    except ImportError as exc:  # pragma: no cover - depends on LibreOffice runtime
        raise RuntimeError("UNO runtime is not available") from exc

    ctx = None
    boot_exc = None
    try:
        # bootstrap spins up a LibreOffice instance and returns a live component context
        ctx = unohelper.Bootstrap.bootstrap()
    except Exception as exc:  # pragma: no cover - defensive guard
        boot_exc = exc

    if ctx is None:
        try:
            # fallback: connect to an already-running soffice accepting sockets
            local_ctx = uno.getComponentContext()
            resolver = local_ctx.ServiceManager.createInstanceWithContext(
                "com.sun.star.bridge.UnoUrlResolver", local_ctx)
            ctx = resolver.resolve(
                "uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")
        except Exception:
            pass

    if ctx is None:  # pragma: no cover - depends on runtime
        raise RuntimeError("Failed to bootstrap or connect to LibreOffice UNO") from boot_exc

    smgr = ctx.getServiceManager()
    desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)

    url = uno.systemPathToFileUrl(path)
    properties = (PropertyValue("Hidden", 0, True, 0),)
    document = desktop.loadComponentFromURL(url, "_blank", 0, properties)

    doc_wrapper = CalcDocument(document)
    spreadsheet_doc = doc_wrapper.iface(InterfaceNames.X_SPREADSHEET_DOCUMENT)
    sheets = spreadsheet_doc.getSheets()
    first_sheet = sheets.getByIndex(0)
    current_desktop = desktop
    return desktop, doc_wrapper, Sheet(first_sheet, document=doc_wrapper)

def connect_calc() -> Tuple[Any, CalcDocument, Sheet]:
    global current_desktop
    try:
        import uno
        from com.sun.star.beans import PropertyValue
    except ImportError as exc:  # pragma: no cover - depends on LibreOffice runtime
        raise RuntimeError("UNO runtime is not available") from exc

    try:
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
        if doc is None:
            raise RuntimeError("No active Calc document found")

        doc_wrapper = CalcDocument(doc)
        spreadsheet_doc = doc_wrapper.iface(InterfaceNames.X_SPREADSHEET_DOCUMENT)
        controller = spreadsheet_doc.getCurrentController()
        sheet = controller.getActiveSheet()
        current_desktop = desktop
        return desktop, doc_wrapper, Sheet(sheet, document=doc_wrapper)
    except Exception as exc:  # pragma: no cover - depends on LibreOffice runtime
        raise RuntimeError("Failed to connect to Calc") from exc

# XSCRIPTCONTEXT に接続する
def connect_calc_script(xscriptcontext) -> Tuple[Any, CalcDocument, Sheet]:
    global current_desktop
    desktop = xscriptcontext.getDesktop()
    doc = CalcDocument(desktop.getCurrentComponent())
    controller = doc.raw.getCurrentController()
    sheet = Sheet(controller.getActiveSheet())
    current_desktop = desktop
    return desktop, doc, sheet


def connect_writer() -> Tuple[Any, WriterDocument]:
    """Connect to an active Writer document.

    Returns:
        (desktop, document_wrapper)

    Raises:
        RuntimeError: when UNO runtime is unavailable or no Writer document is active.
    """

    try:
        import uno
    except ImportError as exc:  # pragma: no cover - depends on LibreOffice runtime
        raise RuntimeError("UNO runtime is not available") from exc

    try:
        local_ctx = uno.getComponentContext()
        resolver = local_ctx.ServiceManager.createInstanceWithContext(
            "com.sun.star.bridge.UnoUrlResolver", local_ctx)

        ctx = resolver.resolve(
            "uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")

        smgr = ctx.ServiceManager
        desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)

        doc = desktop.getCurrentComponent()
        if doc is None or not doc.supportsService("com.sun.star.text.TextDocument"):
            raise RuntimeError("No active Writer document found")

        # Wrap the Writer document for convenience while still allowing raw access via .raw
        return desktop, WriterDocument(doc)
    except Exception as exc:  # pragma: no cover - depends on runtime
        raise RuntimeError("Failed to connect to Writer") from exc


def wrap_sheet(sheet_obj: Any) -> Sheet:
    """Wrap an existing UNO sheet object in a Sheet helper."""
    return Sheet(sheet_obj)
