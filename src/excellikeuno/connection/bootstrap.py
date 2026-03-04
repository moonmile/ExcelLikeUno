from __future__ import annotations

from typing import Any, Tuple

from ..core.calc_document import CalcDocument
from ..core.writer_document import WriterDocument
from ..typing import InterfaceNames
from ..table import Sheet

current_desktop: Any = None


class _LazyAlias:
    """Resolves the target callable on each access to mimic a live alias."""

    def __init__(self, resolver: Any) -> None:
        self._resolver = resolver

    def __call__(self, *args: Any, **kwargs: Any) -> Any:
        return self._resolver(*args, **kwargs)

    def __getattr__(self, name: str) -> Any:  # pragma: no cover - simple delegation
        target = self._resolver()
        return getattr(target, name)

    def __repr__(self) -> str:  # pragma: no cover - debug aid
        try:
            target = self._resolver()
            return f"<_LazyAlias to {target!r}>"
        except Exception:
            return "<_LazyAlias unresolved>"


def _make_properties(options: dict[str, Any]) -> tuple[Any, ...]:
    from com.sun.star.beans import PropertyValue  # type: ignore

    return tuple(PropertyValue(name, 0, value, 0) for name, value in options.items() if value is not None)


def _bootstrap_desktop() -> Any:
    try:
        import uno  # type: ignore
        import unohelper  # type: ignore
    except ImportError as exc:  # pragma: no cover - depends on LibreOffice runtime
        raise RuntimeError("UNO runtime is not available") from exc

    ctx = None
    boot_exc = None
    try:
        ctx = unohelper.Bootstrap.bootstrap()
    except Exception as exc:  # pragma: no cover - defensive guard
        boot_exc = exc

    if ctx is None:
        try:
            local_ctx = uno.getComponentContext()
            resolver = local_ctx.ServiceManager.createInstanceWithContext(
                "com.sun.star.bridge.UnoUrlResolver", local_ctx
            )
            ctx = resolver.resolve("uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")
        except Exception:
            pass

    if ctx is None:  # pragma: no cover - depends on runtime
        raise RuntimeError("Failed to bootstrap or connect to LibreOffice UNO") from boot_exc

    smgr = ctx.getServiceManager()
    desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
    return desktop


def this_desktop() -> Any:
    """Alias that returns the current desktop (connects if needed)."""

    return get_desktop()


def get_desktop() -> Any:
    """Connect to a running LibreOffice desktop via UNO socket."""

    global current_desktop
    if current_desktop is not None:
        return current_desktop

    try:
        import uno  # type: ignore
    except ImportError as exc:  # pragma: no cover - depends on LibreOffice runtime
        raise RuntimeError("UNO runtime is not available") from exc

    try:
        local_ctx = uno.getComponentContext()
        resolver = local_ctx.ServiceManager.createInstanceWithContext(
            "com.sun.star.bridge.UnoUrlResolver", local_ctx
        )
        ctx = resolver.resolve("uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")
        smgr = ctx.ServiceManager
        desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
        current_desktop = desktop
        return desktop
    except Exception as exc:  # pragma: no cover - depends on runtime
        raise RuntimeError("Failed to connect to LibreOffice desktop") from exc


def new_calc_document(hidden: bool = True) -> Tuple[Any, CalcDocument, Sheet]:
    desktop = _bootstrap_desktop()
    global current_desktop
    current_desktop = desktop

    try:
        import uno  # type: ignore
    except ImportError as exc:  # pragma: no cover - depends on runtime
        raise RuntimeError("UNO runtime is not available") from exc

    properties = _make_properties({"Hidden": hidden})
    document = desktop.loadComponentFromURL("private:factory/scalc", "_blank", 0, properties)

    doc_wrapper = CalcDocument(document)
    spreadsheet_doc = doc_wrapper.iface(InterfaceNames.X_SPREADSHEET_DOCUMENT)
    sheets = spreadsheet_doc.getSheets()
    first_sheet = sheets.getByIndex(0)
    return desktop, doc_wrapper, Sheet(first_sheet, document=doc_wrapper)


def add_calc_document(hidden: bool = True) -> Tuple[Any, CalcDocument, Sheet]:
    """Alias for new_calc_document to match desktop-oriented naming."""

    return new_calc_document(hidden=hidden)


def open_calc_document(
    path: str,
    *,
    hidden: bool = True,
    read_only: bool = False,
    as_template: bool = False,
    filter_name: str | None = None,
) -> Tuple[Any, CalcDocument, Sheet]:
    """Open a Calc document and return (desktop, document_wrapper, first_sheet)."""

    desktop = _bootstrap_desktop()
    global current_desktop
    current_desktop = desktop

    try:
        import uno  # type: ignore
    except ImportError as exc:  # pragma: no cover - depends on runtime
        raise RuntimeError("UNO runtime is not available") from exc

    url = uno.systemPathToFileUrl(path)
    properties = _make_properties(
        {"Hidden": hidden, "ReadOnly": read_only, "AsTemplate": as_template, "FilterName": filter_name}
    )
    document = desktop.loadComponentFromURL(url, "_blank", 0, properties)

    doc_wrapper = CalcDocument(document)
    spreadsheet_doc = doc_wrapper.iface(InterfaceNames.X_SPREADSHEET_DOCUMENT)
    sheets = spreadsheet_doc.getSheets()
    first_sheet = sheets.getByIndex(0)
    return desktop, doc_wrapper, Sheet(first_sheet, document=doc_wrapper)


def connect_calc() -> Tuple[Any, CalcDocument, Sheet]:
    desktop = get_desktop()
    doc = desktop.getCurrentComponent()
    if doc is None:
        raise RuntimeError("No active Calc document found")

    doc_wrapper = CalcDocument(doc)
    spreadsheet_doc = doc_wrapper.iface(InterfaceNames.X_SPREADSHEET_DOCUMENT)
    controller = spreadsheet_doc.getCurrentController()
    sheet = controller.getActiveSheet()
    return desktop, doc_wrapper, Sheet(sheet, document=doc_wrapper)


def active_document() -> CalcDocument:
    desktop = get_desktop()
    doc = desktop.getCurrentComponent()
    if doc is None:
        raise RuntimeError("No active Calc document found")
    return CalcDocument(doc)


def get_active_calc_document() -> CalcDocument:
    """Alias for active_document to mirror plan naming."""

    return active_document()


def this_document() -> CalcDocument:
    return active_document()


# XSCRIPTCONTEXT に接続する
def connect_calc_script(xscriptcontext) -> Tuple[Any, CalcDocument, Sheet]:
    global current_desktop
    desktop = xscriptcontext.getDesktop()
    doc = CalcDocument(desktop.getCurrentComponent())
    controller = doc.raw.getCurrentController()
    sheet = Sheet(controller.getActiveSheet(), document=doc)
    current_desktop = desktop
    return desktop, doc, sheet


def connect_writer() -> Tuple[Any, WriterDocument]:
    """Connect to an active Writer document.

    Returns:
        (desktop, document_wrapper)

    Raises:
        RuntimeError: when UNO runtime is unavailable or no Writer document is active.
    """

    desktop = get_desktop()
    doc = desktop.getCurrentComponent()
    if doc is None or not doc.supportsService("com.sun.star.text.TextDocument"):
        raise RuntimeError("No active Writer document found")

    return desktop, WriterDocument(doc)


def active_sheet() -> Sheet:
    doc = active_document()
    controller = doc.raw.getCurrentController()
    sheet = controller.getActiveSheet()
    return Sheet(sheet, document=doc)


# Aliases for naming convenience
ActiveCalcDocument = _LazyAlias(get_active_calc_document)
ActiveSheet = _LazyAlias(active_sheet)
ThisDesktop = _LazyAlias(this_desktop)


def this_sheet() -> Sheet:
    return active_sheet()


def wrap_sheet(sheet_obj: Any) -> Sheet:
    """Wrap an existing UNO sheet object in a Sheet helper."""
    return Sheet(sheet_obj)
