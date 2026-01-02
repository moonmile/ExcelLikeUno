from __future__ import annotations

from typing import List, cast

from ..table import Sheet
from ..typing import InterfaceNames, XSpreadsheet, XSpreadsheetDocument
from .base import UnoObject


class CalcDocument(UnoObject):
    """Wraps a Calc XSpreadsheetDocument."""

    def _sheets(self):
        doc = cast(XSpreadsheetDocument, self.iface(InterfaceNames.X_SPREADSHEET_DOCUMENT))
        return doc.getSheets()

    def sheet(self, index: int) -> Sheet:
        sheets = self._sheets()
        sheet_obj = cast(XSpreadsheet, sheets.getByIndex(index))
        return Sheet(sheet_obj, document=self)

    def sheet_by_name(self, name: str) -> Sheet:
        sheets = self._sheets()
        sheet_obj = cast(XSpreadsheet, sheets.getByName(name))
        return Sheet(sheet_obj, document=self)

    def createInstance(self, service: str):
        """Create a UNO service using the document or component context.

        Calc documents normally implement XMultiServiceFactory, so delegate to
        the wrapped object's ``createInstance`` when present. If the document
        does not expose that surface (e.g., during tests with simplified mocks),
        fall back to the global component context service manager.
        """

        creator = getattr(self.raw, "createInstance", None)
        if callable(creator):
            try:
                return creator(service)
            except Exception:
                # fall back to component context if the document refuses
                pass

        try:
            import uno  # type: ignore

            ctx = uno.getComponentContext()
            smgr = ctx.getServiceManager()
            return smgr.createInstanceWithContext(service, ctx)
        except Exception as exc:  # pragma: no cover - depends on runtime
            raise AttributeError("Document cannot create UNO instance") from exc

    def add_sheet(self, name: str, index: int | None = None) -> Sheet:
        sheets = self._sheets()
        position = sheets.getCount() if index is None else int(index)
        sheets.insertNewByName(name, position)
        return self.sheet_by_name(name)

    def remove_sheet(self, name: str) -> None:
        sheets = self._sheets()
        sheets.removeByName(name)

    @property
    def sheet_names(self) -> List[str]:
        sheets = self._sheets()
        return list(sheets.getElementNames())

    @property
    def active_sheet(self) -> Sheet:
        doc = cast(XSpreadsheetDocument, self.iface(InterfaceNames.X_SPREADSHEET_DOCUMENT))
        controller = doc.getCurrentController()
        sheet_obj = cast(XSpreadsheet, controller.getActiveSheet())
        return Sheet(sheet_obj, document=self)
