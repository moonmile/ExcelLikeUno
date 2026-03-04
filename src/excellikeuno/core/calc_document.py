from __future__ import annotations

from typing import Any, List, cast

from ..table import Sheet
from ..typing import InterfaceNames, XSpreadsheet, XSpreadsheetDocument, XStorable
from .base import UnoObject


class CalcDocument(UnoObject):
    """Wraps a Calc XSpreadsheetDocument."""

    @property
    def storable(self) -> XStorable:
        return cast(XStorable, self.iface(InterfaceNames.X_STORABLE))

    @staticmethod
    def _to_url(path: str) -> str:
        import uno  # type: ignore

        return uno.systemPathToFileUrl(path)

    @staticmethod
    def _make_properties(options: dict[str, Any]) -> tuple[Any, ...]:
        from com.sun.star.beans import PropertyValue  # type: ignore

        return tuple(
            PropertyValue(name, 0, value, 0) for name, value in options.items() if value is not None
        )

    def save(self) -> None:
        self.storable.store()

    def save_as(self, path: str, filter_name: str | None = None, overwrite: bool = True) -> None:
        options: dict[str, Any] = {"FilterName": filter_name, "Overwrite": overwrite}
        self.storable.storeAsURL(self._to_url(path), self._make_properties(options))

    def save_copy(self, path: str, filter_name: str | None = None, overwrite: bool = True) -> None:
        options: dict[str, Any] = {"FilterName": filter_name, "Overwrite": overwrite}
        self.storable.storeToURL(self._to_url(path), self._make_properties(options))

    @property
    def is_modified(self) -> bool:
        return bool(self.storable.isModified())

    @property
    def has_location(self) -> bool:
        return bool(self.storable.hasLocation())

    @property
    def title(self) -> str:
        getter = getattr(self.raw, "getTitle", None)
        if not callable(getter):
            raise AttributeError("Underlying document does not expose getTitle")
        return cast(str, getter())

    @property
    def url(self) -> str:
        getter = getattr(self.raw, "getURL", None)
        if not callable(getter):
            raise AttributeError("Underlying document does not expose getURL")
        return cast(str, getter())

    def activate(self) -> None:
        """Bring this document to the foreground when the runtime supports it."""

        controller = getattr(self.raw, "getCurrentController", lambda: None)()
        frame = getattr(controller, "getFrame", lambda: None)()
        activate = getattr(frame, "activate", None)
        if not callable(activate):
            raise AttributeError("Document frame does not support activate")
        activate()

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

    def add_sheet(self, name: str, index: int | None = None, from_sheet_name: str | None = None) -> Sheet:
        """Add a new sheet.

        When ``from_sheet_name`` is provided and the UNO runtime supports ``copyByName``,
        duplicate the source sheet into the requested position; otherwise insert a blank sheet.
        """

        sheets = self._sheets()
        position = sheets.getCount() if index is None else int(index)

        if from_sheet_name:
            copier = getattr(sheets, "copyByName", None)
            if callable(copier):
                copier(from_sheet_name, name, position)
                return self.sheet_by_name(name)

        sheets.insertNewByName(name, position)
        return self.sheet_by_name(name)

    def remove_sheet(self, name: str) -> None:
        sheets = self._sheets()
        sheets.removeByName(name)

    def copy_sheet(self, source_name: str, new_name: str, index: int | None = None) -> Sheet:
        """Copy an existing sheet to a new sheet name.

        Uses UNO ``copyByName`` when available; raises AttributeError if the interface is missing.
        """

        sheets = self._sheets()
        position = sheets.getCount() if index is None else int(index)
        copier = getattr(sheets, "copyByName", None)
        if not callable(copier):
            raise AttributeError("Underlying sheets object does not support copyByName")
        copier(source_name, new_name, position)
        return self.sheet_by_name(new_name)

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

    @property
    def this_sheet(self) -> Sheet:
        return self.active_sheet


    # XModel interface methods
    @property
    def title(self) -> str:
        return self.raw.getTitle()
    @property
    def url(self) -> str:
        return self.raw.getURL()
