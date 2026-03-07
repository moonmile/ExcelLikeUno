from __future__ import annotations

from typing import Any, Iterable, List, cast, TYPE_CHECKING

from ..core import UnoObject
from ..typing import InterfaceNames
from ..typing.interfaces import StructNames
from ..typing.calc import (
    DataPilotFieldOrientation,
    XDataPilotDescriptor,
    XDataPilotTable2,
    XDataPilotTables,
    XNamed,
    XPropertySet,
)
from ..typing.structs import CellAddress, CellRangeAddress
from .range import Range

if TYPE_CHECKING:  # pragma: no cover - typing only
    from ..sheet import Spreadsheet


def _coerce_range_address(target: Any) -> Any:
    candidate = target
    if isinstance(target, Range):
        try:
            addrable = target.iface(InterfaceNames.X_CELL_RANGE_ADDRESSABLE)
            candidate = addrable.getRangeAddress()
        except Exception:
            candidate = target
    elif hasattr(target, "to_raw"):
        try:
            candidate = target.to_raw()
        except Exception:
            candidate = target
    elif isinstance(target, CellRangeAddress):
        candidate = target.to_raw()

    try:
        import uno  # type: ignore

        struct = uno.createUnoStruct(StructNames.CELL_RANGE_ADDRESS)
    except Exception:
        struct = None

    if struct is None:
        return candidate

    for name in ("Sheet", "StartColumn", "StartRow", "EndColumn", "EndRow"):
        try:
            setattr(struct, name, getattr(candidate, name))
        except Exception:
            try:
                setattr(struct, name, candidate.get(name))  # type: ignore[arg-type]
            except Exception:
                pass
    return struct


def _coerce_cell_address(target: Any, default_sheet: int = 0) -> Any:
    if isinstance(target, CellAddress):
        return target.to_raw()
    if isinstance(target, Range):
        try:
            addrable = target.iface(InterfaceNames.X_CELL_RANGE_ADDRESSABLE)
            raddr = addrable.getRangeAddress()
            return CellAddress(Sheet=raddr.Sheet, Column=raddr.StartColumn, Row=raddr.StartRow).to_raw()
        except Exception:
            pass
    try:
        import uno  # type: ignore

        addr = uno.createUnoStruct(StructNames.CELL_ADDRESS)
    except Exception:
        addr = None
    if addr is None:
        addr = type("CellAddress", (), {})()
    try:
        col, row = target
    except Exception:
        col, row = 0, 0
    addr.Sheet = getattr(target, "Sheet", default_sheet)
    addr.Column = getattr(target, "Column", col)
    addr.Row = getattr(target, "Row", row)
    return addr


class PivotTable(UnoObject):
    def __init__(self, table_obj: Any) -> None:
        super().__init__(table_obj)

    @property
    def name(self) -> str:
        try:
            named = cast(XNamed, self.iface(InterfaceNames.X_NAMED))
            return named.getName()
        except Exception:
            try:
                return cast(str, getattr(self.raw, "Name", ""))
            except Exception:
                return ""

    @name.setter
    def name(self, value: str) -> None:
        target = str(value)
        try:
            named = cast(XNamed, self.iface(InterfaceNames.X_NAMED))
            named.setName(target)
            return
        except Exception:
            pass
        try:
            setattr(self.raw, "Name", target)
        except Exception:
            pass

    def refresh(self) -> None:
        table2 = None
        try:
            table2 = cast(XDataPilotTable2, self.iface(InterfaceNames.X_DATA_PILOT_TABLE2))
        except Exception:
            table2 = None
        if table2 is not None:
            try:
                table2.refresh()
                return
            except Exception:
                pass
        raw_refresh = getattr(self.raw, "refresh", None)
        if callable(raw_refresh):
            raw_refresh()

    @property
    def descriptor(self) -> Any | None:
        try:
            table2 = cast(XDataPilotTable2, self.iface(InterfaceNames.X_DATA_PILOT_TABLE2))
            return table2.getDataPilotDescriptor()
        except Exception:
            pass
        return None

    @property
    def source_range(self) -> Any:
        desc = self.descriptor
        if desc is not None:
            try:
                return desc.getSourceRange()
            except Exception:
                pass
        return None

    @property
    def output_range(self) -> Any:
        table2 = None
        try:
            table2 = cast(XDataPilotTable2, self.iface(InterfaceNames.X_DATA_PILOT_TABLE2))
        except Exception:
            table2 = None
        if table2 is not None:
            try:
                return table2.getOutputRange()
            except Exception:
                pass
        raw_get = getattr(self.raw, "getOutputRange", None)
        if callable(raw_get):
            try:
                return raw_get()
            except Exception:
                pass
        return None


class PivotTables:
    def __init__(self, sheet: "Spreadsheet") -> None:
        self.sheet = sheet

    def _tables(self) -> XDataPilotTables:
        raw = self.sheet.raw
        getter = getattr(raw, "getDataPilotTables", None)
        if callable(getter):
            tables = getter()
            return cast(XDataPilotTables, tables)
        raise RuntimeError("Sheet does not expose DataPilot tables")

    def _names(self) -> List[str]:
        tables = self._tables()
        try:
            names = tables.getElementNames()
            return list(names)
        except Exception:
            pass
        try:
            count = tables.getCount()
            return [cast(str, tables.getByName(tables.getElementNames()[i])) for i in range(count)]
        except Exception:
            return []

    def exists(self, name: str) -> bool:
        tables = self._tables()
        try:
            return bool(tables.hasByName(name))
        except Exception:
            return name in self._names()

    def __len__(self) -> int:
        try:
            return self._tables().getCount()
        except Exception:
            return len(self._names())

    def __iter__(self) -> Iterable[PivotTable]:
        return (self[name] for name in self._names())

    def __getitem__(self, key: str | int) -> PivotTable:
        tables = self._tables()
        if isinstance(key, int):
            names = self._names()
            idx = key if key >= 0 else len(names) + key
            if idx < 0 or idx >= len(names):
                raise IndexError("pivot table index out of range")
            key = names[idx]
        try:
            raw = tables.getByName(key)
        except Exception as exc:
            raise KeyError(f"pivot table not found: {key}") from exc
        return PivotTable(raw)

    def add(
        self,
        name: str,
        source_range: Any,
        target_cell: Any,
        row_fields: list[str] | None = None,
        column_fields: list[str] | None = None,
        data_fields: list[str] | None = None,
        filter_fields: list[str] | None = None,
    ) -> PivotTable:
        tables = self._tables()

        try:
            if hasattr(tables, "hasByName") and tables.hasByName(name):
                tables.removeByName(name)
        except Exception:
            pass

        target_name = str(name)
        doc = getattr(self.sheet, "document", None)

        def _has_name_on_any_sheet(candidate: str) -> bool:
            if doc is None:
                return False
            try:
                sheet_names = doc.sheet_names
            except Exception:
                return False
            for sname in sheet_names:
                try:
                    sh = doc.sheet_by_name(sname)
                    col = PivotTables(sh)
                    if candidate in col._names():
                        return True
                except Exception:
                    continue
            return False

        if _has_name_on_any_sheet(target_name):
            base = f"{target_name}_{getattr(self.sheet, 'name', 'sheet')}"
            candidate = base
            suffix = 1
            while _has_name_on_any_sheet(candidate):
                suffix += 1
                candidate = f"{base}_{suffix}"
            target_name = candidate

        source_addr = _coerce_range_address(source_range)
        if hasattr(source_addr, "Sheet"):
            default_sheet = getattr(source_addr, "Sheet", 0)
        else:
            default_sheet = 0
        output_addr = _coerce_cell_address(target_cell, default_sheet=default_sheet)

        descriptor: XDataPilotDescriptor | Any
        try:
            descriptor = tables.createDataPilotDescriptor()
        except Exception as exc:
            raise RuntimeError(f"failed to create DataPilot descriptor: {exc}")

        try:
            descriptor.setSourceRange(source_addr)
        except Exception:
            pass

        def _set_fields(names: list[str] | None, orientation: DataPilotFieldOrientation) -> None:
            if not names:
                return
            try:
                fields = descriptor.getDataPilotFields()
            except Exception:
                return
            for fname in names:
                try:
                    fld = fields.getByName(fname)
                except Exception:
                    continue
                try:
                    fld.setOrientation(int(orientation))
                except Exception:
                    try:
                        fld.Orientation = int(orientation)  # type: ignore[attr-defined]
                    except Exception:
                        pass

        _set_fields(row_fields, DataPilotFieldOrientation.ROW)
        _set_fields(column_fields, DataPilotFieldOrientation.COLUMN)
        _set_fields(filter_fields, DataPilotFieldOrientation.PAGE)

        if data_fields:
            try:
                fields = descriptor.getDataPilotFields()
            except Exception:
                fields = None
            if fields is not None:
                for fname in data_fields:
                    try:
                        fld = fields.getByName(fname)
                    except Exception:
                        continue
                    try:
                        fld.setOrientation(int(DataPilotFieldOrientation.DATA))
                    except Exception:
                        try:
                            fld.Orientation = int(DataPilotFieldOrientation.DATA)  # type: ignore[attr-defined]
                        except Exception:
                            pass
                    try:
                        props = cast(XPropertySet, fld)
                        props.setPropertyValue("Function", 2)  # SUM
                    except Exception:
                        pass

        try:
            tables.insertNewByName(str(target_name), output_addr, descriptor)
        except Exception as exc:
            raise RuntimeError(f"failed to insert pivot table '{target_name}': {exc}")

        pivot = self[target_name]
        try:
            current_name = pivot.name
        except Exception:
            current_name = None
        if not current_name or str(current_name) != str(target_name):
            try:
                pivot.name = target_name
            except Exception:
                pass

        return pivot
