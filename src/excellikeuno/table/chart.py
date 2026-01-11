from __future__ import annotations

from typing import Any, Iterable, List, cast, TYPE_CHECKING

from ..core import UnoObject
from ..typing import InterfaceNames
from ..typing.interfaces import StructNames
from ..typing.calc import XNamed, XPropertySet, XTableChart, XTableCharts
from ..typing.structs import CellRangeAddress, Point, Rectangle, Size
from .range import Range

if TYPE_CHECKING:  # pragma: no cover - only for type hints
    from ..core.calc_document import CalcDocument
    from .sheet import Sheet


class Chart(UnoObject):
    """Wraps a Calc table chart anchored on a sheet."""

    def __init__(self, chart_obj: Any, document: "CalcDocument | None" = None) -> None:
        super().__init__(chart_obj)
        self._document = document
        self._last_range: Any | None = None
        self._last_position: Any | None = None
        self._last_size: Any | None = None
        self._last_name: str | None = None
        self._last_column_headers: bool | None = None
        self._last_row_headers: bool | None = None
        self._last_has_legend: bool | None = None
        self._last_chart_type: str | None = None

    @property
    def props(self) -> XPropertySet:
        return cast(XPropertySet, self.iface(InterfaceNames.X_PROPERTY_SET))

    @property
    def name(self) -> str:
        try:
            named = cast(XNamed, self.iface(InterfaceNames.X_NAMED))
            return named.getName()
        except Exception:
            try:
                return cast(str, self.props.getPropertyValue("Name"))
            except Exception:
                if self._last_name is not None:
                    return self._last_name
        return ""

    @name.setter
    def name(self, value: str) -> None:
        target = str(value)
        try:
            named = cast(XNamed, self.iface(InterfaceNames.X_NAMED))
            named.setName(target)
            self._last_name = target
            return
        except Exception:
            pass
        try:
            self.props.setPropertyValue("Name", target)
            self._last_name = target
            return
        except Exception:
            self._last_name = target

    @staticmethod
    def _coerce_range_address(target: Any) -> Any:
        """Normalize range-like input to a UNO CellRangeAddress struct."""
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
            for name in ("Sheet", "StartColumn", "StartRow", "EndColumn", "EndRow"):
                try:
                    setattr(struct, name, getattr(candidate, name))
                except Exception:
                    try:
                        setattr(struct, name, candidate.get(name))  # type: ignore[arg-type]
                    except Exception:
                        pass
            return struct
        except Exception:
            return candidate

    @property
    def range(self) -> Any:
        chart_iface = None
        try:
            chart_iface = cast(XTableChart, self.iface(InterfaceNames.X_TABLE_CHART))
            return chart_iface.getRangeAddress()
        except Exception:
            pass
        for prop_name in ("RangeAddress", "Ranges"):
            try:
                return self.props.getPropertyValue(prop_name)
            except Exception:
                continue
        if chart_iface is None:
            raw_get = getattr(self.raw, "getRangeAddress", None)
            if callable(raw_get):
                try:
                    return raw_get()
                except Exception:
                    pass
        return self._last_range

    @range.setter
    def range(self, value: Any) -> None:
        addr = self._coerce_range_address(value)
        try:
            chart_iface = cast(XTableChart, self.iface(InterfaceNames.X_TABLE_CHART))
            chart_iface.setRangeAddress(addr)
            self._last_range = addr
            return
        except Exception:
            pass
        for prop_name in ("RangeAddress", "Ranges"):
            try:
                self.props.setPropertyValue(prop_name, addr)
                self._last_range = addr
                return
            except Exception:
                continue
        raw_set = getattr(self.raw, "setRangeAddress", None)
        if callable(raw_set):
            try:
                raw_set(addr)
                self._last_range = addr
                return
            except Exception:
                pass
        self._last_range = addr

    @property
    def position(self) -> Any:
        chart_iface = None
        try:
            chart_iface = cast(XTableChart, self.iface(InterfaceNames.X_TABLE_CHART))
            pos = chart_iface.getPosition()
            return Point(pos.X, pos.Y)
        except Exception:
            pass
        for prop_name in ("Position", "RectPosition"):
            try:
                pos = self.props.getPropertyValue(prop_name)
                return pos
            except Exception:
                continue
        if chart_iface is None:
            raw_get = getattr(self.raw, "getPosition", None)
            if callable(raw_get):
                try:
                    return raw_get()
                except Exception:
                    pass
        return self._last_position

    @position.setter
    def position(self, value: Any) -> None:
        target = value
        if hasattr(value, "to_raw"):
            try:
                target = value.to_raw()
            except Exception:
                target = value
        try:
            chart_iface = cast(XTableChart, self.iface(InterfaceNames.X_TABLE_CHART))
            chart_iface.setPosition(target)
            self._last_position = target
            return
        except Exception:
            pass
        for prop_name in ("Position", "RectPosition"):
            try:
                self.props.setPropertyValue(prop_name, target)
                self._last_position = target
                return
            except Exception:
                continue
        raw_set = getattr(self.raw, "setPosition", None)
        if callable(raw_set):
            try:
                raw_set(target)
                self._last_position = target
                return
            except Exception:
                pass
        self._last_position = target

    @property
    def size(self) -> Any:
        chart_iface = None
        try:
            chart_iface = cast(XTableChart, self.iface(InterfaceNames.X_TABLE_CHART))
            size = chart_iface.getSize()
            return Size(size.Width, size.Height)
        except Exception:
            pass
        for prop_name in ("Size", "RectSize"):
            try:
                return self.props.getPropertyValue(prop_name)
            except Exception:
                continue
        if chart_iface is None:
            raw_get = getattr(self.raw, "getSize", None)
            if callable(raw_get):
                try:
                    return raw_get()
                except Exception:
                    pass
        return self._last_size

    @size.setter
    def size(self, value: Any) -> None:
        target = value
        if hasattr(value, "to_raw"):
            try:
                target = value.to_raw()
            except Exception:
                target = value
        try:
            chart_iface = cast(XTableChart, self.iface(InterfaceNames.X_TABLE_CHART))
            chart_iface.setSize(target)
            self._last_size = target
            return
        except Exception:
            pass
        for prop_name in ("Size", "RectSize"):
            try:
                self.props.setPropertyValue(prop_name, target)
                self._last_size = target
                return
            except Exception:
                continue
        raw_set = getattr(self.raw, "setSize", None)
        if callable(raw_set):
            try:
                raw_set(target)
                self._last_size = target
                return
            except Exception:
                pass
        self._last_size = target

    @property
    def embedded_document(self) -> Any:
        try:
            chart_iface = cast(XTableChart, self.iface(InterfaceNames.X_TABLE_CHART))
            return chart_iface.getEmbeddedObject()
        except Exception:
            raw_get = getattr(self.raw, "getEmbeddedObject", None)
            if callable(raw_get):
                try:
                    return raw_get()
                except Exception:
                    pass
        return None

    # Header flags
    @property
    def column_headers(self) -> bool | None:
        try:
            return bool(self.props.getPropertyValue("HasColumnHeaders"))
        except Exception:
            return self._last_column_headers

    @column_headers.setter
    def column_headers(self, value: bool) -> None:
        target = bool(value)
        try:
            self.props.setPropertyValue("HasColumnHeaders", target)
            self._last_column_headers = target
            return
        except Exception:
            self._last_column_headers = target

    @property
    def chart_type(self) -> str | None:
        diagram = None
        chart_doc = self.embedded_document
        if chart_doc is not None:
            try:
                diagram = chart_doc.getDiagram()
            except Exception:
                pass
        if diagram is None:
            try:
                diagram = self.props.getPropertyValue("Diagram")
            except Exception:
                diagram = None

        known_services = [
            "com.sun.star.chart.BarDiagram",
            "com.sun.star.chart.AreaDiagram",
            "com.sun.star.chart.BubbleDiagram",
            "com.sun.star.chart.Dim2DDiagram",
            "com.sun.star.chart.DonutDiagram",
            "com.sun.star.chart.FilledNetDiagram",
            "com.sun.star.chart.LineDiagram",
            "com.sun.star.chart.NetDiagram",
            "com.sun.star.chart.PieDiagram",
            "com.sun.star.chart.StackableDiagram",
            "com.sun.star.chart.StockDiagram",
            "com.sun.star.chart.XYDiagram",
        ]
        if diagram is not None:
            supports = getattr(diagram, "supportsService", None)
            if callable(supports):
                for svc in known_services:
                    try:
                        if supports(svc):
                            self._last_chart_type = svc
                            return svc
                    except Exception:
                        continue
            impl = getattr(diagram, "getImplementationName", None)
            if callable(impl):
                try:
                    name = impl()
                    if name:
                        self._last_chart_type = str(name)
                        return self._last_chart_type
                except Exception:
                    pass
        return self._last_chart_type

    @chart_type.setter
    def chart_type(self, service_name: str) -> None:
        target = str(service_name)
        chart_doc = self.embedded_document
        if chart_doc is None:
            self._last_chart_type = target
            return
        try:
            diagram = chart_doc.createInstance(target)
            chart_doc.setDiagram(diagram)
            self._last_chart_type = target
        except Exception:
            self._last_chart_type = target

    @property
    def row_headers(self) -> bool | None:
        try:
            return bool(self.props.getPropertyValue("HasRowHeaders"))
        except Exception:
            return self._last_row_headers

    @row_headers.setter
    def row_headers(self, value: bool) -> None:
        target = bool(value)
        try:
            self.props.setPropertyValue("HasRowHeaders", target)
            self._last_row_headers = target
            return
        except Exception:
            self._last_row_headers = target

    @property
    def has_legend(self) -> bool | None:
        try:
            return bool(self.props.getPropertyValue("HasLegend"))
        except Exception:
            return self._last_has_legend

    @has_legend.setter
    def has_legend(self, value: bool) -> None:
        target = bool(value)
        try:
            self.props.setPropertyValue("HasLegend", target)
            self._last_has_legend = target
            return
        except Exception:
            self._last_has_legend = target


class ChartCollection:
    """Helper for managing charts on a sheet."""

    def __init__(self, sheet: "Sheet") -> None:
        self.sheet = sheet

    def _charts(self) -> XTableCharts:
        try:
            supplier = self.sheet.iface(InterfaceNames.X_TABLE_CHARTS_SUPPLIER)
            charts = supplier.getCharts()
            if charts is not None:
                return cast(XTableCharts, charts)
        except Exception:
            pass
        return cast(XTableCharts, self.sheet.raw.getCharts())

    def _names(self) -> List[str]:
        charts = self._charts()
        try:
            names = charts.getElementNames()
            names_list = list(names)
            if names_list:
                return names_list
        except Exception:
            pass
        try:
            count = charts.getCount()
            return [cast(str, charts.getByIndex(i).Name) for i in range(count)]  # type: ignore[attr-defined]
        except Exception:
            pass

        # Fallback: scan draw page for embedded chart shapes (cross-sheet robustness).
        names_dp: list[str] = []
        try:
            draw_page = self.sheet.iface(InterfaceNames.X_DRAW_PAGE_SUPPLIER).getDrawPage()
            get_count = getattr(draw_page, "getCount", None)
            get_by_index = getattr(draw_page, "getByIndex", None)
            if callable(get_count) and callable(get_by_index):
                for idx in range(int(get_count())):
                    try:
                        shape = get_by_index(idx)
                        name_val = None
                        try:
                            name_val = shape.getName() if hasattr(shape, "getName") else None
                        except Exception:
                            name_val = None
                        if name_val is None:
                            try:
                                name_val = getattr(shape, "Name", None)
                            except Exception:
                                name_val = None
                        if name_val:
                            names_dp.append(str(name_val))
                    except Exception:
                        continue
        except Exception:
            pass

        return names_dp

    def exists(self, name: str) -> bool:
        """Return True if a chart with the given name exists."""
        charts = self._charts()
        try:
            return bool(charts.hasByName(name))
        except Exception:
            return name in self._names()

    def __len__(self) -> int:
        try:
            return self._charts().getCount()
        except Exception:
            return len(self._names())

    def __iter__(self) -> Iterable[Chart]:
        return (self[name] for name in self._names())

    def __getitem__(self, key: str | int) -> Chart:
        charts = self._charts()
        if isinstance(key, int):
            names = self._names()
            idx = key if key >= 0 else len(names) + key
            if idx < 0 or idx >= len(names):
                raise IndexError("chart index out of range")
            key = names[idx]
        try:
            raw = charts.getByName(key)
        except Exception as exc:
            raise KeyError(f"chart not found: {key}") from exc
        return Chart(raw, document=self.sheet.document)

    @staticmethod
    def _coerce_rectangle(rect: Any | None, pos: Any | None, size: Any | None) -> Any:
        if rect is not None:
            if hasattr(rect, "to_raw"):
                try:
                    return rect.to_raw()
                except Exception:
                    return rect
            return rect
        if pos is None or size is None:
            # Default rectangle if nothing provided (roughly 10x7 cm)
            return Rectangle(1000, 1000, 10000, 7000).to_raw()
        if hasattr(pos, "to_raw"):
            try:
                pos = pos.to_raw()
            except Exception:
                pass
        if hasattr(size, "to_raw"):
            try:
                size = size.to_raw()
            except Exception:
                pass
        rect_obj = Rectangle()
        rect_obj.X = getattr(pos, "X", 0)
        rect_obj.Y = getattr(pos, "Y", 0)
        rect_obj.Width = getattr(size, "Width", 0)
        rect_obj.Height = getattr(size, "Height", 0)
        return rect_obj.to_raw()

    def add(
        self,
        name: str,
        data_range: Any,
        rectangle: Any | None = None,
        position: Any | None = None,
        size: Any | None = None,
        column_headers: bool = True,
        row_headers: bool = False,
        diagram_service: str | None = None,
    ) -> Chart:
        # Step 1: remove same-name chart on this sheet if present
        charts = self._charts()
        try:
            if charts.hasByName(name):
                charts.removeByName(name)
        except Exception:
            try:
                if name in self._names():
                    charts.removeByName(name)
            except Exception:
                pass

        # Step 2: avoid cross-sheet name collisions (UNO bug workaround)
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
                    col = ChartCollection(sh)
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

        rect_raw = self._coerce_rectangle(rectangle, position, size)
        if isinstance(data_range, Range):
            try:
                addr = data_range.raw.getRangeAddress()
            except Exception:
                addr = Chart._coerce_range_address(data_range)
        else:
            addr = Chart._coerce_range_address(data_range)
        strategies = []
        strategies.append(lambda: charts.addNewByName(target_name, rect_raw, (addr,), bool(column_headers), bool(row_headers)))
        try:
            import uno  # type: ignore

            addr_seq = uno.Any("[]com.sun.star.table.CellRangeAddress", (addr,))
            strategies.append(lambda: charts.addNewByName(target_name, rect_raw, addr_seq, bool(column_headers), bool(row_headers)))
            strategies.append(
                lambda: uno.invoke(
                    charts,
                    "addNewByName",
                    (target_name, rect_raw, (addr,), bool(column_headers), bool(row_headers)),
                )
            )
            strategies.append(
                lambda: uno.invoke(
                    charts,
                    "addNewByName",
                    (target_name, rect_raw, addr_seq, bool(column_headers), bool(row_headers)),
                )
            )
        except Exception:
            pass

        def _try_strategies() -> Exception | None:
            last_exc: Exception | None = None
            for call in strategies:
                try:
                    call()
                    return None
                except Exception as exc:  # pragma: no cover - depends on UNO runtime
                    last_exc = exc
            return last_exc

        last_exc = _try_strategies()
        if last_exc is not None:
            # If the name still exists, attempt one forced removal then retry once.
            try:
                if hasattr(charts, "hasByName") and charts.hasByName(target_name):
                    charts.removeByName(target_name)
                elif target_name in self._names():
                    charts.removeByName(target_name)
            except Exception:
                pass
            last_exc = _try_strategies()

        if last_exc is not None:
            exc_type = type(last_exc).__name__
            exc_msg = getattr(last_exc, "Message", "") or str(last_exc) or repr(last_exc)
            try:
                names = self._names()
            except Exception:
                names = []
            def _addr_fields(a: Any) -> str:
                try:
                    return f"Sheet={getattr(a, 'Sheet', '?')},Cols={getattr(a, 'StartColumn', '?')}-{getattr(a, 'EndColumn', '?')},Rows={getattr(a, 'StartRow', '?')}-{getattr(a, 'EndRow', '?')}"
                except Exception:
                    return repr(a)

            detail_addr = _addr_fields(addr)
            detail_rect = None
            try:
                detail_rect = f"Rect X={getattr(rect_raw, 'X', '?')},Y={getattr(rect_raw, 'Y', '?')},W={getattr(rect_raw, 'Width', '?')},H={getattr(rect_raw, 'Height', '?')}"
            except Exception:
                detail_rect = repr(rect_raw)

            raise RuntimeError(
                f"failed to add chart '{target_name}': {exc_type}: {exc_msg} | addr=({detail_addr}) rect=({detail_rect}) existing={names}"
            )

        chart = self[target_name]
        try:
            chart._last_range = addr  # type: ignore[attr-defined]
        except Exception:
            pass
        try:
            chart.range = addr
        except Exception:
            pass
        # Some environments leave the chart name empty after creation; enforce the requested name.
        try:
            current_name = chart.name
        except Exception:
            current_name = None
        if not current_name or str(current_name) != str(target_name):
            try:
                chart.name = target_name
            except Exception:
                pass
        if diagram_service:
            chart_doc = chart.embedded_document
            if chart_doc is not None:
                try:
                    diagram = chart_doc.createInstance(diagram_service)
                    chart_doc.setDiagram(diagram)
                except Exception:
                    pass
        return chart

    def add_bar_diagram(
        self,
        name: str,
        data_range: Any,
        rectangle: Any | None = None,
        position: Any | None = None,
        size: Any | None = None,
        column_headers: bool = True,
        row_headers: bool = False,
    ) -> Chart:
        return self.add(
            name=name,
            data_range=data_range,
            rectangle=rectangle,
            position=position,
            size=size,
            column_headers=column_headers,
            row_headers=row_headers,
            diagram_service="com.sun.star.chart.BarDiagram",
        )

    def add_area_chart(
        self,
        name: str,
        data_range: Any,
        rectangle: Any | None = None,
        position: Any | None = None,
        size: Any | None = None,
        column_headers: bool = True,
        row_headers: bool = False,
    ) -> Chart:
        return self.add(
            name=name,
            data_range=data_range,
            rectangle=rectangle,
            position=position,
            size=size,
            column_headers=column_headers,
            row_headers=row_headers,
            diagram_service="com.sun.star.chart.AreaDiagram",
        )

    def add_bubble_diagram(
        self,
        name: str,
        data_range: Any,
        rectangle: Any | None = None,
        position: Any | None = None,
        size: Any | None = None,
        column_headers: bool = True,
        row_headers: bool = False,
    ) -> Chart:
        return self.add(
            name=name,
            data_range=data_range,
            rectangle=rectangle,
            position=position,
            size=size,
            column_headers=column_headers,
            row_headers=row_headers,
            diagram_service="com.sun.star.chart.BubbleDiagram",
        )

    def add_dim2d_diagram(
        self,
        name: str,
        data_range: Any,
        rectangle: Any | None = None,
        position: Any | None = None,
        size: Any | None = None,
        column_headers: bool = True,
        row_headers: bool = False,
    ) -> Chart:
        return self.add(
            name=name,
            data_range=data_range,
            rectangle=rectangle,
            position=position,
            size=size,
            column_headers=column_headers,
            row_headers=row_headers,
            diagram_service="com.sun.star.chart.Dim2DDiagram",
        )

    def add_donut_diagram(
        self,
        name: str,
        data_range: Any,
        rectangle: Any | None = None,
        position: Any | None = None,
        size: Any | None = None,
        column_headers: bool = True,
        row_headers: bool = False,
    ) -> Chart:
        return self.add(
            name=name,
            data_range=data_range,
            rectangle=rectangle,
            position=position,
            size=size,
            column_headers=column_headers,
            row_headers=row_headers,
            diagram_service="com.sun.star.chart.DonutDiagram",
        )

    def add_filled_net_diagram(
        self,
        name: str,
        data_range: Any,
        rectangle: Any | None = None,
        position: Any | None = None,
        size: Any | None = None,
        column_headers: bool = True,
        row_headers: bool = False,
    ) -> Chart:
        return self.add(
            name=name,
            data_range=data_range,
            rectangle=rectangle,
            position=position,
            size=size,
            column_headers=column_headers,
            row_headers=row_headers,
            diagram_service="com.sun.star.chart.FilledNetDiagram",
        )

    def add_line_diagram(
        self,
        name: str,
        data_range: Any,
        rectangle: Any | None = None,
        position: Any | None = None,
        size: Any | None = None,
        column_headers: bool = True,
        row_headers: bool = False,
    ) -> Chart:
        return self.add(
            name=name,
            data_range=data_range,
            rectangle=rectangle,
            position=position,
            size=size,
            column_headers=column_headers,
            row_headers=row_headers,
            diagram_service="com.sun.star.chart.LineDiagram",
        )

    def add_net_diagram(
        self,
        name: str,
        data_range: Any,
        rectangle: Any | None = None,
        position: Any | None = None,
        size: Any | None = None,
        column_headers: bool = True,
        row_headers: bool = False,
    ) -> Chart:
        return self.add(
            name=name,
            data_range=data_range,
            rectangle=rectangle,
            position=position,
            size=size,
            column_headers=column_headers,
            row_headers=row_headers,
            diagram_service="com.sun.star.chart.NetDiagram",
        )

    def add_pie_diagram(
        self,
        name: str,
        data_range: Any,
        rectangle: Any | None = None,
        position: Any | None = None,
        size: Any | None = None,
        column_headers: bool = True,
        row_headers: bool = False,
    ) -> Chart:
        return self.add(
            name=name,
            data_range=data_range,
            rectangle=rectangle,
            position=position,
            size=size,
            column_headers=column_headers,
            row_headers=row_headers,
            diagram_service="com.sun.star.chart.PieDiagram",
        )

    def add_stackable_diagram(
        self,
        name: str,
        data_range: Any,
        rectangle: Any | None = None,
        position: Any | None = None,
        size: Any | None = None,
        column_headers: bool = True,
        row_headers: bool = False,
    ) -> Chart:
        return self.add(
            name=name,
            data_range=data_range,
            rectangle=rectangle,
            position=position,
            size=size,
            column_headers=column_headers,
            row_headers=row_headers,
            diagram_service="com.sun.star.chart.StackableDiagram",
        )

    def add_stock_diagram(
        self,
        name: str,
        data_range: Any,
        rectangle: Any | None = None,
        position: Any | None = None,
        size: Any | None = None,
        column_headers: bool = True,
        row_headers: bool = False,
    ) -> Chart:
        return self.add(
            name=name,
            data_range=data_range,
            rectangle=rectangle,
            position=position,
            size=size,
            column_headers=column_headers,
            row_headers=row_headers,
            diagram_service="com.sun.star.chart.StockDiagram",
        )

    def add_xy_diagram(
        self,
        name: str,
        data_range: Any,
        rectangle: Any | None = None,
        position: Any | None = None,
        size: Any | None = None,
        column_headers: bool = True,
        row_headers: bool = False,
    ) -> Chart:
        return self.add(
            name=name,
            data_range=data_range,
            rectangle=rectangle,
            position=position,
            size=size,
            column_headers=column_headers,
            row_headers=row_headers,
            diagram_service="com.sun.star.chart.XYDiagram",
        )

    def remove(self, name: str) -> None:
        charts = self._charts()
        try:
            try:
                if hasattr(charts, "hasByName") and not charts.hasByName(name):
                    return
            except Exception:
                if name not in self._names():
                    return

            try:
                charts.removeByName(name)
            except Exception as inner_exc:
                # If remove failed but the name is now absent, treat as success to stay idempotent.
                try:
                    if hasattr(charts, "hasByName") and not charts.hasByName(name):
                        return
                except Exception:
                    try:
                        if name not in self._names():
                            return
                    except Exception:
                        pass
                raise inner_exc

            # Verify removal when possible to catch silent failures, but ignore if capability is missing.
            try:
                if hasattr(charts, "hasByName") and charts.hasByName(name):
                    raise RuntimeError(f"chart '{name}' still present after removeByName")
            except Exception:
                try:
                    if name in self._names():
                        raise RuntimeError(f"chart '{name}' still present after removeByName")
                except Exception:
                    pass
        except Exception as exc:
            raise RuntimeError(f"failed to remove chart '{name}': {exc}")
