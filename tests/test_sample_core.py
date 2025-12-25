from excellikeuno import Cell, InterfaceNames, Sheet, UnoObject


class DummyPropertySet:
    def __init__(self) -> None:
        self._props = {}

    def getPropertyValue(self, name: str):
        return self._props.get(name)

    def setPropertyValue(self, name: str, value):
        self._props[name] = value


class DummyCell:
    def __init__(self) -> None:
        self._value = 0.0
        self._formula = ""
        self._props = DummyPropertySet()

    def getValue(self) -> float:
        return self._value

    def setValue(self, value: float) -> None:
        self._value = value

    def getFormula(self) -> str:
        return self._formula

    def setFormula(self, formula: str) -> None:
        self._formula = formula

    def queryInterface(self, name: str):
        if name == InterfaceNames.X_CELL:
            return self
        if name == InterfaceNames.X_PROPERTY_SET:
            return self._props
        raise AttributeError(f"Unknown interface: {name}")


class DummySheet:
    def __init__(self) -> None:
        self._cells = {}

    def getCellByPosition(self, column: int, row: int):
        key = (column, row)
        if key not in self._cells:
            self._cells[key] = DummyCell()
        return self._cells[key]

    def queryInterface(self, name: str):
        if name == InterfaceNames.X_SPREADSHEET:
            return self
        raise AttributeError(f"Unknown interface: {name}")


def test_iface_cache_returns_same_instance():
    class Dummy:
        def __init__(self) -> None:
            self.calls = 0

        def queryInterface(self, name: str):
            self.calls += 1
            return name

    dummy = Dummy()
    wrapper = UnoObject(dummy)

    first = wrapper.iface("iface.example")
    second = wrapper.iface("iface.example")

    assert first == "iface.example"
    assert second == "iface.example"
    assert dummy.calls == 1


def test_cell_value_and_formula_roundtrip():
    raw_cell = DummyCell()
    cell = Cell(raw_cell)

    cell.value = 10.5
    cell.formula = "=A1+B1"
    cell.props.setPropertyValue("Note", "hello")

    assert raw_cell.getValue() == 10.5
    assert raw_cell.getFormula() == "=A1+B1"
    assert cell.value == 10.5
    assert cell.formula == "=A1+B1"
    assert cell.props.getPropertyValue("Note") == "hello"


def test_sheet_cell_wraps_getCellByPosition():
    raw_sheet = DummySheet()
    sheet = Sheet(raw_sheet)

    c = sheet.cell(1, 2)
    c.value = 3.14

    same = sheet.cell(1, 2)
    assert isinstance(c, Cell)
    assert same.value == 3.14
