import pytest


def _connect_or_skip():
    try:
        from excellikeuno.connection import connect_calc
    except Exception:
        pytest.skip("connect_calc not available")
    try:
        return connect_calc()
    except Exception as exc:
        pytest.skip(f"UNO runtime not available: {exc}")


def test_range_value_matrix_roundtrip():
    _, __, sheet = _connect_or_skip()
    rng = sheet.range(0, 8, 1, 9)  # 2x2 block

    coords = [(0, 0), (1, 0), (0, 1), (1, 1)]
    originals = [rng.cell(c, r).formula for c, r in coords]

    matrix = [[1.25, "alpha"], [2, "beta"]]

    try:
        rng.value = matrix

        assert abs(rng.cell(0, 0).value - 1.25) < 0.001
        assert rng.cell(1, 0).text == "alpha"
        assert abs(rng.cell(0, 1).value - 2.0) < 0.001
        assert rng.cell(1, 1).text == "beta"

        values = rng.value
        assert values == [["1.25", "alpha"], ["2", "beta"]]
    finally:
        for (c, r), formula in zip(coords, originals):
            try:
                rng.cell(c, r).formula = formula
            except Exception:
                pass


def test_range_value_row_roundtrip():
    _, __, sheet = _connect_or_skip()
    rng = sheet.range(0, 10, 1, 10)  # 1x2 block

    coords = [(0, 0), (1, 0)]
    originals = [rng.cell(c, r).formula for c, r in coords]

    try:
        rng.value = [10, "zeta"]

        assert abs(rng.cell(0, 0).value - 10.0) < 0.001
        assert rng.cell(1, 0).text == "zeta"

        assert rng.value == [["10", "zeta"]]
    finally:
        for (c, r), formula in zip(coords, originals):
            try:
                rng.cell(c, r).formula = formula
            except Exception:
                pass
