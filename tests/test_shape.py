import pytest

from excellikeuno.connection import connect_calc


def _connect_or_skip():
    try:
        from com.sun.star.awt import Point, Size  # type: ignore
    except ImportError as exc:  # pragma: no cover - depends on LibreOffice runtime
        pytest.skip(f"UNO runtime not available: {exc}")

    try:
        desktop, doc, sheet = connect_calc()
    except RuntimeError as exc:  # pragma: no cover - depends on LibreOffice runtime
        pytest.skip(f"UNO runtime not available: {exc}")

    return desktop, doc, sheet, Point, Size


def _add_rectangle(doc, sheet, Point, Size):
    draw_page = sheet.raw.getDrawPage()
    rect = doc.createInstance("com.sun.star.drawing.RectangleShape")
    rect.setPosition(Point(1000, 1000))
    rect.setSize(Size(2000, 1200))
    draw_page.add(rect)
    return rect, draw_page


def test_shapes_collection_wraps_draw_page():
    _, doc, sheet, Point, Size = _connect_or_skip()
    rect, draw_page = _add_rectangle(doc, sheet, Point, Size)
    try:
        count_after = draw_page.getCount()
        shapes = sheet.shapes()
        assert len(shapes) == count_after
        rect_pos = rect.getPosition()
        rect_size = rect.getSize()
        def _same_geom(shape):
            pos = shape.position
            size = shape.size
            return (pos.X, pos.Y, size.Width, size.Height) == (
                rect_pos.X,
                rect_pos.Y,
                rect_size.Width,
                rect_size.Height,
            )

        assert any(_same_geom(shape) for shape in shapes)

        wrapper = sheet.shape(count_after - 1)
        assert _same_geom(wrapper)
    finally:
        draw_page.remove(rect)


def test_shape_position_and_size_roundtrip():
    _, doc, sheet, Point, Size = _connect_or_skip()
    rect, draw_page = _add_rectangle(doc, sheet, Point, Size)
    try:
        wrapper = sheet.shapes()[-1]
        new_pos = Point(rect.getPosition().X + 200, rect.getPosition().Y + 150)
        new_size = Size(rect.getSize().Width + 200, rect.getSize().Height + 100)
        wrapper.position = new_pos
        wrapper.size = new_size
        updated_pos = wrapper.position
        updated_size = wrapper.size
        assert (updated_pos.X, updated_pos.Y) == (new_pos.X, new_pos.Y)
        assert (updated_size.Width, updated_size.Height) == (new_size.Width, new_size.Height)
    finally:
        draw_page.remove(rect)


def test_shape_basic_properties_roundtrip():
    _, doc, sheet, Point, Size = _connect_or_skip()
    rect, draw_page = _add_rectangle(doc, sheet, Point, Size)
    shape = None
    original_name = None
    original_visible = None
    original_printable = None
    try:
        shape = sheet.shapes()[-1]
        original_name = shape.name
        original_visible = shape.visible
        original_printable = shape.printable

        shape.name = "Rect_Test"
        shape.visible = not original_visible
        shape.printable = not original_printable

        assert shape.name == "Rect_Test"
        assert shape.visible is (not original_visible)
        assert shape.printable is (not original_printable)
    finally:
        # restore to reduce UI churn
        if shape is not None and original_name is not None:
            shape.name = original_name
            shape.visible = original_visible
            shape.printable = original_printable
        draw_page.remove(rect)
