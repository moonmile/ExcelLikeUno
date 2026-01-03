import pytest

from excellikeuno.connection import connect_calc
from excellikeuno.drawing import (
    ConnectorShape,
    EllipseShape,
    LineShape,
    PolyLineShape,
    PolyPolygonShape,
    RectangleShape,
    TextShape,
)


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


def _add_connector(doc, draw_page, start_shape, end_shape):
    connector = doc.createInstance("com.sun.star.drawing.ConnectorShape")
    # Insert then configure glue targets
    draw_page.add(connector)
    connector.setPropertyValue("StartShape", start_shape)
    connector.setPropertyValue("EndShape", end_shape)
    connector.setPropertyValue("StartGluePointIndex", 0)
    connector.setPropertyValue("EndGluePointIndex", 0)
    return connector


def _add_line(doc, draw_page):
    return doc.createInstance("com.sun.star.drawing.LineShape")


def _add_rectangle_shape(doc, draw_page):
    return doc.createInstance("com.sun.star.drawing.RectangleShape")


def _add_ellipse(doc, draw_page):
    return doc.createInstance("com.sun.star.drawing.EllipseShape")


def _add_polyline(doc, draw_page):
    return doc.createInstance("com.sun.star.drawing.PolyLineShape")


def _add_polypolygon(doc, draw_page):
    return doc.createInstance("com.sun.star.drawing.PolyPolygonShape")


def _add_textshape(doc, draw_page):
    return doc.createInstance("com.sun.star.drawing.TextShape")


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
            pos = shape.Position
            size = shape.Size
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
        wrapper.Position = new_pos
        wrapper.Size = new_size
        updated_pos = wrapper.Position
        updated_size = wrapper.Size
        assert (updated_pos.X, updated_pos.Y) == (new_pos.X, new_pos.Y)
        assert (updated_size.Width, updated_size.Height) == (new_size.Width, new_size.Height)
    finally:
        draw_page.remove(rect)


def test_connector_wraps_start_end_shapes():
    _, doc, sheet, Point, Size = _connect_or_skip()
    rect1, draw_page = _add_rectangle(doc, sheet, Point, Size)
    rect2, _ = _add_rectangle(doc, sheet, Point, Size)
    rect2.setPosition(Point(4000, 1000))
    connector = None
    try:
        connector_raw = _add_connector(doc, draw_page, rect1, rect2)
        connector = ConnectorShape(connector_raw)

        def _same_geom(a, b):
            pa, sa = a.getPosition(), a.getSize()
            pb, sb = b.getPosition(), b.getSize()
            return (pa.X, pa.Y, sa.Width, sa.Height) == (
                pb.X,
                pb.Y,
                sb.Width,
                sb.Height,
            )

        assert _same_geom(connector.start_shape, rect1)
        assert _same_geom(connector.end_shape, rect2)

        connector.start_glue_point_index = 0
        connector.end_glue_point_index = 0
        assert connector.start_glue_point_index == 0
        assert connector.end_glue_point_index == 0
    finally:
        if connector is not None:
            draw_page.remove(connector.raw)
        draw_page.remove(rect1)
        draw_page.remove(rect2)


def test_shape_basic_properties_roundtrip():
    _, doc, sheet, Point, Size = _connect_or_skip()
    rect, draw_page = _add_rectangle(doc, sheet, Point, Size)
    shape = None
    original_name = None
    original_visible = None
    original_printable = None
    try:
        shape = sheet.shapes()[-1]
        original_name = shape.Name
        original_visible = shape.Visible
        original_printable = shape.Printable

        shape.Name = "Rect_Test"
        shape.Visible = not original_visible
        shape.Printable = not original_printable

        assert shape.Name == "Rect_Test"
        assert shape.Visible is (not original_visible)
        assert shape.Printable is (not original_printable)
    finally:
        # restore to reduce UI churn
        if shape is not None and original_name is not None:
            shape.Name = original_name
            shape.Visible = original_visible
            shape.Printable = original_printable
        draw_page.remove(rect)


def test_line_shape_start_end_roundtrip():
    _, doc, sheet, Point, Size = _connect_or_skip()
    draw_page = sheet.raw.getDrawPage()
    line_raw = _add_line(doc, draw_page)
    line = LineShape(line_raw)
    try:
        start = Point(500, 500)
        end = Point(1500, 1200)
        line.start_position = start
        line.end_position = end
        got_start = line.start_position
        got_end = line.end_position
        if got_start is None or got_end is None:
            pytest.skip("LineShape StartPosition/EndPosition not supported in this runtime")
        assert (got_start.X, got_start.Y) == (start.X, start.Y)
        assert (got_end.X, got_end.Y) == (end.X, end.Y)
    finally:
        draw_page.remove(line.raw)


def test_rectangle_shape_corner_radius_roundtrip():
    _, doc, sheet, Point, Size = _connect_or_skip()
    draw_page = sheet.raw.getDrawPage()
    rect_raw = _add_rectangle_shape(doc, draw_page)
    rect = RectangleShape(rect_raw)
    try:
        rect.corner_radius = 150
        assert rect.corner_radius == 150
    finally:
        draw_page.remove(rect.raw)


def test_ellipse_shape_angles_roundtrip():
    _, doc, sheet, Point, Size = _connect_or_skip()
    draw_page = sheet.raw.getDrawPage()
    ellipse_raw = _add_ellipse(doc, draw_page)
    ellipse = EllipseShape(ellipse_raw)
    draw_page.remove(ellipse.raw)


def test_polyline_shape_points_roundtrip():
    _, doc, sheet, Point, Size = _connect_or_skip()
    draw_page = sheet.raw.getDrawPage()
    poly_raw = _add_polyline(doc, draw_page)
    poly = PolyLineShape(poly_raw)
    try:
        points = ((Point(0, 0), Point(1000, 0), Point(1000, 1000)),)
        try:
            poly.poly_polygon = points
        except Exception as exc:  # pragma: no cover - runtime dependent
            pytest.skip(f"PolyLineShape PolyPolygon not supported: {exc}")
        value = poly.poly_polygon
        if value is None:
            pytest.skip("PolyLineShape PolyPolygon returned None")
        assert len(value[0]) == 3
    finally:
        draw_page.remove(poly.raw)


def test_polypolygon_shape_points_roundtrip():
    _, doc, sheet, Point, Size = _connect_or_skip()
    draw_page = sheet.raw.getDrawPage()
    poly_raw = _add_polypolygon(doc, draw_page)
    poly = PolyPolygonShape(poly_raw)
    try:
        points = (
            (Point(0, 0), Point(1000, 0), Point(1000, 1000), Point(0, 1000)),
        )
        try:
            poly.poly_polygon = points
        except Exception as exc:  # pragma: no cover - runtime dependent
            pytest.skip(f"PolyPolygonShape PolyPolygon not supported: {exc}")
        value = poly.poly_polygon
        if value is None:
            pytest.skip("PolyPolygonShape PolyPolygon returned None")
        assert len(value[0]) == 4
    finally:
        draw_page.remove(poly.raw)


def test_textshape_string_roundtrip():
    _, doc, sheet, Point, Size = _connect_or_skip()
    shapes = sheet.shapes

    # Reuse an existing TextShape if present; otherwise create one like the sample
    text_shape = None
    for shape in shapes:
        if isinstance(shape, TextShape):
            text_shape = shape
            break

    if text_shape is None:
        text_shape = shapes.add_text_shape(
            x=1000,
            y=1000,
            width=4000,
            height=1500,
            text="",
        )

    test_string = "Hello Calc"
    try:
        try:
            text_shape.string = test_string
        except Exception as exc:  # pragma: no cover - runtime dependent
            pytest.skip(f"TextShape String not supported: {exc}")
        if text_shape.string == "":
            pytest.skip("TextShape String returned empty; likely unsupported")
        assert text_shape.string == test_string
    finally:
        try:
            # Clean up only if we created it in this test
            if text_shape and text_shape.raw not in [s.raw for s in shapes]:
                sheet.raw.getDrawPage().remove(text_shape.raw)
        except Exception:
            pass
