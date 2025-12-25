from __future__ import annotations

from typing import Any, cast

from ..core import InterfaceNames, UnoObject
from ..typing import XPropertySet, XShape


class Shape(UnoObject):
    """Wraps a drawing Shape from Calc draw page."""

    def _get_prop(self, name: str) -> Any:
        props = cast(XPropertySet, self.iface(InterfaceNames.X_PROPERTY_SET))
        return props.getPropertyValue(name)

    def _set_prop(self, name: str, value: Any) -> None:
        props = cast(XPropertySet, self.iface(InterfaceNames.X_PROPERTY_SET))
        props.setPropertyValue(name, value)

    @property
    def position(self) -> Any:
        shape = cast(XShape, self.iface(InterfaceNames.X_SHAPE))
        return shape.getPosition()

    @position.setter
    def position(self, value: Any) -> None:
        shape = cast(XShape, self.iface(InterfaceNames.X_SHAPE))
        shape.setPosition(value)

    @property
    def size(self) -> Any:
        shape = cast(XShape, self.iface(InterfaceNames.X_SHAPE))
        return shape.getSize()

    @size.setter
    def size(self, value: Any) -> None:
        shape = cast(XShape, self.iface(InterfaceNames.X_SHAPE))
        shape.setSize(value)

    @property
    def props(self) -> XPropertySet:
        return cast(XPropertySet, self.iface(InterfaceNames.X_PROPERTY_SET))

    # Common Shape properties
    @property
    def name(self) -> str:
        return cast(str, self._get_prop("Name"))

    @name.setter
    def name(self, value: str) -> None:
        self._set_prop("Name", value)

    @property
    def title(self) -> str:
        return cast(str, self._get_prop("Title"))

    @title.setter
    def title(self, value: str) -> None:
        self._set_prop("Title", value)

    @property
    def description(self) -> str:
        return cast(str, self._get_prop("Description"))

    @description.setter
    def description(self, value: str) -> None:
        self._set_prop("Description", value)

    @property
    def visible(self) -> bool:
        return bool(self._get_prop("Visible"))

    @visible.setter
    def visible(self, value: bool) -> None:
        self._set_prop("Visible", bool(value))

    @property
    def printable(self) -> bool:
        return bool(self._get_prop("Printable"))

    @printable.setter
    def printable(self, value: bool) -> None:
        self._set_prop("Printable", bool(value))

    @property
    def move_protect(self) -> bool:
        return bool(self._get_prop("MoveProtect"))

    @move_protect.setter
    def move_protect(self, value: bool) -> None:
        self._set_prop("MoveProtect", bool(value))

    @property
    def size_protect(self) -> bool:
        return bool(self._get_prop("SizeProtect"))

    @size_protect.setter
    def size_protect(self, value: bool) -> None:
        self._set_prop("SizeProtect", bool(value))

    @property
    def z_order(self) -> int:
        return int(self._get_prop("ZOrder"))

    @z_order.setter
    def z_order(self, value: int) -> None:
        self._set_prop("ZOrder", int(value))

    @property
    def layer_id(self) -> int:
        return int(self._get_prop("LayerID"))

    @layer_id.setter
    def layer_id(self, value: int) -> None:
        self._set_prop("LayerID", int(value))

    @property
    def layer_name(self) -> str:
        return cast(str, self._get_prop("LayerName"))

    @layer_name.setter
    def layer_name(self, value: str) -> None:
        self._set_prop("LayerName", value)

    @property
    def hyperlink(self) -> str:
        return cast(str, self._get_prop("Hyperlink"))

    @hyperlink.setter
    def hyperlink(self, value: str) -> None:
        self._set_prop("Hyperlink", value)

    @property
    def navigation_order(self) -> int:
        return int(self._get_prop("NavigationOrder"))

    @navigation_order.setter
    def navigation_order(self, value: int) -> None:
        self._set_prop("NavigationOrder", int(value))

    @property
    def style(self) -> Any:
        return self._get_prop("Style")

    @style.setter
    def style(self, value: Any) -> None:
        self._set_prop("Style", value)

    @property
    def transformation(self) -> Any:
        return self._get_prop("Transformation")

    @transformation.setter
    def transformation(self, value: Any) -> None:
        self._set_prop("Transformation", value)

    @property
    def shape_user_defined_attributes(self) -> Any:
        return self._get_prop("ShapeUserDefinedAttributes")

    @shape_user_defined_attributes.setter
    def shape_user_defined_attributes(self, value: Any) -> None:
        self._set_prop("ShapeUserDefinedAttributes", value)

    @property
    def relative_height(self) -> int:
        return int(self._get_prop("RelativeHeight"))

    @relative_height.setter
    def relative_height(self, value: int) -> None:
        self._set_prop("RelativeHeight", int(value))

    @property
    def relative_width(self) -> int:
        return int(self._get_prop("RelativeWidth"))

    @relative_width.setter
    def relative_width(self, value: int) -> None:
        self._set_prop("RelativeWidth", int(value))

    @property
    def relative_height_relation(self) -> int:
        return int(self._get_prop("RelativeHeightRelation"))

    @relative_height_relation.setter
    def relative_height_relation(self, value: int) -> None:
        self._set_prop("RelativeHeightRelation", int(value))

    @property
    def relative_width_relation(self) -> int:
        return int(self._get_prop("RelativeWidthRelation"))

    @relative_width_relation.setter
    def relative_width_relation(self, value: int) -> None:
        self._set_prop("RelativeWidthRelation", int(value))

    @property
    def decorative(self) -> bool:
        return bool(self._get_prop("Decorative"))

    @decorative.setter
    def decorative(self, value: bool) -> None:
        self._set_prop("Decorative", bool(value))

    @property
    def interop_grab_bag(self) -> Any:
        return self._get_prop("InteropGrabBag")

    @interop_grab_bag.setter
    def interop_grab_bag(self, value: Any) -> None:
        self._set_prop("InteropGrabBag", value)

    # Text anchoring and wrap related properties
    @property
    def anchor_page_no(self) -> int:
        return int(self._get_prop("AnchorPageNo"))

    @anchor_page_no.setter
    def anchor_page_no(self, value: int) -> None:
        self._set_prop("AnchorPageNo", int(value))

    @property
    def anchor_type(self) -> Any:
        return self._get_prop("AnchorType")

    @anchor_type.setter
    def anchor_type(self, value: Any) -> None:
        self._set_prop("AnchorType", value)

    @property
    def anchor_frame(self) -> Any:
        return self._get_prop("AnchorFrame")

    @anchor_frame.setter
    def anchor_frame(self, value: Any) -> None:
        self._set_prop("AnchorFrame", value)

    @property
    def text_range(self) -> Any:
        return self._get_prop("TextRange")

    @text_range.setter
    def text_range(self, value: Any) -> None:
        self._set_prop("TextRange", value)

    @property
    def surround(self) -> Any:
        return self._get_prop("Surround")

    @surround.setter
    def surround(self, value: Any) -> None:
        self._set_prop("Surround", value)

    @property
    def surround_anchor_only(self) -> bool:
        return bool(self._get_prop("SurroundAnchorOnly"))

    @surround_anchor_only.setter
    def surround_anchor_only(self, value: bool) -> None:
        self._set_prop("SurroundAnchorOnly", bool(value))

    @property
    def surround_contour(self) -> bool:
        return bool(self._get_prop("SurroundContour"))

    @surround_contour.setter
    def surround_contour(self, value: bool) -> None:
        self._set_prop("SurroundContour", bool(value))

    @property
    def contour_outside(self) -> bool:
        return bool(self._get_prop("ContourOutside"))

    @contour_outside.setter
    def contour_outside(self, value: bool) -> None:
        self._set_prop("ContourOutside", bool(value))

    @property
    def opaque(self) -> bool:
        return bool(self._get_prop("Opaque"))

    @opaque.setter
    def opaque(self, value: bool) -> None:
        self._set_prop("Opaque", bool(value))

    @property
    def wrap_influence_on_position(self) -> int:
        return int(self._get_prop("WrapInfluenceOnPosition"))

    @wrap_influence_on_position.setter
    def wrap_influence_on_position(self, value: int) -> None:
        self._set_prop("WrapInfluenceOnPosition", int(value))

    # Orientation and positioning
    @property
    def hori_orient(self) -> int:
        return int(self._get_prop("HoriOrient"))

    @hori_orient.setter
    def hori_orient(self, value: int) -> None:
        self._set_prop("HoriOrient", int(value))

    @property
    def hori_orient_position(self) -> int:
        return int(self._get_prop("HoriOrientPosition"))

    @hori_orient_position.setter
    def hori_orient_position(self, value: int) -> None:
        self._set_prop("HoriOrientPosition", int(value))

    @property
    def hori_orient_relation(self) -> int:
        return int(self._get_prop("HoriOrientRelation"))

    @hori_orient_relation.setter
    def hori_orient_relation(self, value: int) -> None:
        self._set_prop("HoriOrientRelation", int(value))

    @property
    def vert_orient(self) -> int:
        return int(self._get_prop("VertOrient"))

    @vert_orient.setter
    def vert_orient(self, value: int) -> None:
        self._set_prop("VertOrient", int(value))

    @property
    def vert_orient_position(self) -> int:
        return int(self._get_prop("VertOrientPosition"))

    @vert_orient_position.setter
    def vert_orient_position(self, value: int) -> None:
        self._set_prop("VertOrientPosition", int(value))

    @property
    def vert_orient_relation(self) -> int:
        return int(self._get_prop("VertOrientRelation"))

    @vert_orient_relation.setter
    def vert_orient_relation(self, value: int) -> None:
        self._set_prop("VertOrientRelation", int(value))

    @property
    def left_margin(self) -> int:
        return int(self._get_prop("LeftMargin"))

    @left_margin.setter
    def left_margin(self, value: int) -> None:
        self._set_prop("LeftMargin", int(value))

    @property
    def right_margin(self) -> int:
        return int(self._get_prop("RightMargin"))

    @right_margin.setter
    def right_margin(self, value: int) -> None:
        self._set_prop("RightMargin", int(value))

    @property
    def top_margin(self) -> int:
        return int(self._get_prop("TopMargin"))

    @top_margin.setter
    def top_margin(self, value: int) -> None:
        self._set_prop("TopMargin", int(value))

    @property
    def bottom_margin(self) -> int:
        return int(self._get_prop("BottomMargin"))

    @bottom_margin.setter
    def bottom_margin(self, value: int) -> None:
        self._set_prop("BottomMargin", int(value))

    # Layout direction helpers
    @property
    def transformation_in_hori_l2r(self) -> Any:
        return self._get_prop("TransformationInHoriL2R")

    @property
    def position_layout_dir(self) -> int:
        return int(self._get_prop("PositionLayoutDir"))

    @position_layout_dir.setter
    def position_layout_dir(self, value: int) -> None:
        self._set_prop("PositionLayoutDir", int(value))

    @property
    def start_position_in_hori_l2r(self) -> Any:
        return self._get_prop("StartPositionInHoriL2R")

    @property
    def end_position_in_hori_l2r(self) -> Any:
        return self._get_prop("EndPositionInHoriL2R")

    # Overlap handling
    @property
    def allow_overlap(self) -> bool:
        return bool(self._get_prop("AllowOverlap"))

    @allow_overlap.setter
    def allow_overlap(self, value: bool) -> None:
        self._set_prop("AllowOverlap", bool(value))
