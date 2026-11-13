"""Shapes, images, lines, connectors, grouping and z-order."""

from __future__ import annotations

import copy
from pathlib import Path
from typing import Any

from lxml import etree
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.util import Emu, Pt

from engines.ooxml.common import logger
from engines.ooxml_rels import assign_fresh_shape_ids, transfer_relationships
from errors import ensure
from utils.colors import normalize_color
from utils.units import emu_to_inches, to_emu


class OOXMLShapesMixin:
    """Shapes, images, lines, connectors, grouping and z-order."""

    def get_shape_details(self, params: dict[str, Any]) -> dict[str, Any]:
        session = self._get_session(str(params["presentation_id"]))
        slide_index = int(params["slide_index"])
        slide = self._get_slide_obj(session, slide_index)

        shape = self._require_shape(slide, params)
        detail = self._detailed_shape_payload(shape)

        return {
            "presentation_id": session.id,
            "slide_index": slide_index,
            **detail,
        }

    def add_shape(self, params: dict[str, Any]) -> dict[str, Any]:
        session = self._get_session(str(params["presentation_id"]))
        slide_index = int(params["slide_index"])
        slide = self._get_slide_obj(session, slide_index)

        shape_type_str = str(params["shape_type"]).lower()
        mso_shape = self._SHAPE_TYPE_MAP.get(shape_type_str)
        ensure(
            mso_shape is not None,
            "validation_error",
            f"Unknown shape_type '{shape_type_str}'.",
            {"available": sorted(self._SHAPE_TYPE_MAP.keys())},
        )

        left = to_emu(params["left"])
        top = to_emu(params["top"])
        width = to_emu(params["width"])
        height = to_emu(params["height"])

        from pptx.util import Emu as EmuVal

        shape = slide.shapes.add_shape(mso_shape, EmuVal(left), EmuVal(top), EmuVal(width), EmuVal(height))

        # Optional fill color
        fill_hex = params.get("fill_hex")
        if fill_hex:
            shape.fill.solid()
            shape.fill.fore_color.rgb = RGBColor.from_string(normalize_color(str(fill_hex)))

        # Optional line color
        line_hex = params.get("line_hex")
        if line_hex:
            shape.line.color.rgb = RGBColor.from_string(normalize_color(str(line_hex)))

        # Optional text
        text = params.get("text")
        if text and hasattr(shape, "text_frame"):
            shape.text_frame.paragraphs[0].text = str(text)

        self._mark_modified(session)

        return {
            "success": True,
            "presentation_id": session.id,
            "slide_index": slide_index,
            "shape_name": shape.name,
            "shape_id": int(shape.shape_id),
            "presentation_state": self.get_presentation_state({"presentation_id": session.id}),
        }

    def delete_shape(self, params: dict[str, Any]) -> dict[str, Any]:
        session = self._get_session(str(params["presentation_id"]))
        slide_index = int(params["slide_index"])
        slide = self._get_slide_obj(session, slide_index)

        shape = self._require_shape(slide, params)
        deleted_shape = shape.name if shape.name else str(shape.shape_id)
        # Remove the shape element from the shape tree
        sp_tree = slide.shapes._spTree
        sp_tree.remove(shape.element)

        self._mark_modified(session)

        return {
            "success": True,
            "presentation_id": session.id,
            "slide_index": slide_index,
            "deleted_shape": deleted_shape,
            "presentation_state": self.get_presentation_state({"presentation_id": session.id}),
        }

    def set_shape_properties(self, params: dict[str, Any]) -> dict[str, Any]:
        session = self._get_session(str(params["presentation_id"]))
        slide_index = int(params["slide_index"])
        slide = self._get_slide_obj(session, slide_index)

        shape = self._require_shape(slide, params)

        # Position / size
        if "left" in params:
            shape.left = Emu(to_emu(params["left"]))
        if "top" in params:
            shape.top = Emu(to_emu(params["top"]))
        if "width" in params:
            shape.width = Emu(to_emu(params["width"]))
        if "height" in params:
            shape.height = Emu(to_emu(params["height"]))

        # Rotation
        if "rotation" in params:
            shape.rotation = float(params["rotation"])

        # Fill
        fill_hex = params.get("fill_hex")
        if fill_hex == "none":
            shape.fill.background()
        elif fill_hex:
            shape.fill.solid()
            shape.fill.fore_color.rgb = RGBColor.from_string(normalize_color(str(fill_hex)))

        # Line / outline
        line_hex = params.get("line_hex")
        if line_hex == "none":
            shape.line.fill.background()
        elif line_hex:
            shape.line.color.rgb = RGBColor.from_string(normalize_color(str(line_hex)))

        line_width_pt = params.get("line_width_pt")
        if line_width_pt is not None:
            shape.line.width = Pt(float(line_width_pt))

        # Name
        if "name" in params:
            shape.name = str(params["name"])

        self._mark_modified(session)

        return {
            "success": True,
            "presentation_id": session.id,
            "slide_index": slide_index,
            "shape_name": shape.name,
            "left_inches": emu_to_inches(shape.left) if shape.left else None,
            "top_inches": emu_to_inches(shape.top) if shape.top else None,
            "width_inches": emu_to_inches(shape.width) if shape.width else None,
            "height_inches": emu_to_inches(shape.height) if shape.height else None,
            "rotation": float(shape.rotation) if hasattr(shape, "rotation") else 0.0,
            "presentation_state": self.get_presentation_state({"presentation_id": session.id}),
        }

    def clone_shape(self, params: dict[str, Any]) -> dict[str, Any]:
        session = self._get_session(str(params["presentation_id"]))
        slide_index = int(params["slide_index"])
        slide = self._get_slide_obj(session, slide_index)

        shape = self._require_shape(slide, params)

        # Target slide (default: same slide)
        target_slide_index = int(params.get("target_slide_index", slide_index))
        target_slide = self._get_slide_obj(session, target_slide_index)

        # Deep copy the shape XML element, then re-home its image/chart/link relationships
        # onto the target slide and give it an id that is unique there.
        new_element = copy.deepcopy(shape.element)
        copy_warnings = transfer_relationships(new_element, slide.part, target_slide.part)
        assign_fresh_shape_ids(new_element, target_slide.shapes._spTree)

        # Optional offset
        offset_left = params.get("offset_left")
        offset_top = params.get("offset_top")

        target_slide.shapes._spTree.append(new_element)

        # Find the new shape so we can apply offset
        # The cloned shape is the last one added
        new_shape = target_slide.shapes[-1] if len(target_slide.shapes) > 0 else None
        if new_shape and offset_left is not None:
            new_shape.left = Emu(int(new_shape.left or 0) + to_emu(offset_left))
        if new_shape and offset_top is not None:
            new_shape.top = Emu(int(new_shape.top or 0) + to_emu(offset_top))

        self._mark_modified(session)

        return {
            "success": True,
            "presentation_id": session.id,
            "source_slide_index": slide_index,
            "target_slide_index": target_slide_index,
            "cloned_shape_name": new_shape.name if new_shape else "unknown",
            "cloned_shape_id": int(new_shape.shape_id) if new_shape else None,
            "warnings": copy_warnings,
            "presentation_state": self.get_presentation_state({"presentation_id": session.id}),
        }

    def group_shapes(self, params: dict[str, Any]) -> dict[str, Any]:
        session = self._get_session(str(params["presentation_id"]))
        slide_index = int(params["slide_index"])
        slide = self._get_slide_obj(session, slide_index)

        shape_names = params.get("shape_names", [])
        shape_ids = params.get("shape_ids", [])
        ensure(
            len(shape_names) >= 2 or len(shape_ids) >= 2,
            "validation_error",
            "At least 2 shapes required for grouping.",
        )

        identifiers = shape_names if shape_names else shape_ids
        shapes_to_group = [self._find_shape(slide, ident) for ident in identifiers]

        sp_tree = slide.shapes._spTree

        # Calculate bounding box
        min_left = min(s.left for s in shapes_to_group if s.left is not None)
        min_top = min(s.top for s in shapes_to_group if s.top is not None)
        max_right = max((s.left or 0) + (s.width or 0) for s in shapes_to_group)
        max_bottom = max((s.top or 0) + (s.height or 0) for s in shapes_to_group)

        # Create group shape element
        grp_sp = etree.SubElement(sp_tree, "{http://schemas.openxmlformats.org/presentationml/2006/main}grpSp")

        # Group shape properties
        grp_sp_pr = etree.SubElement(grp_sp, "{http://schemas.openxmlformats.org/presentationml/2006/main}grpSpPr")
        xfrm = etree.SubElement(grp_sp_pr, "{http://schemas.openxmlformats.org/drawingml/2006/main}xfrm")
        etree.SubElement(
            xfrm,
            "{http://schemas.openxmlformats.org/drawingml/2006/main}off",
            attrib={"x": str(min_left), "y": str(min_top)},
        )
        etree.SubElement(
            xfrm,
            "{http://schemas.openxmlformats.org/drawingml/2006/main}ext",
            attrib={"cx": str(max_right - min_left), "cy": str(max_bottom - min_top)},
        )
        etree.SubElement(
            xfrm,
            "{http://schemas.openxmlformats.org/drawingml/2006/main}chOff",
            attrib={"x": str(min_left), "y": str(min_top)},
        )
        etree.SubElement(
            xfrm,
            "{http://schemas.openxmlformats.org/drawingml/2006/main}chExt",
            attrib={"cx": str(max_right - min_left), "cy": str(max_bottom - min_top)},
        )

        # Non-visual group shape properties
        next_id = max((int(s.shape_id) for s in slide.shapes), default=0) + 1
        nvGrpSpPr = etree.SubElement(grp_sp, "{http://schemas.openxmlformats.org/presentationml/2006/main}nvGrpSpPr")
        cNvPr = etree.SubElement(nvGrpSpPr, "{http://schemas.openxmlformats.org/presentationml/2006/main}cNvPr")
        cNvPr.set("id", str(next_id))
        cNvPr.set("name", str(params.get("group_name", "Group")))
        etree.SubElement(nvGrpSpPr, "{http://schemas.openxmlformats.org/presentationml/2006/main}cNvGrpSpPr")
        etree.SubElement(nvGrpSpPr, "{http://schemas.openxmlformats.org/presentationml/2006/main}nvPr")

        # Move shapes into group
        for shape in shapes_to_group:
            sp_tree.remove(shape.element)
            grp_sp.append(shape.element)

        self._mark_modified(session)

        return {
            "success": True,
            "presentation_id": session.id,
            "slide_index": slide_index,
            "group_name": str(params.get("group_name", "Group")),
            "shapes_grouped": len(shapes_to_group),
            "presentation_state": self.get_presentation_state({"presentation_id": session.id}),
        }

    def ungroup_shapes(self, params: dict[str, Any]) -> dict[str, Any]:
        session = self._get_session(str(params["presentation_id"]))
        slide_index = int(params["slide_index"])
        slide = self._get_slide_obj(session, slide_index)

        shape = self._require_shape(slide, params)

        ensure(
            shape.shape_type == MSO_SHAPE_TYPE.GROUP,
            "conflict",
            "Shape is not a group.",
        )

        sp_tree = slide.shapes._spTree
        group_element = shape.element

        # Collect children before removing group
        children = list(group_element)
        child_names: list[str] = []

        for child in children:
            # Skip non-visual properties (grpSpPr, nvGrpSpPr)
            tag = etree.QName(child).localname
            if tag in ("grpSpPr", "nvGrpSpPr"):
                continue
            group_element.remove(child)
            sp_tree.append(child)
            # Try to get name
            cNvPr = child.find(".//{http://schemas.openxmlformats.org/drawingml/2006/main}cNvPr")
            if cNvPr is None:
                # Try presentation namespace
                for ns_prefix in [
                    "http://schemas.openxmlformats.org/presentationml/2006/main",
                    "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
                ]:
                    cNvPr = child.find(f".//{{{ns_prefix}}}cNvPr")
                    if cNvPr is not None:
                        break
            name = cNvPr.get("name", "Unknown") if cNvPr is not None else "Unknown"
            child_names.append(name)

        # Remove the now-empty group
        sp_tree.remove(group_element)

        self._mark_modified(session)

        return {
            "success": True,
            "presentation_id": session.id,
            "slide_index": slide_index,
            "ungrouped_shapes": child_names,
            "presentation_state": self.get_presentation_state({"presentation_id": session.id}),
        }

    def set_shape_z_order(self, params: dict[str, Any]) -> dict[str, Any]:
        session = self._get_session(str(params["presentation_id"]))
        slide_index = int(params["slide_index"])
        slide = self._get_slide_obj(session, slide_index)

        shape = self._require_shape(slide, params)

        action = str(params.get("action", "front")).lower()
        ensure(
            action in ("front", "back", "forward", "backward"),
            "validation_error",
            "action must be 'front', 'back', 'forward', or 'backward'.",
        )

        sp_tree = slide.shapes._spTree
        element = shape.element

        if action == "front":
            sp_tree.remove(element)
            sp_tree.append(element)
        elif action == "back":
            sp_tree.remove(element)
            sp_tree.insert(2, element)  # index 0,1 are typically nvGrpSpPr and grpSpPr
        elif action == "forward":
            idx = list(sp_tree).index(element)
            sp_tree.remove(element)
            sp_tree.insert(min(idx + 1, len(sp_tree)), element)
        elif action == "backward":
            idx = list(sp_tree).index(element)
            sp_tree.remove(element)
            sp_tree.insert(max(idx - 1, 2), element)

        self._mark_modified(session)

        return {
            "success": True,
            "presentation_id": session.id,
            "slide_index": slide_index,
            "shape_name": shape.name,
            "action": action,
            "presentation_state": self.get_presentation_state({"presentation_id": session.id}),
        }

    def add_image(self, params: dict[str, Any]) -> dict[str, Any]:
        session = self._get_session(str(params["presentation_id"]))
        slide_index = int(params["slide_index"])
        slide = self._get_slide_obj(session, slide_index)

        image_path = str(params["image_path"])
        ensure(Path(image_path).is_file(), "not_found", f"Image file not found: {image_path}")

        left = to_emu(params.get("left", "0in"))
        top = to_emu(params.get("top", "0in"))

        # Width/height: at least one should be specified, or both
        width_param = params.get("width")
        height_param = params.get("height")

        kwargs: dict[str, Any] = {
            "image_file": image_path,
            "left": Emu(left),
            "top": Emu(top),
        }
        if width_param:
            kwargs["width"] = Emu(to_emu(width_param))
        if height_param:
            kwargs["height"] = Emu(to_emu(height_param))

        pic = slide.shapes.add_picture(**kwargs)

        self._mark_modified(session)

        return {
            "success": True,
            "presentation_id": session.id,
            "slide_index": slide_index,
            "shape_name": pic.name,
            "shape_id": int(pic.shape_id),
            "left_inches": emu_to_inches(pic.left),
            "top_inches": emu_to_inches(pic.top),
            "width_inches": emu_to_inches(pic.width),
            "height_inches": emu_to_inches(pic.height),
            "presentation_state": self.get_presentation_state({"presentation_id": session.id}),
        }

    def add_line(self, params: dict[str, Any]) -> dict[str, Any]:
        session = self._get_session(str(params["presentation_id"]))
        slide_index = int(params["slide_index"])
        slide = self._get_slide_obj(session, slide_index)

        begin_x = to_emu(params["begin_x"])
        begin_y = to_emu(params["begin_y"])
        end_x = to_emu(params["end_x"])
        end_y = to_emu(params["end_y"])

        # python-pptx add_connector doesn't exist easily, use XML approach
        from pptx.oxml.ns import qn

        sp_tree = slide.shapes._spTree

        # Calculate position and dimensions from endpoints
        left = min(begin_x, end_x)
        top = min(begin_y, end_y)
        cx = abs(end_x - begin_x)
        cy = abs(end_y - begin_y)

        # Determine if line needs flipping
        flip_h = "1" if end_x < begin_x else "0"
        flip_v = "1" if end_y < begin_y else "0"

        # Build connector shape XML
        cxn_sp = etree.SubElement(sp_tree, qn("p:cxnSp"))

        # Non-visual properties — generate a unique shape ID
        next_id = max((int(s.shape_id) for s in slide.shapes), default=0) + 1
        nv = etree.SubElement(cxn_sp, qn("p:nvCxnSpPr"))
        cNvPr = etree.SubElement(nv, qn("p:cNvPr"))
        cNvPr.set("id", str(next_id))
        cNvPr.set("name", str(params.get("line_name", "Connector")))
        etree.SubElement(nv, qn("p:cNvCxnSpPr"))
        etree.SubElement(nv, qn("p:nvPr"))

        # Shape properties
        sp_pr = etree.SubElement(cxn_sp, qn("p:spPr"))
        xfrm = etree.SubElement(sp_pr, qn("a:xfrm"))
        if flip_h == "1":
            xfrm.set("flipH", "1")
        if flip_v == "1":
            xfrm.set("flipV", "1")
        etree.SubElement(xfrm, qn("a:off"), attrib={"x": str(left), "y": str(top)})
        etree.SubElement(xfrm, qn("a:ext"), attrib={"cx": str(max(cx, 1)), "cy": str(max(cy, 1))})

        # Preset geometry (straight line)
        prst_geom = etree.SubElement(sp_pr, qn("a:prstGeom"), attrib={"prst": "line"})
        etree.SubElement(prst_geom, qn("a:avLst"))

        # Line properties
        line_hex = params.get("color_hex")
        line_width_pt = params.get("width_pt", 1.0)

        ln = etree.SubElement(sp_pr, qn("a:ln"), attrib={"w": str(int(Pt(float(line_width_pt))))})
        if line_hex:
            solid = etree.SubElement(ln, qn("a:solidFill"))
            etree.SubElement(solid, qn("a:srgbClr"), attrib={"val": normalize_color(str(line_hex))})
        else:
            solid = etree.SubElement(ln, qn("a:solidFill"))
            etree.SubElement(solid, qn("a:srgbClr"), attrib={"val": "000000"})

        # Dash style
        dash_style = params.get("dash_style")
        if dash_style:
            etree.SubElement(ln, qn("a:prstDash"), attrib={"val": str(dash_style)})

        self._mark_modified(session)

        return {
            "success": True,
            "presentation_id": session.id,
            "slide_index": slide_index,
            "line_name": str(params.get("line_name", "Connector")),
            "begin": {"x_inches": emu_to_inches(begin_x), "y_inches": emu_to_inches(begin_y)},
            "end": {"x_inches": emu_to_inches(end_x), "y_inches": emu_to_inches(end_y)},
            "presentation_state": self.get_presentation_state({"presentation_id": session.id}),
        }

    def copy_shape_between_decks(self, params: dict[str, Any]) -> dict[str, Any]:
        source_session = self._get_session(str(params["source_presentation_id"]))
        target_session = self._get_session(str(params["target_presentation_id"]))

        source_slide_index = int(params["source_slide_index"])
        target_slide_index = int(params["target_slide_index"])

        source_slide = self._get_slide_obj(source_session, source_slide_index)
        target_slide = self._get_slide_obj(target_session, target_slide_index)

        source_shape = self._require_shape(source_slide, params)

        # Deep copy the shape XML element
        new_element = copy.deepcopy(source_shape.element)

        # Copy every part the shape depends on (images, charts, media, links) into the target
        # deck and repoint the shape at them; give the shape an id unique on the target slide.
        copy_warnings = transfer_relationships(new_element, source_slide.part, target_slide.part)
        assign_fresh_shape_ids(new_element, target_slide.shapes._spTree)

        # Apply offset if specified
        offset_left = params.get("offset_left")
        offset_top = params.get("offset_top")
        if offset_left is not None:
            try:
                off_elem = new_element.find(".//{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetml}off")
                if off_elem is None:
                    # Try presentation ML namespace
                    for off in new_element.iter():
                        if off.tag.endswith("}off") or off.tag == "off":
                            off_elem = off
                            break
                if off_elem is not None and off_elem.get("x") is not None:
                    current_x = int(off_elem.get("x", "0"))
                    off_elem.set("x", str(current_x + to_emu(offset_left)))
            except Exception as exc:
                logger.warning("could not apply offset_left: %s", exc)

        if offset_top is not None:
            try:
                for off in new_element.iter():
                    if (off.tag.endswith("}off") or off.tag == "off") and off.get("y") is not None:
                        current_y = int(off.get("y", "0"))
                        off.set("y", str(current_y + to_emu(offset_top)))
                        break
            except Exception as exc:
                logger.warning("could not apply offset_top: %s", exc)

        # Insert into target slide's shape tree
        target_sp_tree = target_slide.shapes._spTree
        target_sp_tree.append(new_element)

        self._mark_modified(target_session)

        # Get the new shape's name from the inserted element
        new_shape_name = ""
        try:
            nv_elem = new_element.find(".//{http://schemas.openxmlformats.org/presentationml/2006/main}cNvPr")
            if nv_elem is None:
                nv_elem = new_element.find(".//{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetml}cNvPr")
            if nv_elem is None:
                # Generic search for cNvPr
                for elem in new_element.iter():
                    if elem.tag.endswith("}cNvPr") or elem.tag == "cNvPr":
                        nv_elem = elem
                        break
            if nv_elem is not None:
                new_shape_name = nv_elem.get("name", "")
        except Exception:
            pass

        return {
            "success": True,
            "source_presentation_id": source_session.id,
            "target_presentation_id": target_session.id,
            "source_slide_index": source_slide_index,
            "target_slide_index": target_slide_index,
            "source_shape_name": source_shape.name,
            "new_shape_name": new_shape_name,
            "warnings": copy_warnings,
            "presentation_state": self.get_presentation_state({"presentation_id": target_session.id}),
        }

    def set_shape_fill_gradient(self, params: dict[str, Any]) -> dict[str, Any]:
        session = self._get_session(str(params["presentation_id"]))
        slide_index = int(params["slide_index"])
        slide = self._get_slide_obj(session, slide_index)

        shape = self._require_shape(slide, params)

        fill = shape.fill
        fill.gradient()

        if "angle" in params:
            fill.gradient_angle = float(params["angle"])

        stops = params.get("stops", [])
        if stops:
            # Clear existing stops (python-pptx usually creates 2 default stops)
            while len(fill.gradient_stops) > 0:
                # Remove from underlying XML to clear it
                gs_lst = shape.element.find(".//{http://schemas.openxmlformats.org/drawingml/2006/main}gsLst")
                if gs_lst is not None:
                    for gs in gs_lst:
                        gs_lst.remove(gs)
                # python-pptx doesn't have a direct clear() method for gradient_stops
                break

            # Adding new stops requires XML manipulation because python-pptx has limited write support for stops
            gs_lst = shape.element.find(".//{http://schemas.openxmlformats.org/drawingml/2006/main}gsLst")
            if gs_lst is None:
                grad_fill = shape.element.find(".//{http://schemas.openxmlformats.org/drawingml/2006/main}gradFill")
                if grad_fill is not None:
                    gs_lst = etree.Element("{http://schemas.openxmlformats.org/drawingml/2006/main}gsLst")
                    grad_fill.insert(0, gs_lst)

            if gs_lst is not None:
                for stop in stops:
                    pos = int(float(stop.get("position", 0.0)) * 100000)  # 0 to 100000 (0% to 100%)
                    hex_color = str(stop.get("color_hex", "FFFFFF")).strip("#")

                    gs_elem = etree.SubElement(gs_lst, "{http://schemas.openxmlformats.org/drawingml/2006/main}gs")
                    gs_elem.set("pos", str(pos))

                    srgb_clr = etree.SubElement(
                        gs_elem, "{http://schemas.openxmlformats.org/drawingml/2006/main}srgbClr"
                    )
                    srgb_clr.set("val", hex_color)

        self._mark_modified(session)
        return {
            "success": True,
            "shape_name": shape.name,
            "presentation_state": self.get_presentation_state({"presentation_id": session.id}),
        }

    def add_connector(self, params: dict[str, Any]) -> dict[str, Any]:
        session = self._get_session(str(params["presentation_id"]))
        slide_index = int(params["slide_index"])
        slide = self._get_slide_obj(session, slide_index)

        from pptx.enum.shapes import MSO_CONNECTOR

        type_map = {
            "straight": MSO_CONNECTOR.STRAIGHT,
            "elbow": MSO_CONNECTOR.ELBOW,
            "curve": MSO_CONNECTOR.CURVE,
        }
        ctype_str = str(params.get("connector_type", "straight")).lower()
        connector_type = type_map.get(ctype_str, MSO_CONNECTOR.STRAIGHT)

        # Temporary coordinates for creation
        connector = slide.shapes.add_connector(connector_type, Emu(0), Emu(0), Emu(100), Emu(100))

        if "begin_shape_name" in params or "begin_shape_id" in params:
            b_shape = self._find_shape(slide, params.get("begin_shape_name") or params.get("begin_shape_id"))
            b_idx = int(params.get("begin_connection_site", 0))
            connector.begin_connect(b_shape, b_idx)

        if "end_shape_name" in params or "end_shape_id" in params:
            e_shape = self._find_shape(slide, params.get("end_shape_name") or params.get("end_shape_id"))
            e_idx = int(params.get("end_connection_site", 0))
            connector.end_connect(e_shape, e_idx)

        # Apply formatting
        if "color_hex" in params:
            connector.line.color.rgb = RGBColor.from_string(str(params["color_hex"]).strip("#"))

        if "width_pt" in params:
            connector.line.width = Pt(float(params["width_pt"]))

        self._mark_modified(session)
        return {
            "success": True,
            "shape_name": connector.name,
            "shape_id": int(connector.shape_id),
            "presentation_state": self.get_presentation_state({"presentation_id": session.id}),
        }
