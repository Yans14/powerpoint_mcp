"""Relationship-aware copying of slide XML for the OOXML engine.

Shape XML refers to images, charts, media and hyperlinks through relationship ids (``r:embed``,
``r:id``, ...) that are only meaningful inside the part that owns the XML. Copying an element to
another slide without carrying those relationships over leaves dangling ids, which PowerPoint
reports as a corrupt file. The helpers here copy XML and re-home every relationship it uses.
"""

from __future__ import annotations

import copy
import io
import re
from typing import Any

from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.opc.package import XmlPart
from pptx.opc.packuri import PackURI
from pptx.parts.image import ImagePart

_R_NS_PREFIX = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"

# Parts that hold per-shape data. Sharing one between two shapes would make an edit to one
# silently change the other, so they are copied instead of re-linked.
_CLONE_RELTYPES = frozenset({RT.CHART})

# Targets that only make sense inside their own package and cannot follow a shape to another deck.
_PACKAGE_LOCAL_RELTYPES = frozenset({RT.SLIDE, RT.SLIDE_LAYOUT, RT.SLIDE_MASTER, RT.NOTES_SLIDE})

_PARTNAME_SUFFIX_RE = re.compile(r"\d*(\.[^./]*)?$")

_SHAPE_ID_REF_TAGS = ("stCxn", "endCxn")


class _CopyContext:
    """Tracks cloned parts and reserved partnames across one copy operation."""

    def __init__(self) -> None:
        self.clones: dict[int, Any] = {}
        self.reserved_partnames: set[str] = set()


def _local_name(tag: Any) -> str:
    return tag.rsplit("}", 1)[-1] if isinstance(tag, str) else ""


def _relationship_attributes(element):
    """Yield ``(element, attribute_name, rId)`` for every relationship reference under ``element``."""
    for node in element.iter():
        if not isinstance(node.tag, str):
            continue
        for attr_name, value in node.attrib.items():
            if attr_name.startswith(_R_NS_PREFIX) and value:
                yield node, attr_name, value


def _next_partname(package, source_partname: str, context: _CopyContext) -> PackURI:
    template = _PARTNAME_SUFFIX_RE.sub(lambda m: "%d" + (m.group(1) or ""), source_partname.replace("%", "%%"), count=1)
    taken = {str(part.partname) for part in package.iter_parts()} | context.reserved_partnames
    number = 1
    while template % number in taken:
        number += 1
    partname = template % number
    context.reserved_partnames.add(partname)
    return PackURI(partname)


def _clone_part(part, package, context: _CopyContext):
    """Deep-copy ``part`` and everything it links to into ``package`` under fresh partnames."""
    cached = context.clones.get(id(part))
    if cached is not None:
        return cached

    partname = _next_partname(package, str(part.partname), context)
    clone = type(part).load(partname, part.content_type, package, part.blob)
    context.clones[id(part)] = clone

    rid_map: dict[str, str] = {}
    for rid, rel in part.rels.items():
        if rel.is_external:
            rid_map[rid] = clone.relate_to(rel.target_ref, rel.reltype, is_external=True)
        else:
            rid_map[rid] = clone.relate_to(_clone_part(rel.target_part, package, context), rel.reltype)

    if isinstance(clone, XmlPart):
        for node, attr_name, rid in list(_relationship_attributes(clone._element)):
            if rid in rid_map:
                node.set(attr_name, rid_map[rid])

    return clone


def _drop_reference(node, attr_name: str) -> None:
    """Remove a relationship reference that cannot be carried over."""
    parent = node.getparent()
    if _local_name(node.tag) in ("hlinkClick", "hlinkHover") and parent is not None:
        parent.remove(node)
    else:
        del node.attrib[attr_name]


def transfer_relationships(element, source_part, target_part) -> list[str]:
    """Point every relationship reference under ``element`` at ``target_part``.

    ``element`` must be XML copied from ``source_part``. Each referenced relationship is recreated
    on ``target_part`` and the reference is rewritten to the new id. Returns human-readable warnings
    for references that could not be carried over.
    """
    warnings: list[str] = []
    context = _CopyContext()
    same_package = source_part.package is target_part.package
    rid_map: dict[str, str | None] = {}

    for node, attr_name, rid in list(_relationship_attributes(element)):
        if rid not in rid_map:
            rid_map[rid] = _transfer_one(rid, source_part, target_part, same_package, context, warnings)

        new_rid = rid_map[rid]
        if new_rid is None:
            _drop_reference(node, attr_name)
        else:
            node.set(attr_name, new_rid)

    return warnings


def _transfer_one(
    rid: str,
    source_part,
    target_part,
    same_package: bool,
    context: _CopyContext,
    warnings: list[str],
) -> str | None:
    rel = source_part.rels.get(rid)
    if rel is None:
        warnings.append(f"Source references missing relationship '{rid}'; the reference was removed.")
        return None

    if rel.is_external:
        return target_part.relate_to(rel.target_ref, rel.reltype, is_external=True)

    linked_part = rel.target_part

    if same_package:
        if rel.reltype in _CLONE_RELTYPES:
            linked_part = _clone_part(linked_part, target_part.package, context)
        return target_part.relate_to(linked_part, rel.reltype)

    if rel.reltype in _PACKAGE_LOCAL_RELTYPES:
        warnings.append("A link to a slide of the source presentation was removed; it has no target in this deck.")
        return None

    if isinstance(linked_part, ImagePart) and hasattr(target_part, "get_or_add_image_part"):
        # Lets the target package de-duplicate the image and pick a non-colliding partname.
        _, new_rid = target_part.get_or_add_image_part(io.BytesIO(linked_part.blob))
        return new_rid

    return target_part.relate_to(_clone_part(linked_part, target_part.package, context), rel.reltype)


def copy_slide_content(source_slide, target_slide) -> list[str]:
    """Make ``target_slide`` a faithful copy of ``source_slide``.

    Copies every shape (placeholders included, with their formatting), the background, and
    slide-level transition and animation data, then re-homes all relationships. The target's
    ``spTree`` element is refilled in place because python-pptx caches a reference to it.
    """
    src_sld, dst_sld = source_slide._element, target_slide._element
    src_csld, dst_csld = src_sld.cSld, dst_sld.cSld
    src_tree, dst_tree = src_csld.spTree, dst_csld.spTree

    for child in list(dst_tree):
        dst_tree.remove(child)
    for child in src_tree:
        dst_tree.append(copy.deepcopy(child))

    for child in list(dst_csld):
        if child is not dst_tree:
            dst_csld.remove(child)
    before_tree = True
    for child in src_csld:
        if child is src_tree:
            before_tree = False
        elif before_tree:
            dst_tree.addprevious(copy.deepcopy(child))
        else:
            dst_csld.append(copy.deepcopy(child))

    for child in list(dst_sld):
        if child is not dst_csld:
            dst_sld.remove(child)
    for child in src_sld:
        if child is not src_csld:
            dst_sld.append(copy.deepcopy(child))

    for owner_src, owner_dst in ((src_csld, dst_csld), (src_sld, dst_sld)):
        for attr_name, value in owner_src.attrib.items():
            owner_dst.set(attr_name, value)

    return transfer_relationships(dst_sld, source_slide.part, target_slide.part)


def assign_fresh_shape_ids(element, sp_tree) -> None:
    """Give a copied shape (and any shapes nested in it) ids unused in ``sp_tree``.

    Call before inserting ``element`` into ``sp_tree``. Connector endpoints inside the copy that
    point at shapes also inside the copy are updated to follow the new ids.
    """
    used = [
        int(node.get("id"))
        for node in sp_tree.iter()
        if _local_name(node.tag) == "cNvPr" and (node.get("id") or "").isdigit()
    ]
    next_id = max(used, default=0) + 1

    id_map: dict[str, str] = {}
    for node in element.iter():
        if _local_name(node.tag) == "cNvPr" and node.get("id") is not None:
            id_map[node.get("id")] = str(next_id)
            node.set("id", str(next_id))
            next_id += 1

    for node in element.iter():
        if _local_name(node.tag) in _SHAPE_ID_REF_TAGS and node.get("id") in id_map:
            node.set("id", id_map[node.get("id")])
