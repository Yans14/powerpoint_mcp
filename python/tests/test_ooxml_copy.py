"""Copy operations must leave a package that reopens with every relationship intact."""

from __future__ import annotations

import zipfile

import pytest
from PIL import Image
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.oxml.ns import qn

from service import PowerPointService

_R_NS_PREFIX = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"


@pytest.fixture
def service():
    svc = PowerPointService()
    yield svc
    svc.shutdown()


@pytest.fixture
def image_path(tmp_path):
    path = tmp_path / "red.png"
    Image.new("RGB", (32, 32), "red").save(path)
    return str(path)


def _new_deck(service, slide_count=1):
    pid = service.dispatch("pptx_create_presentation", {})["presentation_id"]
    for _ in range(slide_count):
        service.dispatch("pptx_add_slide", {"presentation_id": pid, "layout_name": "Title and Content"})
    return pid


def _save_and_reopen(service, pid, tmp_path, name="out.pptx"):
    out = tmp_path / name
    service.dispatch("pptx_save_presentation", {"presentation_id": pid, "output_path": str(out)})
    with zipfile.ZipFile(out) as archive:
        names = archive.namelist()
        assert len(names) == len(set(names)), "saved package contains duplicate part names"
    return Presentation(str(out))


def _assert_relationships_resolve(prs):
    """Every r:* reference on every slide must resolve to a relationship of that slide."""
    for index, slide in enumerate(prs.slides, start=1):
        for node in slide._element.iter():
            if not isinstance(node.tag, str):
                continue
            for attr_name, rid in node.attrib.items():
                if attr_name.startswith(_R_NS_PREFIX) and rid:
                    assert rid in slide.part.rels, f"slide {index}: dangling relationship {rid} on {node.tag}"


def _pictures(slide):
    return [shape for shape in slide.shapes if shape.shape_type == MSO_SHAPE_TYPE.PICTURE]


def _charts(slide):
    return [shape for shape in slide.shapes if shape.has_chart]


def _add_chart(service, pid, slide_index=1):
    service.dispatch(
        "pptx_add_chart",
        {
            "presentation_id": pid,
            "slide_index": slide_index,
            "chart_type": "column_clustered",
            "categories": ["A", "B"],
            "series": [{"name": "S1", "values": [1, 2]}],
        },
    )


def test_duplicate_slide_keeps_images_readable(service, image_path, tmp_path):
    pid = _new_deck(service)
    service.dispatch("pptx_add_image", {"presentation_id": pid, "slide_index": 1, "image_path": image_path})

    result = service.dispatch("pptx_duplicate_slide", {"presentation_id": pid, "source_index": 1})

    assert result["duplicated_slide_index"] == 2
    prs = _save_and_reopen(service, pid, tmp_path)
    _assert_relationships_resolve(prs)
    source_blob = _pictures(prs.slides[0])[0].image.blob
    assert _pictures(prs.slides[1])[0].image.blob == source_blob


def test_duplicate_slide_copies_placeholders_text_and_background(service, tmp_path):
    pid = _new_deck(service)
    service.dispatch(
        "pptx_set_placeholder_text",
        {"presentation_id": pid, "slide_index": 1, "placeholder_name": "Title 1", "text_content": "Quarterly results"},
    )
    service.dispatch("pptx_set_slide_background", {"presentation_id": pid, "slide_index": 1, "color_hex": "112233"})

    service.dispatch("pptx_duplicate_slide", {"presentation_id": pid, "source_index": 1})

    prs = _save_and_reopen(service, pid, tmp_path)
    source, duplicate = prs.slides[0], prs.slides[1]
    assert duplicate.shapes.title.text == "Quarterly results"
    assert [shape.name for shape in duplicate.shapes] == [shape.name for shape in source.shapes]
    assert duplicate._element.cSld.find(qn("p:bg")) is not None


def test_duplicate_slide_gives_the_copy_its_own_chart(service, tmp_path):
    pid = _new_deck(service)
    _add_chart(service, pid)

    service.dispatch("pptx_duplicate_slide", {"presentation_id": pid, "source_index": 1})
    copied_shape = _charts(service.engine.sessions[pid].extra["prs"].slides[1])[0]
    service.dispatch(
        "pptx_update_chart_data",
        {
            "presentation_id": pid,
            "slide_index": 2,
            "shape_id": int(copied_shape.shape_id),
            "categories": ["A", "B"],
            "series": [{"name": "S1", "values": [30, 40]}],
        },
    )

    prs = _save_and_reopen(service, pid, tmp_path)
    _assert_relationships_resolve(prs)
    source_chart = _charts(prs.slides[0])[0].chart
    copied_chart = _charts(prs.slides[1])[0].chart
    assert source_chart.part.partname != copied_chart.part.partname
    assert list(source_chart.plots[0].series[0].values) == [1, 2]
    assert list(copied_chart.plots[0].series[0].values) == [30, 40]
    assert copied_chart.part.chart_workbook.xlsx_part is not None


def test_clone_shape_to_another_slide_keeps_image(service, image_path, tmp_path):
    pid = _new_deck(service, slide_count=2)
    service.dispatch("pptx_add_image", {"presentation_id": pid, "slide_index": 1, "image_path": image_path})
    picture_id = int(_pictures(service.engine.sessions[pid].extra["prs"].slides[0])[0].shape_id)

    result = service.dispatch(
        "pptx_clone_shape",
        {"presentation_id": pid, "slide_index": 1, "shape_id": picture_id, "target_slide_index": 2},
    )

    prs = _save_and_reopen(service, pid, tmp_path)
    _assert_relationships_resolve(prs)
    cloned = _pictures(prs.slides[1])[0]
    assert cloned.image.blob == _pictures(prs.slides[0])[0].image.blob
    assert result["cloned_shape_id"] == cloned.shape_id


def test_clone_shape_on_same_slide_gets_a_unique_id(service, tmp_path):
    pid = _new_deck(service)
    service.dispatch(
        "pptx_add_shape",
        {
            "presentation_id": pid,
            "slide_index": 1,
            "shape_type": "rectangle",
            "left": "1in",
            "top": "1in",
            "width": "1in",
            "height": "1in",
        },
    )
    slide = service.engine.sessions[pid].extra["prs"].slides[0]
    source_id = int(list(slide.shapes)[-1].shape_id)

    service.dispatch("pptx_clone_shape", {"presentation_id": pid, "slide_index": 1, "shape_id": source_id})

    prs = _save_and_reopen(service, pid, tmp_path)
    ids = [shape.shape_id for shape in prs.slides[0].shapes]
    assert len(ids) == len(set(ids))


def test_copy_shape_between_decks_copies_image_into_target(service, image_path, tmp_path):
    source_pid = _new_deck(service)
    target_pid = _new_deck(service)
    other_image = tmp_path / "blue.png"
    Image.new("RGB", (32, 32), "blue").save(other_image)
    # Both decks now own a /ppt/media/image1.png with different content.
    service.dispatch("pptx_add_image", {"presentation_id": source_pid, "slide_index": 1, "image_path": image_path})
    service.dispatch(
        "pptx_add_image", {"presentation_id": target_pid, "slide_index": 1, "image_path": str(other_image)}
    )
    source_picture = _pictures(service.engine.sessions[source_pid].extra["prs"].slides[0])[0]

    service.dispatch(
        "pptx_copy_shape_between_decks",
        {
            "source_presentation_id": source_pid,
            "source_slide_index": 1,
            "shape_id": int(source_picture.shape_id),
            "target_presentation_id": target_pid,
            "target_slide_index": 1,
        },
    )

    prs = _save_and_reopen(service, target_pid, tmp_path)
    _assert_relationships_resolve(prs)
    blobs = {picture.image.blob for picture in _pictures(prs.slides[0])}
    assert blobs == {source_picture.image.blob, other_image.read_bytes()}


def test_copy_chart_between_decks_copies_chart_and_workbook(service, tmp_path):
    source_pid = _new_deck(service)
    target_pid = _new_deck(service)
    _add_chart(service, source_pid)
    _add_chart(service, target_pid)
    source_chart = _charts(service.engine.sessions[source_pid].extra["prs"].slides[0])[0]

    service.dispatch(
        "pptx_copy_shape_between_decks",
        {
            "source_presentation_id": source_pid,
            "source_slide_index": 1,
            "shape_id": int(source_chart.shape_id),
            "target_presentation_id": target_pid,
            "target_slide_index": 1,
        },
    )

    prs = _save_and_reopen(service, target_pid, tmp_path)
    _assert_relationships_resolve(prs)
    charts = _charts(prs.slides[0])
    assert len(charts) == 2
    assert len({chart.chart.part.partname for chart in charts}) == 2
    for chart in charts:
        assert chart.chart.part.chart_workbook.xlsx_part is not None
