from unittest.mock import MagicMock

from checkers.position_checker import PositionChecker, _rects_overlap


def test_rects_overlap_clear_separation():
    assert not _rects_overlap((0, 0, 100, 100), (200, 0, 300, 100), tolerance=0)


def test_rects_overlap_touching_boundary():
    assert not _rects_overlap((0, 0, 100, 100), (100, 0, 200, 100), tolerance=0)


def test_rects_overlap_within_tolerance():
    assert not _rects_overlap((0, 0, 103, 100), (100, 0, 200, 100), tolerance=50000)


def test_rects_overlap_actual_overlap():
    assert _rects_overlap((0, 0, 200, 200), (100, 100, 300, 300), tolerance=0)


def _make_mock_shape(name, left, top, width, height):
    shape = MagicMock()
    shape.name = name
    shape.left = left
    shape.top = top
    shape.width = width
    shape.height = height
    return shape


def _make_mock_prs(slide_width, slide_height, shapes):
    prs = MagicMock()
    prs.slide_width = slide_width
    prs.slide_height = slide_height
    slide = MagicMock()
    slide.shapes = shapes
    prs.slides = [slide]
    return prs


def test_out_of_bounds_detection():
    sw, sh = 9144000, 6858000  # 10in x 7.5in
    shapes = [_make_mock_shape("box", 9000000, 0, 300000, 300000)]
    prs = _make_mock_prs(sw, sh, shapes)
    result = PositionChecker().check(prs, [1], check_bounds=True, check_overlaps=False, tolerance_px=0)
    assert len(result["issues"]) == 1
    assert result["issues"][0]["issue_type"] == "out_of_bounds"


def test_no_issues_for_contained_shapes():
    sw, sh = 9144000, 6858000
    shapes = [
        _make_mock_shape("box1", 0, 0, 914400, 914400),
        _make_mock_shape("box2", 2000000, 2000000, 914400, 914400),
    ]
    prs = _make_mock_prs(sw, sh, shapes)
    result = PositionChecker().check(prs, [1], check_bounds=True, check_overlaps=True, tolerance_px=5)
    assert result["summary"]["total_issues"] == 0
