from unittest.mock import MagicMock

from checkers.content_checker import _DEFAULT_TEXT_PATTERNS, ContentChecker


def test_default_text_patterns_match():
    assert any(p.search("Click to add title") for p in _DEFAULT_TEXT_PATTERNS)
    assert any(p.search("click here to edit") for p in _DEFAULT_TEXT_PATTERNS)
    assert not any(p.search("Q3 Revenue Results") for p in _DEFAULT_TEXT_PATTERNS)


def _make_placeholder(name, text):
    ph = MagicMock()
    ph.name = name
    ph.placeholder_format = MagicMock()
    ph.placeholder_format.type = "TITLE"
    ph.has_text_frame = True
    ph.text_frame = MagicMock()
    ph.text_frame.text = text
    ph.shape_type = 14  # not picture
    return ph


def _make_prs(placeholders_per_slide):
    prs = MagicMock()
    slides = []
    for phs in placeholders_per_slide:
        slide = MagicMock()
        slide.placeholders = phs
        slide.shapes = phs
        slides.append(slide)
    prs.slides = slides
    return prs


def test_empty_placeholder_detected():
    ph = _make_placeholder("Title 1", "")
    prs = _make_prs([[ph]])
    result = ContentChecker().check(prs, check_empty=True, check_default_text=False)
    assert len(result["empty_placeholders"]) == 1


def test_default_text_detected():
    ph = _make_placeholder("Title 1", "Click to add title")
    prs = _make_prs([[ph]])
    result = ContentChecker().check(prs, check_empty=False, check_default_text=True)
    assert len(result["default_text_remaining"]) == 1


def test_real_content_no_issues():
    ph = _make_placeholder("Title 1", "Q3 2025 Revenue Summary")
    prs = _make_prs([[ph]])
    result = ContentChecker().check(prs, check_empty=True, check_default_text=True)
    assert result["summary"]["total_issues"] == 0
