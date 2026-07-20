import io

from pptx import Presentation

from openppt_action import (
    _find_layout_by_name,
    build_pptx,
    list_layouts,
    parse_outline,
)

SAMPLE = """Here's your deck:

# Q3 Results
## Revenue up 12% YoY

# Key Wins
- Launched EU region
- Churn down to 2.1%
  - Best quarter ever
Notes: mention the EU launch timeline if asked

# Next Quarter
1. Ship mobile app
* Hire 2 SREs

# Team Photo
- ![Team offsite](https://example.com/team.jpg)
"""

TEMPLATED = """# Roadmap @layout: Section Header

# Details {layout: Two Content}
- Point one
- Point two

# Plain slide
Layout: Title and Content
- A bullet
"""


def test_parse_outline():
    slides = parse_outline(SAMPLE)
    assert len(slides) == 4
    assert slides[0] == {
        "title": "Q3 Results",
        "subtitle": "Revenue up 12% YoY",
        "layout": "",
        "notes": "",
        "bullets": [],
    }
    assert slides[1]["bullets"] == [
        (0, "Launched EU region"),
        (0, "Churn down to 2.1%"),
        (1, "Best quarter ever"),
    ]
    assert slides[1]["notes"] == "mention the EU launch timeline if asked"
    assert slides[2]["bullets"] == [(0, "Ship mobile app"), (0, "Hire 2 SREs")]
    assert slides[3]["bullets"] == [
        (0, {"alt": "Team offsite", "image": "https://example.com/team.jpg"})
    ]
    assert parse_outline("no headings here") == []


def test_parse_layout_directives():
    slides = parse_outline(TEMPLATED)
    # inline '@layout:', inline '{layout: ...}', and a standalone 'Layout:' line
    assert slides[0]["title"] == "Roadmap"
    assert slides[0]["layout"] == "Section Header"
    assert slides[1]["title"] == "Details"
    assert slides[1]["layout"] == "Two Content"
    assert slides[1]["bullets"] == [(0, "Point one"), (0, "Point two")]
    assert slides[2]["title"] == "Plain slide"
    assert slides[2]["layout"] == "Title and Content"
    assert slides[2]["bullets"] == [(0, "A bullet")]


def test_build_pptx():
    data = build_pptx(parse_outline(SAMPLE))
    assert data[:2] == b"PK"  # .pptx is a zip


def test_build_pptx_skips_unreachable_image():
    # image slide with an unfetchable URL should still produce a valid deck
    slides = parse_outline("# Photo\n- ![x](https://nonexistent.invalid/x.png)\n")
    data = build_pptx(slides)
    assert data[:2] == b"PK"


def _make_template_bytes():
    """A 'template' is just an ordinary .pptx; build one and re-serialize it."""
    prs = Presentation()
    prs.slides.add_slide(prs.slide_layouts[6])  # a stray sample slide
    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()


def test_build_with_template_clears_sample_slides():
    template = _make_template_bytes()
    slides = parse_outline(SAMPLE)
    data = build_pptx(slides, template=template)
    out = Presentation(io.BytesIO(data))
    # 4 outline slides, and the template's stray sample slide is gone.
    assert len(out.slides) == 4
    assert out.slides[0].shapes.title.text == "Q3 Results"


def test_layout_directive_selects_named_layout():
    template = _make_template_bytes()
    slides = parse_outline("# Roadmap @layout: Section Header\n- item\n")
    data = build_pptx(slides, template=template)
    out = Presentation(io.BytesIO(data))
    assert out.slides[0].slide_layout.name == "Section Header"


def test_body_bullets_land_in_a_placeholder():
    slides = parse_outline("# Ideas\n- first\n- second\n")
    data = build_pptx(slides)
    out = Presentation(io.BytesIO(data))
    texts = [
        p.text
        for shape in out.slides[0].placeholders
        if shape.has_text_frame
        for p in shape.text_frame.paragraphs
    ]
    assert "first" in texts and "second" in texts


def test_find_layout_by_name_is_fuzzy():
    prs = Presentation()
    assert _find_layout_by_name(prs, "section header").name == "Section Header"
    assert _find_layout_by_name(prs, "TWO CONTENT").name == "Two Content"
    assert _find_layout_by_name(prs, "nope") is None


def test_list_layouts_reports_names():
    names = [n for n, _ in list_layouts(None)]
    assert "Title Slide" in names and "Title and Content" in names


if __name__ == "__main__":
    test_parse_outline()
    test_parse_layout_directives()
    test_build_pptx()
    test_build_pptx_skips_unreachable_image()
    test_build_with_template_clears_sample_slides()
    test_layout_directive_selects_named_layout()
    test_body_bullets_land_in_a_placeholder()
    test_find_layout_by_name_is_fuzzy()
    test_list_layouts_reports_names()
    print("ok")
