import io

import openppt_action

from pptx import Presentation
from pptx.util import Pt

from openppt_action import (
    APPENDIX_MARKER,
    VERSION,
    _find_layout_by_name,
    _pick_outline,
    build_pptx,
    list_layouts,
    parse_outline,
)

SAMPLE = """Here's your deck:

<slide title="Q3 Results" subtitle="Revenue up 12% YoY"></slide>

<slide title="Key Wins">
<ul>
<li>Launched EU region</li>
<li>Churn down to 2.1%
<ul><li>Best quarter ever</li></ul>
</li>
</ul>
<notes>mention the EU launch timeline if asked</notes>
</slide>

<slide title="Next Quarter">
<ol>
<li>Ship mobile app</li>
<li>Hire 2 SREs</li>
</ol>
</slide>

<slide title="Team Photo">
<ul><li><img src="https://example.com/team.jpg" alt="Team offsite"></li></ul>
</slide>
"""

TEMPLATED = """<slide title="Roadmap" layout="Section Header"></slide>

<slide title="Details" layout="Two Content">
<ul><li>Point one</li><li>Point two</li></ul>
</slide>
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
        "runs": [],
        "table": [],
        "code": "",
        "columns": [],
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
    assert parse_outline("no slide tags here") == []


def test_parse_layout_attribute():
    slides = parse_outline(TEMPLATED)
    assert slides[0]["title"] == "Roadmap"
    assert slides[0]["layout"] == "Section Header"
    assert slides[1]["title"] == "Details"
    assert slides[1]["layout"] == "Two Content"
    assert slides[1]["bullets"] == [(0, "Point one"), (0, "Point two")]


def test_subtitle_attribute_works_on_any_slide():
    slides = parse_outline(
        '<slide title="A" subtitle="first"></slide><slide title="B" subtitle="second"><li>x</li></slide>'
    )
    assert slides[0]["subtitle"] == "first"
    assert slides[1]["subtitle"] == "second"
    assert slides[1]["bullets"] == [(0, "x")]


def test_unclosed_slide_tags_are_tolerated():
    """A model that forgets '</slide>' shouldn't lose the next slide."""
    slides = parse_outline(
        '<slide title="One"><li>a<slide title="Two"><li>b</slide>'
    )
    assert [s["title"] for s in slides] == ["One", "Two"]
    assert slides[0]["bullets"] == [(0, "a")]
    assert slides[1]["bullets"] == [(0, "b")]


def test_bare_text_inside_slide_becomes_a_bullet():
    """Text not wrapped in '<p>'/'<li>' shouldn't be silently dropped."""
    (slide,) = parse_outline('<slide title="Bare">Just some text</slide>')
    assert slide["bullets"] == [(0, "Just some text")]


def test_p_tag_is_a_bullet():
    (slide,) = parse_outline(
        '<slide title="T"><p>First paragraph</p><p>Second paragraph</p></slide>'
    )
    assert slide["bullets"] == [(0, "First paragraph"), (0, "Second paragraph")]


def test_inline_tags_flatten_to_plain_text():
    (slide,) = parse_outline(
        '<slide title="T"><ul><li><b>bold</b> <code>code</code> '
        '<a href="http://x">link</a></li></ul></slide>'
    )
    assert slide["bullets"] == [(0, "bold code link")]


def test_notes_and_blockquote_both_feed_notes():
    (slide,) = parse_outline(
        '<slide title="T"><li>a</li><blockquote>quoted note</blockquote>'
        "<notes>and this</notes></slide>"
    )
    assert slide["notes"] == "quoted note\nand this"
    assert slide["bullets"] == [(0, "a")]


def test_deeply_nested_bullets_cap_at_level_4():
    nested = "<slide title=\"T\">"
    for _ in range(7):
        nested += "<ul><li>"
    nested += "deep"
    nested += "</li></ul>" * 7
    nested += "</slide>"
    (slide,) = parse_outline(nested)
    assert slide["bullets"] == [(4, "deep")]


def test_ignores_appended_download_link():
    """Re-exporting a message that already has a download link shouldn't
    turn that link into a bogus trailing bullet."""
    appended = (
        SAMPLE
        + f"\n{APPENDIX_MARKER}\n📊 [Download deck.pptx](/api/v1/files/abc/content)"
        + "\n\nhttp://host/api/v1/files/abc/content\n"
    )
    assert parse_outline(appended) == parse_outline(SAMPLE)


def test_appendix_marker_stops_even_with_an_unclosed_slide():
    """The marker has to win even when the model's own last slide was never
    closed — otherwise openPPT's own appended text gets folded into it."""
    unclosed = '<slide title="Real"><li>x' + f"\n\n{APPENDIX_MARKER}\n<slide title=\"stray\"><li>y</slide>"
    slides = parse_outline(unclosed)
    assert len(slides) == 1
    assert slides[0]["title"] == "Real"
    assert slides[0]["bullets"] == [(0, "x")]


def test_never_raises():
    junk = [
        "",
        "\x00�🙂" * 100,
        "<" * 500,
        "<li>" * 5000,
        "<pre>\nunclosed",
        "<slide title='" + "x" * 2000,
        "```\nunclosed fence",
    ]
    for j in junk:
        assert isinstance(parse_outline(j), list)


def test_pick_outline_falls_back_when_the_clicked_message_is_empty():
    """Open WebUI sometimes hands the action a message with no content at all;
    reporting 'no outline' at a chat that plainly has one is the failure that
    started the v0.3.5 diagnostic-exported-as-a-deck chain."""
    messages = [
        {"id": "a", "role": "assistant", "content": SAMPLE},
        {"id": "b", "role": "user", "content": "export it"},
        {"id": "c", "role": "assistant", "content": ""},
    ]
    assert _pick_outline(messages, "c") == SAMPLE
    assert _pick_outline(messages, None) == SAMPLE
    assert _pick_outline([], None) == ""
    assert _pick_outline([{"id": "c", "role": "assistant", "content": ""}], "c") == ""


def test_pick_outline_strips_what_openppt_appended():
    """Exporting twice must not turn our own download link into a bullet."""
    content = (
        SAMPLE
        + f"\n\n{APPENDIX_MARKER}\n\n"
        + "📊 [Download deck.pptx](/api/v1/files/x/content)\n"
    )
    slides = parse_outline(_pick_outline([{"id": "a", "role": "assistant", "content": content}], "a"))
    assert len(slides) == 4
    assert slides[-1]["bullets"] == [
        (0, {"alt": "Team offsite", "image": "https://example.com/team.jpg"})
    ]


def test_build_pptx():
    data = build_pptx(parse_outline(SAMPLE))
    assert data[:2] == b"PK"  # .pptx is a zip


def test_build_pptx_skips_unreachable_image():
    slides = parse_outline(
        '<slide title="Photo"><li><img src="https://nonexistent.invalid/x.png" alt="x"></li></slide>'
    )
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
    slides = parse_outline('<slide title="Roadmap" layout="Section Header"><li>item</li></slide>')
    data = build_pptx(slides, template=template)
    out = Presentation(io.BytesIO(data))
    assert out.slides[0].slide_layout.name == "Section Header"


def test_body_bullets_land_in_a_placeholder():
    slides = parse_outline('<slide title="Ideas"><ul><li>first</li><li>second</li></ul></slide>')
    data = build_pptx(slides)
    out = Presentation(io.BytesIO(data))
    texts = [
        p.text
        for shape in out.slides[0].placeholders
        if shape.has_text_frame
        for p in shape.text_frame.paragraphs
    ]
    assert "first" in texts and "second" in texts


def test_long_title_shrinks_instead_of_overlapping_the_body():
    """python-pptx can't recompute PowerPoint's autofit, so a long title keeps
    the template's ~44pt and spills across the slide over the bullets."""
    long_title = "openPPT: nothing slide-shaped in that message. " * 4
    out = Presentation(
        io.BytesIO(build_pptx(parse_outline(f'<slide title="{long_title}"><li>a</li></slide>')))
    )
    frame = out.slides[0].shapes.title.text_frame
    assert frame.word_wrap is True
    assert frame.paragraphs[0].font.size.pt <= 24  # the template's is 44


def test_short_title_keeps_the_templates_own_size():
    out = Presentation(
        io.BytesIO(build_pptx(parse_outline('<slide title="Q3 Results"><li>a</li></slide>')))
    )
    # None = inherited from the layout; we must not restyle a title that fits
    assert out.slides[0].shapes.title.text_frame.paragraphs[0].font.size is None


def test_dense_bullet_list_shrinks():
    items = "".join(f"<li>point {i}</li>" for i in range(20))
    out = Presentation(
        io.BytesIO(build_pptx(parse_outline(f'<slide title="Ideas"><ul>{items}</ul></slide>')))
    )
    (body,) = [p for p in out.slides[0].placeholders if p.placeholder_format.idx == 1]
    assert len(body.text_frame.paragraphs) == 20
    assert all(p.font.size.pt <= 14 for p in body.text_frame.paragraphs)
    # the size is on the runs too: PowerPoint renders a run's own size and
    # ignores a paragraph-level default for runs that already exist
    assert all(r.font.size.pt <= 14 for p in body.text_frame.paragraphs for r in p.runs)


def test_sparse_bullet_list_stays_readable():
    """Three short bullets used to be set at 20pt, leaving the slide empty."""
    out = Presentation(io.BytesIO(build_pptx(parse_outline(
        '<slide title="T"><ul><li>One</li><li>Two</li><li>Three</li></ul></slide>'
    ))))
    (body,) = [p for p in out.slides[0].placeholders if p.placeholder_format.idx == 1]
    assert all(p.font.size.pt >= 24 for p in body.text_frame.paragraphs)


def _shapes_inside_slide(prs, slide):
    for shape in slide.shapes:
        assert shape.left >= 0 and shape.top >= 0, shape.name
        assert shape.left + shape.width <= prs.slide_width + 1, shape.name
        assert shape.top + shape.height <= prs.slide_height + 1, shape.name


def test_bullets_keep_their_width_next_to_a_code_block():
    """Shrinking the body by setting only its height wrote width 0, so every
    slide mixing bullets with a table or code lost its bullets."""
    out = Presentation(io.BytesIO(build_pptx(parse_outline(
        '<slide title="Deploy"><ul><li>Run the script</li></ul>'
        "<pre><code>make ship</code></pre></slide>"
    ))))
    slide = out.slides[0]
    (body,) = [p for p in slide.placeholders if p.placeholder_format.idx == 1]
    layout_body = [p for p in slide.slide_layout.placeholders if p.placeholder_format.idx == 1][0]
    assert body.width == layout_body.width and body.left == layout_body.left
    assert body.top == layout_body.top and 0 < body.height < layout_body.height
    (code,) = [s for s in slide.shapes if s.has_text_frame and "make ship" in s.text_frame.text]
    assert code.top >= body.top + body.height
    _shapes_inside_slide(out, slide)


def _layout_shapes_inside(prs):
    for owner in [prs.slide_master] + list(prs.slide_layouts):
        for shape in owner.shapes:
            if shape.width is None:
                continue
            assert shape.left + shape.width <= prs.slide_width + 1, (owner.name, shape.name)


def test_default_deck_is_16_9_with_layouts_rescaled():
    prs = Presentation(io.BytesIO(build_pptx(parse_outline(SAMPLE))))
    assert abs(prs.slide_width / prs.slide_height - 16 / 9) < 0.01
    _layout_shapes_inside(prs)
    # placeholders were widened with the slide, not left bunched on the left
    title = prs.slides[1].shapes.title
    assert title.width > prs.slide_width * 0.8
    for slide in prs.slides:
        _shapes_inside_slide(prs, slide)


def test_accent_color_reaches_the_theme():
    from pptx.opc.constants import RELATIONSHIP_TYPE as RT

    prs = Presentation(io.BytesIO(build_pptx(parse_outline(SAMPLE), accent="#0F766E")))
    theme = prs.slide_master.part.part_related_by(RT.THEME).blob.decode()
    assert '<a:accent1><a:srgbClr val="0F766E"/></a:accent1>' in theme
    # the title page uses it as its background
    cover = prs.slides[0].slide_layout
    assert str(cover.background.fill.fore_color.rgb) == "0F766E"
    # an invalid color falls back instead of failing the export
    bad = Presentation(io.BytesIO(build_pptx(parse_outline(SAMPLE), accent="not-a-color")))
    assert bad.slides


def test_template_decks_keep_their_own_look():
    template = _make_template_bytes()
    prs = Presentation(io.BytesIO(build_pptx(parse_outline(SAMPLE), template=template)))
    assert prs.slide_width == Presentation(io.BytesIO(template)).slide_width  # 4:3 kept
    assert "openPPT accent bar" not in [s.name for s in prs.slide_master.shapes]


def test_slide_numbers_on_every_slide_but_the_first():
    from pptx.enum.shapes import PP_PLACEHOLDER

    def numbers(prs):
        return [
            [ph.text_frame.text for ph in slide.placeholders
             if ph.placeholder_format.type == PP_PLACEHOLDER.SLIDE_NUMBER]
            for slide in prs.slides
        ]

    prs = Presentation(io.BytesIO(build_pptx(parse_outline(SAMPLE))))
    assert numbers(prs) == [[], ["2"], ["3"], ["4"]]
    # and not on a section divider either
    divided = Presentation(io.BytesIO(build_pptx(parse_outline(
        '<slide title="Deck"></slide><slide title="Part 2"></slide><slide title="X"><li>a</li></slide>'
    ))))
    assert numbers(divided) == [[], [], ["3"]]
    assert prs.slides[1].placeholders[12]._element.xpath(".//a:fld[@type='slidenum']")
    off = Presentation(io.BytesIO(build_pptx(parse_outline(SAMPLE), slide_numbers=False)))
    assert numbers(off) == [[], [], [], []]


def _png(w, h):
    from PIL import Image

    buf = io.BytesIO()
    Image.new("RGB", (w, h), (200, 30, 30)).save(buf, "PNG")
    return buf.getvalue()


def test_a_non_image_src_never_aborts_the_export():
    """An <img> pointing at a text file or HTML page used to raise out of
    add_picture and fail the whole deck."""
    import tempfile

    with tempfile.NamedTemporaryFile("w", suffix=".png", delete=False) as fh:
        fh.write("not an image")
    outline = f'<slide title="Photo"><img src="{fh.name}" alt="x"><li>caption</li></slide>'
    for local in (False, True):
        out = Presentation(io.BytesIO(
            build_pptx(parse_outline(outline), allow_local_images=local)
        ))
        assert not [s for s in out.slides[0].shapes if s.shape_type == 13]


def test_local_paths_only_load_when_allowed():
    """In Open WebUI a local path is the server's filesystem — the outline is
    model output, so it must not be able to pull arbitrary server files in."""
    import tempfile

    with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as fh:
        fh.write(_png(40, 20))
    outline = f'<slide title="Photo"><img src="{fh.name}" alt="x"></slide>'
    blocked = Presentation(io.BytesIO(build_pptx(parse_outline(outline))))
    assert not [s for s in blocked.slides[0].shapes if s.shape_type == 13]
    allowed = Presentation(io.BytesIO(build_pptx(parse_outline(outline), allow_local_images=True)))
    assert [s for s in allowed.slides[0].shapes if s.shape_type == 13]


def test_private_network_image_urls_are_not_fetched():
    from openppt_action import _load_image, _public_host

    assert not _public_host("127.0.0.1")
    assert not _public_host("169.254.169.254")  # cloud metadata endpoint
    assert not _public_host("10.0.0.5")
    assert not _public_host("localhost")
    assert _load_image("http://127.0.0.1:9/x.png") is None


def test_data_uri_and_prefetched_images_embed_side_by_side():
    import base64

    uri = "data:image/png;base64," + base64.b64encode(_png(400, 200)).decode()
    outline = (
        f'<slide title="Two"><img src="{uri}" alt="wide">'
        '<img src="/api/v1/files/abc/content" alt="tall"></slide>'
    )
    data = build_pptx(parse_outline(outline), images={"/api/v1/files/abc/content": _png(100, 300)})
    prs = Presentation(io.BytesIO(data))
    slide = prs.slides[0]
    pics = [s for s in slide.shapes if s.shape_type == 13]
    assert len(pics) == 2
    wide, tall = pics
    # aspect ratio kept
    assert abs(wide.width / wide.height - 2.0) < 0.02
    assert abs(tall.width / tall.height - 1 / 3) < 0.02
    # no overlap, all on the slide
    assert wide.left + wide.width <= tall.left or tall.left + tall.width <= wide.left
    _shapes_inside_slide(prs, slide)
    assert wide._element.nvPicPr.cNvPr.get("descr") == "wide"


def test_images_sit_beside_bullets_not_on_top_of_them():
    import base64

    uri = "data:image/png;base64," + base64.b64encode(_png(300, 300)).decode()
    prs = Presentation(io.BytesIO(build_pptx(parse_outline(
        f'<slide title="T"><ul><li>point one</li><li>point two</li></ul><img src="{uri}"></slide>'
    ))))
    slide = prs.slides[0]
    (pic,) = [s for s in slide.shapes if s.shape_type == 13]
    (body,) = [p for p in slide.placeholders if p.placeholder_format.idx == 1]
    assert body.left + body.width <= pic.left
    _shapes_inside_slide(prs, slide)


def test_image_only_slide_keeps_its_title():
    """_body_placeholder matched by Python id(), which never matches a fresh
    python-pptx proxy — so on a Title Only layout it returned the title as the
    'body', the images were laid over it, and the title was dropped."""
    import base64

    uri = "data:image/png;base64," + base64.b64encode(_png(300, 200)).decode()
    prs = Presentation(io.BytesIO(build_pptx(parse_outline(
        f'<slide title="Deck"></slide><slide title="Team"><img src="{uri}"></slide>'
    ))))
    slide = prs.slides[1]
    assert slide.shapes.title is not None and slide.shapes.title.text == "Team"
    (pic,) = [s for s in slide.shapes if s.shape_type == 13]
    assert pic.top >= slide.shapes.title.top + slide.shapes.title.height
    _shapes_inside_slide(prs, slide)


def test_content_slide_subtitle_is_kept():
    """It used to be written into the body, then overwritten by bullet one."""
    prs = Presentation(io.BytesIO(build_pptx(parse_outline(
        '<slide title="Intro"></slide>'
        '<slide title="T" subtitle="Why it matters"><li>one</li><li>two</li></slide>'
    ))))
    texts = [s.text_frame.text for s in prs.slides[1].shapes if s.has_text_frame]
    assert "Why it matters" in texts
    assert any("one" in t and "two" in t for t in texts)
    _shapes_inside_slide(prs, prs.slides[1])


def test_no_empty_placeholder_is_left_behind():
    """An unfilled placeholder shows as 'Click to add text' in PowerPoint."""
    prs = Presentation(io.BytesIO(build_pptx(parse_outline(
        '<slide title="Deck"></slide>'
        '<slide title="Next" layout="Two Content"><li>a</li></slide>'
        '<slide title="Section break"></slide>'
        '<slide title="Code only"><pre>x = 1</pre></slide>'
    ))))
    for slide in prs.slides:
        for ph in slide.placeholders:
            assert ph.text_frame.text.strip(), (slide.slide_layout.name, ph.name)


def test_empty_later_slide_becomes_a_section_divider():
    prs = Presentation(io.BytesIO(build_pptx(parse_outline(
        '<slide title="Deck"></slide>'
        '<slide title="Part two" subtitle="Rollout"></slide>'
        '<slide title="Body"><li>x</li></slide>'
    ))))
    divider = prs.slides[1]
    assert divider.slide_layout.name == "Section Header"
    texts = [s.text_frame.text for s in divider.shapes if s.has_text_frame]
    assert "Part two" in texts and "Rollout" in texts


def test_table_rows_size_to_their_text():
    """Rows used to be stretched to fill the whole content area."""
    prs = Presentation(io.BytesIO(build_pptx(parse_outline(TABLE_OUTLINE))))
    slide = prs.slides[0]
    (shape,) = [s for s in slide.shapes if s.has_table]
    rows = list(shape.table.rows)
    assert all(r.height < Pt(60) for r in rows)
    assert sum(c.width for c in shape.table.columns) == shape.width
    _shapes_inside_slide(prs, slide)


def test_find_layout_by_name_is_fuzzy():
    prs = Presentation()
    assert _find_layout_by_name(prs, "section header").name == "Section Header"
    assert _find_layout_by_name(prs, "TWO CONTENT").name == "Two Content"
    assert _find_layout_by_name(prs, "nope") is None


TABLE_OUTLINE = """<slide title="Q3 Numbers">
<p>revenue held</p>
<table>
<tr><th>Region</th><th>Rev</th><th>Growth</th></tr>
<tr><td>EMEA</td><td>4.2</td><td>12%</td></tr>
<tr><td>APAC</td><td>3.1</td><td>30%</td></tr>
</table>
</slide>
"""


COLUMNS_OUTLINE = """<slide title="Build vs buy">
<p>Decision due Friday</p>
<column heading="Build"><ul><li>Full control</li><li>6 months</li></ul></column>
<column heading="Buy"><ul><li><b>Live in 2 weeks</b></li><li>$40k/yr</li></ul></column>
</slide>"""


def test_parse_columns():
    (slide,) = parse_outline(COLUMNS_OUTLINE)
    assert slide["bullets"] == [(0, "Decision due Friday")]
    assert [c["heading"] for c in slide["columns"]] == ["Build", "Buy"]
    assert slide["columns"][0]["bullets"] == [(0, "Full control"), (0, "6 months")]
    assert slide["columns"][1]["bullets"] == [(0, "Live in 2 weeks"), (0, "$40k/yr")]
    assert slide["columns"][1]["runs"][0] == [("Live in 2 weeks", ("b",))]


def test_unclosed_column_is_closed_by_the_next_one():
    (slide,) = parse_outline(
        '<slide title="T"><col heading="A"><li>a1<col heading="B"><li>b1</slide>'
    )
    assert [c["bullets"] for c in slide["columns"]] == [[(0, "a1")], [(0, "b1")]]


def test_col_inside_a_table_is_not_a_column():
    (slide,) = parse_outline(
        '<slide title="T"><table><colgroup><col></colgroup>'
        "<tr><td>x</td></tr></table></slide>"
    )
    assert slide["columns"] == [] and slide["table"] == [["x"]]


def test_build_renders_columns_into_the_two_content_layout():
    prs = Presentation(io.BytesIO(build_pptx(parse_outline(COLUMNS_OUTLINE))))
    slide = prs.slides[0]
    assert slide.slide_layout.name == "Two Content"
    bodies = sorted(
        (p for p in slide.placeholders if p.placeholder_format.type in (2, 7)),
        key=lambda p: p.left,
    )
    assert [b.text_frame.paragraphs[0].text for b in bodies] == ["Build", "Buy"]
    assert bodies[0].left + bodies[0].width <= bodies[1].left
    assert "Full control" in bodies[0].text_frame.text
    assert "$40k/yr" in bodies[1].text_frame.text
    # the slide-level line sits above both columns
    (intro,) = [s for s in slide.shapes if s.has_text_frame and s.text_frame.text == "Decision due Friday"]
    assert intro.top + intro.height <= bodies[0].top
    # headings are bold and bullet-less
    heading = bodies[0].text_frame.paragraphs[0]
    assert heading.runs[0].font.bold and heading._p.pPr.find(
        "{http://schemas.openxmlformats.org/drawingml/2006/main}buNone"
    ) is not None
    _shapes_inside_slide(prs, slide)


def test_columns_fall_back_to_text_boxes_without_a_matching_layout():
    three = (
        '<slide title="T"><column heading="A"><li>a</li></column>'
        '<column heading="B"><li>b</li></column><column heading="C"><li>c</li></column></slide>'
    )
    prs = Presentation(io.BytesIO(build_pptx(parse_outline(three))))
    slide = prs.slides[0]
    texts = [s.text_frame.text for s in slide.shapes if s.has_text_frame]
    for col in ("A\na", "B\nb", "C\nc"):
        assert col in texts, texts
    _shapes_inside_slide(prs, slide)


def test_inline_formatting_survives_as_runs():
    (slide,) = parse_outline(
        '<slide title="T"><ul><li>Ship <b>billing v2</b> by <i>March</i>, run <code>make ship</code></li>'
        "<li>plain</li></ul></slide>"
    )
    assert slide["bullets"] == [(0, "Ship billing v2 by March, run make ship"), (0, "plain")]
    assert slide["runs"][1] is None
    prs = Presentation(io.BytesIO(build_pptx([slide])))
    (body,) = [p for p in prs.slides[0].placeholders if p.placeholder_format.idx == 1]
    runs = {r.text: r for r in body.text_frame.paragraphs[0].runs}
    assert runs["billing v2"].font.bold and not runs["Ship "].font.bold
    assert runs["March"].font.italic
    assert runs["make ship"].font.name == "Consolas"


def test_bare_lines_become_separate_bullets():
    (slide,) = parse_outline('<slide title="T">\nFirst point\nSecond <b>point</b>\n</slide>')
    assert slide["bullets"] == [(0, "First point"), (0, "Second point")]


def test_line_breaks_inside_a_bullet_are_spaces():
    (slide,) = parse_outline('<slide title="T"><li>wrapped\n   across lines<br>and a br</li></slide>')
    assert slide["bullets"] == [(0, "wrapped across lines and a br")]


def test_parse_table():
    (slide,) = parse_outline(TABLE_OUTLINE)
    assert slide["table"] == [
        ["Region", "Rev", "Growth"],
        ["EMEA", "4.2", "12%"],
        ["APAC", "3.1", "30%"],
    ]
    assert slide["bullets"] == [(0, "revenue held")]  # the rows aren't bullets


def test_build_renders_table():
    out = Presentation(io.BytesIO(build_pptx(parse_outline(TABLE_OUTLINE))))
    (table,) = [s.table for s in out.slides[0].shapes if s.has_table]
    assert len(table.rows) == 3 and len(table.columns) == 3
    assert table.cell(2, 0).text == "APAC"


def test_parse_pre_code_block():
    (slide,) = parse_outline(
        '<slide title="Deploy"><li>one step</li><pre><code>make ship\nssh box</code></pre></slide>'
    )
    assert slide["code"] == "make ship\nssh box"
    assert slide["bullets"] == [(0, "one step")]


def test_parse_fenced_code_block_inside_a_slide():
    """Models fall back to ``` out of habit even when told to write HTML."""
    (slide,) = parse_outline('<slide title="Deploy">\n```bash\nmake ship\nssh box\n```\n</slide>')
    assert slide["code"] == "make ship\nssh box"


def test_code_with_angle_brackets_does_not_confuse_the_parser():
    """A ``` fence is HTML-escaped before parsing, so code containing '<'/'>'
    (e.g. a real HTML/XML sample) can't be mistaken for slide markup."""
    (slide,) = parse_outline(
        '<slide title="XML">\n```xml\n<a href="x"><b>hi</b></a>\n```\n</slide>'
    )
    assert slide["code"] == '<a href="x"><b>hi</b></a>'
    assert slide["bullets"] == []


def test_build_renders_code_in_a_monospace_box():
    outline = '<slide title="Deploy"><pre><code>print(\'hi\')</code></pre></slide>'
    out = Presentation(io.BytesIO(build_pptx(parse_outline(outline))))
    boxes = [
        s for s in out.slides[0].shapes
        if s.has_text_frame and "print('hi')" in s.text_frame.text
    ]
    (box,) = boxes
    assert box.text_frame.paragraphs[0].font.name == "Consolas"
    # the empty bullet placeholder is gone, not left as a "click to add" box
    assert not [p for p in out.slides[0].placeholders if p.placeholder_format.idx == 1]


def test_outline_wrapped_in_a_fence_still_parses():
    """Models wrap their whole answer in ``` whatever the prompt says; that
    fence must not swallow the deck into one slide's code block."""
    slides = parse_outline(
        'Here you go:\n\n```html\n<slide title="One"><li>a</li></slide>\n'
        '<slide title="Two"><li>b</li></slide>\n```\n'
    )
    assert [s["title"] for s in slides] == ["One", "Two"]
    assert all(not s["code"] for s in slides)


def test_bullets_and_a_table_share_the_slide():
    out = Presentation(io.BytesIO(build_pptx(parse_outline(TABLE_OUTLINE))))
    slide = out.slides[0]
    (body,) = [p for p in slide.placeholders if p.placeholder_format.idx == 1]
    (table,) = [s for s in slide.shapes if s.has_table]
    assert "revenue held" in body.text_frame.text
    assert table.top >= body.top + body.height  # table sits below the bullets


def test_list_layouts_reports_names():
    names = [n for n, _ in list_layouts(None)]
    assert "Title Slide" in names and "Title and Content" in names


class _FakeRequest:
    base_url = "http://host.example/"


def _run_action(async_api: bool, content: str = SAMPLE, __request__=None,
                body_content=None, history=False, event_call=None, valves=None,
                attach=True, user=None, extra_files=None):
    """Drive Action.action against stand-in Open WebUI Files/Storage modules.

    Open WebUI's Files API is sync in older versions and async in current ones,
    so both shapes have to work. The async fakes only record their effect
    *inside* the coroutine — a dropped 'await' therefore shows up as a missing
    file record rather than passing silently.

    Whatever the action *returns* is how it edits the message on current Open
    WebUI, so each returned message is recorded as a 'returned' pseudo-event
    alongside the emitted ones — tests scanning event content see both.
    """
    import asyncio
    import sys
    import tempfile
    import types

    from openppt_action import Action

    with tempfile.NamedTemporaryFile(suffix=".pptx", delete=False) as fh:
        fh.write(_make_template_bytes())
        template_path = fh.name

    inserted = {}
    uploaded = _run_action.uploaded = []
    owners = {"f1": "u1", **{fid: owner for fid, (owner, _) in (extra_files or {}).items()}}

    class Record:
        def __init__(self, fid):
            self.filename, self.path = "brand.pptx", template_path
            if fid in (extra_files or {}):
                self.filename, self.path = "image.png", extra_files[fid][1]

    class Files:
        @staticmethod
        def insert_new_file(user_id, form):
            async def _insert():
                inserted[form["id"]] = form
                return form

            if async_api:
                return _insert()
            inserted[form["id"]] = form
            return form

        @staticmethod
        def get_file_by_id(fid):
            async def _get():
                return Record(fid)

            return _get() if async_api else Record(fid)

        @staticmethod
        def get_file_by_id_and_user_id(fid, user_id):
            rec = Record(fid) if owners.get(fid) == user_id else None

            async def _get():
                return rec

            return _get() if async_api else rec

    class Storage:
        @staticmethod
        def upload_file(stream, filename, tags=None):
            data = stream.read()
            uploaded.append(data)
            return data, f"/data/{filename}"

        @staticmethod
        def get_file(path):
            return path

    stored_messages = [
        {"id": "m0", "role": "user", "content": "deck please",
         "files": [{"type": "file", "id": "f1", "name": "brand.pptx"}] if attach else []},
        {"id": "m1", "role": "assistant", "content": content},
    ]

    class Chat:
        # what Open WebUI persists: the attachment lives on the user message.
        # Some versions keep the flat list, others only the history id->msg map.
        chat = (
            {"history": {"messages": {m["id"]: m for m in stored_messages}}}
            if history
            else {"messages": stored_messages}
        )

    class Chats:
        @staticmethod
        def get_chat_by_id(cid):
            async def _get():
                return Chat()

            return _get() if async_api else Chat()

    files_mod = types.ModuleType("open_webui.models.files")
    files_mod.Files = Files
    files_mod.FileForm = lambda **kw: kw
    chats_mod = types.ModuleType("open_webui.models.chats")
    chats_mod.Chats = Chats
    storage_mod = types.ModuleType("open_webui.storage.provider")
    storage_mod.Storage = Storage
    fakes = {
        "open_webui": types.ModuleType("open_webui"),
        "open_webui.models": types.ModuleType("open_webui.models"),
        "open_webui.models.files": files_mod,
        "open_webui.models.chats": chats_mod,
        "open_webui.storage": types.ModuleType("open_webui.storage"),
        "open_webui.storage.provider": storage_mod,
    }

    events = []

    async def emitter(event):
        events.append(event)

    # Exactly what Open WebUI POSTs to /api/chat/actions: no 'files' anywhere,
    # each message stripped to id/role/content. The template has to be found
    # through chat_id, or every export silently falls back to the plain theme.
    body = {
        "id": "m1",
        "chat_id": "c1",
        "messages": [
            {"id": "m0", "role": "user", "content": "deck please"},
            {"id": "m1", "role": "assistant",
             "content": content if body_content is None else body_content},
        ],
    }
    action = Action()
    for key, value in (valves or {}).items():
        setattr(action.valves, key, value)
    sys.modules.update(fakes)
    try:
        result = asyncio.run(
            action.action(
                body,
                __user__=user or {"id": "u1", "role": "user"},
                __event_emitter__=emitter,
                __event_call__=event_call,
                __request__=__request__,
            )
        )
    finally:
        for key in fakes:
            sys.modules.pop(key, None)
    for msg in (result or {}).get("messages", []):
        events.append({"type": "returned", "data": dict(msg)})
    return events, inserted


def test_action_inserts_the_file_and_links_to_it():
    for async_api in (False, True):
        events, inserted = _run_action(async_api)
        descriptions = " ".join(
            str(e["data"].get("description", "")) + str(e["data"].get("content", ""))
            for e in events
        )
        assert "failed" not in descriptions, descriptions
        # the record has to exist, or the link 404s under a "ready" toast
        (file_id,) = inserted
        assert inserted[file_id]["meta"]["size"] > 0
        link = f"/api/v1/files/{file_id}/content"
        assert any(link in str(e["data"].get("content", "")) for e in events), events
        # and it's offered as a real attachment, not only as markdown
        (attached,) = [e for e in events if e["type"] == "files"]
        (chip,) = attached["data"]["files"]
        assert chip["id"] == file_id and chip["url"] == file_id
        assert chip["type"] == "file" and chip["name"].endswith(".pptx")
        # and the attached .pptx was picked up as the template
        assert "on your template" in descriptions


def test_no_outline_diagnostic_is_not_itself_slide_shaped():
    """v0.3.5 posted an error report written in '---' + '- ' Markdown, so the
    next click exported the error report as a deck — the diagnostic must
    still contain no '<slide>' tag under the new grammar."""
    events, inserted = _run_action(False, content="Sorry, I can't help with that.")
    assert inserted == {}  # nothing was built
    (message,) = [e for e in events if e["type"] in ("message", "returned")]
    posted = message["data"]["content"]
    assert APPENDIX_MARKER in posted
    assert parse_outline(posted.split(APPENDIX_MARKER, 1)[1]) == []
    # the model's own reply is kept, with the diagnostic under it
    assert posted.startswith("Sorry, I can't help with that.")
    # and once it's in the message, it is invisible to the next click
    appended = {"id": "m1", "role": "assistant", "content": posted}
    assert _pick_outline([appended], "m1") == "Sorry, I can't help with that."


def test_no_outline_toast_carries_the_evidence():
    """Some deployments surface only the toast — the appended chat block goes
    unseen — so the toast itself has to say which paste ran and how much text
    the action was actually handed. '0 chars' (the action got nothing) and
    'N chars' (the parser rejected real text) are different root causes, and
    without them in the toast there is no way to tell them apart remotely."""
    for content, expect_chars in (("", 0), ("Sorry, I can't help with that.", 30)):
        events, inserted = _run_action(False, content=content)
        assert inserted == {}  # nothing was built
        toasts = [
            str(e["data"].get("content", "")) + str(e["data"].get("description", ""))
            for e in events
            if e["type"] in ("notification", "status")
        ]
        assert toasts, events
        blob = " ".join(toasts)
        assert VERSION in blob, blob  # a stale paste shows up here
        assert f"read {expect_chars} chars" in blob, blob
    # and the empty case names itself rather than trailing off into nothing
    assert "(empty)" in blob or "(empty)" in " ".join(
        str(e["data"].get("content", ""))
        for e in _run_action(False, content="")[0]
        if e["type"] == "notification"
    )


def test_outline_is_read_from_the_saved_chat_when_the_body_is_blank():
    """Some Open WebUI versions POST the action a body whose assistant message
    content is an empty string — the toast reported exactly
    'read 0 chars from 2 messages [user:str(12), assistant:str(0)]'.
    parse_outline is pure, so with no text there is nothing to parse and every
    export fails on that deployment while an identical chat exports fine on a
    version that populates the body. The outline has to come from the stored
    chat, the same way v0.5.1 had to go there for the template.
    """
    for hist in (False, True):
        events, inserted = _run_action(False, body_content="", history=hist)
        blob = " ".join(
            str(e["data"].get("description", "")) + str(e["data"].get("content", ""))
            for e in events
        )
        assert "nothing slide-shaped" not in blob, blob
        (file_id,) = inserted  # a real deck was built from the stored outline
        assert inserted[file_id]["meta"]["size"] > 0
        assert "on your template" in blob  # and the template still resolved

    # the body still wins when it does carry the outline: no needless chat read
    events, inserted = _run_action(False)
    assert inserted and "nothing slide-shaped" not in " ".join(
        str(e["data"].get("content", "")) for e in events
    )
    # and a genuinely outline-less chat still reports, rather than hanging on
    # a fallback that can't help
    events, inserted = _run_action(False, content="Sorry, I can't help.", body_content="")
    assert inserted == {}
    assert any("nothing slide-shaped" in str(e["data"].get("content", "")) for e in events)


def test_shape_separates_the_ways_content_goes_missing():
    """'read 0 chars' has four causes needing four different fixes. _shape has
    to tell them apart from the toast alone, with no access to the machine."""
    from openppt_action import _shape

    # list-shaped (multimodal) content — _pick_outline's isinstance guard
    # turns this into "" and the outline is lost
    assert "assistant:list(1)" in _shape(
        [{"role": "assistant", "content": [{"type": "text", "text": "<slide>"}]}]
    )
    # genuinely empty content in the POST body
    assert "assistant:str(0)" in _shape([{"role": "assistant", "content": ""}])
    # no assistant role at all — candidates never gets an entry
    assert _shape([{"role": "user", "content": "hi"}]) == "user:str(2)"
    # content absent entirely
    assert "assistant:NoneType(?)" in _shape([{"role": "assistant"}])
    # and it never raises on junk, since it runs inside the failure path
    assert _shape(None) == "(none)"
    assert _shape([]) == "(none)"
    assert "str!" in _shape(["not-a-dict"])


def test_action_appends_full_url_when_request_available():
    events, inserted = _run_action(False, __request__=_FakeRequest())
    (file_id,) = inserted
    full_link = f"http://host.example/api/v1/files/{file_id}/content"
    contents = [str(e["data"].get("content", "")) for e in events]
    (download_msg,) = [c for c in contents if "📊" in c]
    assert full_link in download_msg
    assert APPENDIX_MARKER in download_msg
    # the full URL is repeated as its own line, after everything else
    assert download_msg.rstrip().endswith(full_link)


def _statuses(events):
    return " ".join(str(e["data"].get("description", "")) for e in events if e["type"] == "status")


def test_template_precedence_attached_then_valve_then_builtin():
    import tempfile

    wide = Presentation()
    wide.slide_width = 12192000
    buf = io.BytesIO()
    wide.save(buf)
    with tempfile.NamedTemporaryFile(suffix=".pptx", delete=False) as fh:
        fh.write(buf.getvalue())
    valve = {"default_template": fh.name}

    def deck():
        return Presentation(io.BytesIO(_run_action.uploaded[-1]))

    # attached template wins over the Valve (the attachment is 4:3)
    events, _ = _run_action(False, valves=valve)
    assert "on your template" in _statuses(events)
    assert deck().slide_width == 9144000
    # no attachment: the admin's default template
    events, _ = _run_action(False, valves=valve, attach=False)
    assert "on the default template" in _statuses(events)
    assert deck().slide_width == 12192000
    assert "openPPT accent bar" not in [s.name for s in deck().slide_master.shapes]
    # neither: the built-in theme, in the Valve's accent color
    events, _ = _run_action(False, attach=False, valves={"accent_color": "0F766E"})
    assert "template" not in _statuses(events)
    assert "openPPT accent bar" in [s.name for s in deck().slide_master.shapes]
    # a broken Valve setting falls back rather than failing the export
    events, inserted = _run_action(False, attach=False, valves={"default_template": "/nope.pptx"})
    assert inserted and "failed" not in _statuses(events)


def test_open_webui_image_links_are_read_from_storage_for_their_owner_only():
    import tempfile

    with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as fh:
        fh.write(_png(60, 30))
    outline = (
        '<slide title="Pics"><img src="/api/v1/files/mine/content" alt="mine">'
        '<img src="http://host.example/api/v1/files/theirs/content" alt="theirs"></slide>'
    )
    files = {"mine": ("u1", fh.name), "theirs": ("u2", fh.name)}

    def pictures():
        prs = Presentation(io.BytesIO(_run_action.uploaded[-1]))
        return [s for s in prs.slides[0].shapes if s.shape_type == 13]

    _run_action(True, content=outline, extra_files=files)
    assert len(pictures()) == 1  # someone else's file isn't readable
    _run_action(True, content=outline, extra_files=files, user={"id": "u1", "role": "admin"})
    assert len(pictures()) == 2  # an admin can reach any file, as in the UI


def test_history_wins_when_the_flat_message_list_lags_behind():
    # Open WebUI's saved chat keeps the conversation twice; the flat list is
    # rebuilt on save and can hold only the user turn while history already has
    # the assistant's reply. Reading the emptier store loses the outline.
    chat = {
        "messages": [{"role": "user", "content": "make a deck"}],
        "history": {
            "messages": {
                "a": {"role": "user", "content": "make a deck"},
                "b": {"role": "assistant", "content": "<slide title='X'></slide>"},
            }
        },
    }
    picked = openppt_action._chat_messages(chat)
    assert len(picked) == 2, picked
    assert openppt_action.parse_outline(openppt_action._pick_outline(picked))

    # ...but a complete flat list still wins over an empty history.
    flat_only = {"messages": [{"role": "assistant", "content": "hi"}], "history": {}}
    assert len(openppt_action._chat_messages(flat_only)) == 1
    assert openppt_action._chat_messages({}) == []
    assert openppt_action._saved_shape(chat) == "flat=1 hist=2"
    assert openppt_action._saved_shape(None) == "none"


def test_missing_assistant_reply_is_not_reported_as_a_bad_outline():
    # flat=1 hist=1 in the wild: the chat holds only the user's turn, so there
    # is no reply to have an outline in. Blaming the outline format here sends
    # people off rewriting a message that was never the problem.
    import asyncio

    events = []

    async def emitter(event):
        events.append(event)

    body = {"chat_id": "c1", "id": "m2", "messages": [{"role": "user", "content": "deck pls"}]}
    action = openppt_action.Action()
    action._saved_chat = lambda b: _async({"messages": [{"role": "user", "content": "deck pls"}]})
    asyncio.run(action.action(body, {}, emitter))
    text = " ".join(
        str(e.get("data", {}).get("content", "")) for e in events
    )
    assert "never received the model's reply" in text, text
    assert "no assistant reply reached openPPT" in text, text
    assert "nothing slide-shaped" not in text, text


async def _async(value):
    return value


def _returned(events):
    return [e["data"]["content"] for e in events if e["type"] == "returned"]


def test_export_edits_the_message_through_the_return_value():
    """v0.10+ renders a reply from its structured 'output', not 'content', so
    a 'message' event (which only appends to 'content') put the download link
    where nobody could see it. Returning the edited message is Open WebUI's
    own contract, and the frontend syncs it into 'output'."""
    events, inserted = _run_action(False, __request__=_FakeRequest())
    (file_id,) = inserted
    (edited,) = _returned(events)
    # the model's outline is kept intact, the link sits under it
    assert edited.startswith(SAMPLE.rstrip())
    assert f"/api/v1/files/{file_id}/content" in edited
    assert edited.count(APPENDIX_MARKER) == 1
    # and nothing is double-appended through the legacy event as well
    assert not [e for e in events if e["type"] == "message"]


def test_reexport_replaces_the_old_link_instead_of_stacking():
    previous = SAMPLE + f"\n\n{APPENDIX_MARKER}\n📊 [Download old.pptx](/api/v1/files/old/content)"
    events, inserted = _run_action(False, content=previous)
    (edited,) = _returned(events)
    assert edited.count(APPENDIX_MARKER) == 1
    assert "old.pptx" not in edited


def test_empty_clicked_text_never_overwrites_the_reply():
    """With no clicked text in the request, returning a message would replace
    the model's reply with just openPPT's link — so fall back to appending."""
    events, inserted = _run_action(False, body_content="")
    assert inserted  # the outline still came from the saved chat
    assert _returned(events) == []
    (appended,) = [e for e in events if e["type"] == "message"]
    assert appended["data"]["content"].lstrip().startswith(APPENDIX_MARKER)


def test_clicked_text_is_read_from_structured_output():
    """v0.10+ keeps the reply in 'output' items; 'content' can be empty."""
    from openppt_action import _message_text

    msg = {
        "role": "assistant",
        "content": "",
        "output": [
            {"type": "reasoning", "content": [{"type": "text", "text": "thinking"}]},
            {"type": "message", "content": [{"type": "output_text", "text": SAMPLE}]},
        ],
    }
    assert _message_text(msg) == SAMPLE
    assert len(parse_outline(_pick_outline([dict(msg, id="m1")], "m1"))) == 4
    # multimodal list content
    listed = {"role": "assistant", "content": [{"type": "text", "text": SAMPLE}]}
    assert _message_text(listed) == SAMPLE
    assert _message_text({"role": "assistant", "content": None}) == ""
    assert _message_text("junk") == ""


def test_auto_download_runs_in_the_browser():
    calls = []

    async def event_call(event):
        calls.append(event)
        return "downloaded"

    events, inserted = _run_action(False, event_call=event_call, __request__=_FakeRequest())
    (file_id,) = inserted
    (call,) = calls
    assert call["type"] == "execute"
    code = call["data"]["code"]
    assert f'"/api/v1/files/{file_id}/content"' in code  # relative URL tried first
    assert f"http://host.example/api/v1/files/{file_id}/content" in code
    assert "localStorage.token" in code and ".download = name" in code
    statuses = [e["data"].get("description", "") for e in events if e["type"] == "status"]
    assert any("downloaded" in s for s in statuses), statuses


def test_hanging_auto_download_does_not_block_the_export():
    import asyncio

    async def event_call(event):
        await asyncio.sleep(30)  # a frontend that never answers

    saved = openppt_action.AUTO_DOWNLOAD_TIMEOUT
    openppt_action.AUTO_DOWNLOAD_TIMEOUT = 0.05
    try:
        events, inserted = _run_action(False, event_call=event_call)
    finally:
        openppt_action.AUTO_DOWNLOAD_TIMEOUT = saved
    (file_id,) = inserted
    assert any(file_id in c for c in _returned(events))  # the link still lands
    assert not any("failed" in str(e["data"]) for e in events)


def test_auto_download_valve_off_makes_no_browser_call():
    calls = []

    async def event_call(event):
        calls.append(event)
        return "downloaded"

    events, inserted = _run_action(False, event_call=event_call, valves={"auto_download": False})
    assert inserted and calls == []


# Kept last: anything defined below this block would be skipped when the file
# is run standalone (two tests silently were, before v0.8.0).
if __name__ == "__main__":
    for _name, _fn in sorted(list(globals().items())):
        if _name.startswith("test_") and callable(_fn):
            _fn()
    print("ok")
