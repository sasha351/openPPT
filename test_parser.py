import io

from pptx import Presentation

from openppt_action import (
    APPENDIX_MARKER,
    _find_layout_by_name,
    _pick_outline,
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


def test_h2_deck_slides_on_most_common_level():
    """A '## '-per-slide deck: '#' above it is the deck title, not a bullet."""
    slides = parse_outline("# Deck\n## One\n- a\n## Two\n- b\n## Three\n- c\n")
    assert [s["title"] for s in slides] == ["Deck", "One", "Two", "Three"]
    assert slides[1]["bullets"] == [(0, "a")]


def test_headingless_fallbacks():
    assert [s["title"] for s in parse_outline("Slide 1: Intro\n- a\nSlide 2: Body\n- b\n")] == [
        "Intro",
        "Body",
    ]
    assert [s["title"] for s in parse_outline("**Intro**\n- a\n\n**Body**\n- b\n")] == [
        "Intro",
        "Body",
    ]
    sep = parse_outline("Intro\n- a\n\n---\n\nBody\n- b\n")
    assert [s["title"] for s in sep] == ["Body"]  # prose before the first marker is dropped
    assert sep[0]["bullets"] == [(0, "b")]


def test_extra_bullet_markers_and_inline_markdown():
    (slide,) = parse_outline("# T\n+ plus\n2) paren\n- **bold** `code` [link](http://x)\n")
    assert slide["bullets"] == [
        (0, "plus"),
        (0, "paren"),
        (0, "bold code link"),
    ]


def test_blockquote_notes():
    (slide,) = parse_outline("# T\n- a\n> quoted note\nNotes: and this\n")
    assert slide["notes"] == "quoted note\nand this"
    assert slide["bullets"] == [(0, "a")]


def test_tolerance_does_not_break_layouts_or_images():
    """New grammar must not eat the v0.3.0 directives it runs alongside."""
    slides = parse_outline(
        "## Roadmap @layout: Section Header\n## Photo\n- ![alt](http://x/y.png)\n"
    )
    assert slides[0]["layout"] == "Section Header"
    assert slides[0]["title"] == "Roadmap"
    assert slides[1]["bullets"] == [(0, {"alt": "alt", "image": "http://x/y.png"})]


def test_ignores_appended_download_link():
    """Re-exporting a message that already has a download link shouldn't
    turn that link into a bogus trailing bullet."""
    appended = (
        SAMPLE
        + f"\n{APPENDIX_MARKER}\n📊 [Download deck.pptx](/api/v1/files/abc/content)"
        + "\n\nhttp://host/api/v1/files/abc/content\n"
    )
    slides = parse_outline(appended)
    assert slides == parse_outline(SAMPLE)


def test_never_raises():
    for junk in ["", "\x00�🙂" * 100, "#" * 500, "- \n" * 5000, "```\nunclosed"]:
        assert isinstance(parse_outline(junk), list)


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


def test_long_title_shrinks_instead_of_overlapping_the_body():
    """python-pptx can't recompute PowerPoint's autofit, so a long title keeps
    the template's ~44pt and spills across the slide over the bullets."""
    long_title = "openPPT: nothing slide-shaped in that message. " * 4
    out = Presentation(io.BytesIO(build_pptx(parse_outline(f"# {long_title}\n- a\n"))))
    frame = out.slides[0].shapes.title.text_frame
    assert frame.word_wrap is True
    assert frame.paragraphs[0].font.size.pt <= 20


def test_short_title_keeps_the_templates_own_size():
    out = Presentation(io.BytesIO(build_pptx(parse_outline("# Q3 Results\n- a\n"))))
    # None = inherited from the layout; we must not restyle a title that fits
    assert out.slides[0].shapes.title.text_frame.paragraphs[0].font.size is None


def test_dense_bullet_list_shrinks():
    outline = "# Ideas\n" + "".join(f"- point {i}\n" for i in range(14))
    out = Presentation(io.BytesIO(build_pptx(parse_outline(outline))))
    (body,) = [p for p in out.slides[0].placeholders if p.placeholder_format.idx == 1]
    assert len(body.text_frame.paragraphs) == 14
    assert body.text_frame.paragraphs[0].font.size.pt <= 14


def test_find_layout_by_name_is_fuzzy():
    prs = Presentation()
    assert _find_layout_by_name(prs, "section header").name == "Section Header"
    assert _find_layout_by_name(prs, "TWO CONTENT").name == "Two Content"
    assert _find_layout_by_name(prs, "nope") is None


def test_list_layouts_reports_names():
    names = [n for n, _ in list_layouts(None)]
    assert "Title Slide" in names and "Title and Content" in names


class _FakeRequest:
    base_url = "http://host.example/"


def _run_action(async_api: bool, content: str = SAMPLE, __request__=None):
    """Drive Action.action against stand-in Open WebUI Files/Storage modules.

    Open WebUI's Files API is sync in older versions and async in current ones,
    so both shapes have to work. The async fakes only record their effect
    *inside* the coroutine — a dropped 'await' therefore shows up as a missing
    file record rather than passing silently.
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

    class Record:
        filename = "brand.pptx"
        path = template_path

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
                return Record()

            return _get() if async_api else Record()

    class Storage:
        @staticmethod
        def upload_file(stream, filename, tags=None):
            return stream.read(), f"/data/{filename}"

        @staticmethod
        def get_file(path):
            return path

    files_mod = types.ModuleType("open_webui.models.files")
    files_mod.Files = Files
    files_mod.FileForm = lambda **kw: kw
    storage_mod = types.ModuleType("open_webui.storage.provider")
    storage_mod.Storage = Storage
    fakes = {
        "open_webui": types.ModuleType("open_webui"),
        "open_webui.models": types.ModuleType("open_webui.models"),
        "open_webui.models.files": files_mod,
        "open_webui.storage": types.ModuleType("open_webui.storage"),
        "open_webui.storage.provider": storage_mod,
    }

    events = []

    async def emitter(event):
        events.append(event)

    body = {
        "id": "m1",
        "messages": [
            {"id": "m0", "role": "user", "content": "deck please", "files": [{"id": "f1"}]},
            {"id": "m1", "role": "assistant", "content": content},
        ],
    }
    sys.modules.update(fakes)
    try:
        asyncio.run(
            Action().action(
                body,
                __user__={"id": "u1"},
                __event_emitter__=emitter,
                __request__=__request__,
            )
        )
    finally:
        for key in fakes:
            sys.modules.pop(key, None)
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
    next click exported the error report as a deck."""
    events, inserted = _run_action(False, content="Sorry, I can't help with that.")
    assert inserted == {}  # nothing was built
    (message,) = [e for e in events if e["type"] == "message"]
    posted = message["data"]["content"]
    assert APPENDIX_MARKER in posted
    assert parse_outline(posted) == []
    # and once appended to the message, it is invisible to the next click
    appended = {"id": "m1", "role": "assistant", "content": "Sorry, I can't help with that." + posted}
    assert _pick_outline([appended], "m1") == "Sorry, I can't help with that."


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


if __name__ == "__main__":
    for _name, _fn in sorted(list(globals().items())):
        if _name.startswith("test_"):
            _fn()
    print("ok")
