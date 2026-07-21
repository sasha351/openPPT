"""
title: Export to PowerPoint
author: openPPT
version: 0.3.6
requirements: python-pptx
description: Adds an "Export to PowerPoint" button that converts a Markdown outline in the assistant message into a downloadable .pptx, with speaker notes, images, and custom templates. Attach a .pptx/.potx to the chat and the deck inherits its theme, fonts, and layouts — the model just dumps content as an outline and picks layouts by name. Works with any model (no tool calling needed).
"""

import inspect
import io
import re
import urllib.request
import uuid

from pptx import Presentation
from pptx.enum.shapes import PP_PLACEHOLDER
from pptx.enum.text import MSO_AUTO_SIZE
from pptx.util import Pt

PPTX_MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation"

IMAGE_RE = re.compile(r"^!\[(.*?)\]\((.*?)\)$")
NOTES_RE = re.compile(r"^notes?:\s*(.*)$", re.IGNORECASE)
# '# Title @layout: Section Header' or '# Title {layout: Section Header}'
LAYOUT_RE = re.compile(r"\s*(?:@layout:|\{layout:)\s*([^}]+?)\}?\s*$", re.IGNORECASE)
LAYOUT_LINE_RE = re.compile(r"^layout:\s*(.*)$", re.IGNORECASE)

# Tolerant slide-start grammar: real models rarely stick to plain '# '.
HEADING_RE = re.compile(r"^(#{1,6})\s+(.+)")
SLIDE_PREFIX_RE = re.compile(r"^\**slide\s*\d+\s*[:.\-]\s*(.+?)\**$", re.IGNORECASE)
BOLD_ONLY_RE = re.compile(r"^\*\*(.+?)\*\*:?$")
SEPARATOR_RE = re.compile(r"^(-{3,}|\*{3,}|_{3,})$")
QUOTE_RE = re.compile(r"^>\s*(.*)")
BULLET_RE = re.compile(r"^([-*+]|\d+[.)])\s+(.+)")
LINK_RE = re.compile(r"\[([^\]]+)\]\([^)]*\)")

# Precedes anything openPPT appends to a chat message — the download link, the
# no-outline diagnostic. An HTML comment renders invisibly in chat but survives
# a round trip back in, and parse_outline stops at this line, so re-clicking
# Export can't parse openPPT's own output back into slides (v0.3.5 exported its
# own error report as a deck).
APPENDIX_MARKER = "<!-- openppt -->"

# Kept equal to the docstring's 'version:' line — a stale paste in Open WebUI's
# Functions list is otherwise invisible, and the diagnostic is where it shows.
VERSION = "0.4.0"


def _strip_md(text: str) -> str:
    """Flatten inline Markdown so bullets read as plain text on a slide."""
    return re.sub(r"[*`]+", "", LINK_RE.sub(r"\1", text)).strip()


def _new_slide(title: str, layout: str = "") -> dict:
    return {"title": title, "subtitle": "", "layout": layout, "notes": "", "bullets": []}


def _split_layout(title: str):
    """Pull a trailing '@layout: X' / '{layout: X}' off a heading title."""
    m = LAYOUT_RE.search(title)
    if m:
        return title[: m.start()].strip(), m.group(1).strip()
    return title, ""


def _slide_level(lines) -> int:
    """Heading level that marks slides: whichever is most frequent.

    A deck written with '## ' slides and one '# ' deck title should not treat
    every '## ' as a bullet, which is what keying on '# ' alone does.
    """
    counts = {}
    for raw in lines:
        m = HEADING_RE.match(raw.strip())
        if m:
            counts[len(m.group(1))] = counts.get(len(m.group(1)), 0) + 1
    if not counts:
        return 0
    top = max(counts.values())
    return min(lvl for lvl, n in counts.items() if n == top)

# Placeholder type groups (see PP_PLACEHOLDER).
TITLE_TYPES = (PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.CENTER_TITLE)
SUBTITLE_TYPES = (PP_PLACEHOLDER.SUBTITLE,)
BODY_TYPES = (PP_PLACEHOLDER.BODY, PP_PLACEHOLDER.OBJECT)
# Placeholders that are chrome, never content we want to fill.
CHROME_TYPES = (
    PP_PLACEHOLDER.DATE,
    PP_PLACEHOLDER.FOOTER,
    PP_PLACEHOLDER.SLIDE_NUMBER,
    PP_PLACEHOLDER.HEADER,
)


def parse_outline(markdown: str) -> list:
    """Parse a Markdown outline into slides.

    Each slide is {"title", "subtitle", "layout", "notes",
    "bullets": [(level, text) | (level, {"image": url, "alt": alt})]}.

    '# Heading' starts a slide, '-'/'*'/'1.' lines are bullets (indent = nesting),
    '## ' on a bullet-less first slide is its subtitle. A bullet written as
    '![alt](url)' embeds an image instead of text. A line starting with
    'Notes:' (anywhere after the slide heading) is appended to that slide's
    speaker notes.

    A slide can name a template layout so the model chooses the design per
    slide: either inline on the heading — '# Title @layout: Section Header'
    (or '{layout: Section Header}') — or on its own 'Layout: <name>' line.
    The name is matched against the template's layouts in build_pptx.

    The grammar is tolerant, because local models rarely emit plain '# '
    exactly. Slides start at whichever heading level is most frequent, so a
    '## '-per-slide deck works and a lone '# ' above it becomes the deck
    title. With no headings at all, 'Slide 3: Title' prefixes, lone
    '**Bold Title**' lines and '---' separators start slides instead.
    Bullets may use '-', '*', '+', '1.' or '1)', and '> ' quoted lines are
    speaker notes alongside 'Notes:'. Inline '**bold**', '`code`' and
    '[links](url)' are flattened to plain text.

    Text before the first slide marker (chat prose) is ignored. Returns []
    if no slides found. Everything from the `APPENDIX_MARKER` line onward is
    ignored too, so re-exporting a message openPPT already appended to (a
    download link, a diagnostic) doesn't parse that text back into slides.
    """
    lines = markdown.splitlines()
    level_marker = _slide_level(lines)

    slides = []
    slide = None
    expect_title = False

    def start(title, layout=""):
        nonlocal slide
        slide = _new_slide(title, layout)
        slides.append(slide)

    for raw in lines:
        line = raw.strip()
        if line == APPENDIX_MARKER:
            break  # everything below is openPPT's own appended output
        if not line or line.startswith("```"):
            continue

        heading = HEADING_RE.match(line)
        if heading and level_marker:
            lvl = len(heading.group(1))
            title, layout = _split_layout(heading.group(2).strip())
            if lvl <= level_marker:
                start(_strip_md(title), layout)
                continue
            if lvl == level_marker + 1 and len(slides) == 1 and not slide["bullets"]:
                slide["subtitle"] = _strip_md(title)
                continue
            if slide is not None:
                slide["bullets"].append((0, _strip_md(title)))
            continue

        if not level_marker:
            if SEPARATOR_RE.match(line):
                expect_title = True
                continue
            prefixed = SLIDE_PREFIX_RE.match(line)
            bolded = BOLD_ONLY_RE.match(line)
            if prefixed or bolded:
                title, layout = _split_layout((prefixed or bolded).group(1).strip())
                start(_strip_md(title), layout)
                expect_title = False
                continue
            if expect_title and not BULLET_RE.match(line):
                title, layout = _split_layout(line)
                start(_strip_md(title), layout)
                expect_title = False
                continue

        if slide is None:
            continue

        notes_match = NOTES_RE.match(line) or QUOTE_RE.match(line)
        if notes_match:
            slide["notes"] = (slide["notes"] + "\n" + notes_match.group(1)).strip()
            continue
        layout_line = LAYOUT_LINE_RE.match(line)
        if layout_line:
            slide["layout"] = layout_line.group(1).strip()
            continue

        indent = len(raw) - len(raw.lstrip())
        level = min(indent // 2, 4)
        bullet = BULLET_RE.match(line)
        text = bullet.group(2).strip() if bullet else line
        image_match = IMAGE_RE.match(text)
        if image_match:
            slide["bullets"].append((level, {"alt": image_match.group(1), "image": image_match.group(2)}))
        else:
            slide["bullets"].append((level, _strip_md(text)))
    return slides


async def _maybe_await(value):
    """Open WebUI's Files API turned async; a pasted Function must fit both.

    Calling an async Files method without awaiting returns a coroutine that
    never runs — the insert silently doesn't happen and the download 404s.
    """
    return await value if inspect.isawaitable(value) else value


def _fetch_image(url: str):
    """Return image bytes for a local path or http(s) URL, or None on failure."""
    try:
        if url.startswith("http://") or url.startswith("https://"):
            with urllib.request.urlopen(url, timeout=10) as resp:
                return io.BytesIO(resp.read())
        with open(url, "rb") as f:
            return io.BytesIO(f.read())
    except Exception:
        return None


def _ph_type(placeholder):
    """Placeholder type, tolerant of odd templates that leave it unset."""
    try:
        return placeholder.placeholder_format.type
    except Exception:
        return None


def _clear_slides(prs) -> None:
    """Remove any example slides from a template so the deck starts clean.

    A .potx has none, but a .pptx used as a template often carries sample
    slides; the template's masters, layouts, and theme are untouched.
    """
    sld_id_lst = prs.slides._sldIdLst
    for sld_id in list(sld_id_lst):
        rid = sld_id.get(
            "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
        )
        if rid:
            prs.part.drop_rel(rid)
        sld_id_lst.remove(sld_id)


def _layout_stats(layout):
    """Summarise a layout by placeholder role for role-based selection."""
    titles = subtitles = bodies = 0
    for ph in layout.placeholders:
        t = _ph_type(ph)
        if t in TITLE_TYPES:
            titles += 1
        elif t in SUBTITLE_TYPES:
            subtitles += 1
        elif t in BODY_TYPES:
            bodies += 1
    return {"title": titles, "subtitle": subtitles, "body": bodies}


def list_layouts(prs_or_bytes) -> list:
    """Return [(name, {'title','subtitle','body' counts})] for a template.

    Accepts a Presentation, template bytes, or a file-like/stream. Handy for
    telling the model which layout names it can pick from.
    """
    prs = _as_presentation(prs_or_bytes)
    return [(lay.name, _layout_stats(lay)) for lay in prs.slide_layouts]


def _as_presentation(template):
    """Coerce None / bytes / stream / Presentation into a Presentation."""
    if template is None:
        return Presentation()
    if hasattr(template, "slide_layouts"):
        return template
    if isinstance(template, (bytes, bytearray)):
        return Presentation(io.BytesIO(template))
    return Presentation(template)  # file path or file-like


def _find_layout_by_name(prs, name: str):
    """Case-insensitive match of a layout by name (exact, then substring)."""
    if not name:
        return None
    target = name.strip().lower()
    layouts = list(prs.slide_layouts)
    for lay in layouts:
        if lay.name.lower() == target:
            return lay
    for lay in layouts:
        if target in lay.name.lower() or lay.name.lower() in target:
            return lay
    return None


def _pick_layout(prs, role: str):
    """Choose a layout for a role ('title', 'content', 'title_only') by
    inspecting placeholders, so it works on any template regardless of the
    order or names its layouts happen to use."""
    layouts = list(prs.slide_layouts)
    scored = [(lay, _layout_stats(lay)) for lay in layouts]

    if role == "title":
        # Prefer title + subtitle, no bullet body (a real title slide).
        for lay, s in scored:
            if s["title"] and s["subtitle"] and not s["body"]:
                return lay
        for lay, s in scored:
            if s["title"] and s["subtitle"]:
                return lay
    if role == "title_only":
        # A title, nothing else to fill.
        for lay, s in scored:
            if s["title"] and not s["subtitle"] and not s["body"]:
                return lay
    # role == "content" (and fallbacks for the above): title + exactly one body.
    for lay, s in scored:
        if s["title"] and s["body"] == 1 and not s["subtitle"]:
            return lay
    for lay, s in scored:
        if s["title"] and s["body"]:
            return lay
    for lay, s in scored:
        if s["title"]:
            return lay
    return layouts[0]


def _title_placeholder(slide):
    if slide.shapes.title is not None:
        return slide.shapes.title
    for ph in slide.placeholders:
        if _ph_type(ph) in TITLE_TYPES:
            return ph
    return None


def _subtitle_placeholder(slide):
    for ph in slide.placeholders:
        if _ph_type(ph) in SUBTITLE_TYPES:
            return ph
    return None


def _body_placeholder(slide):
    """First fillable body placeholder that isn't the title/subtitle/chrome."""
    title = _title_placeholder(slide)
    title_id = id(title) if title is not None else None
    for ph in slide.placeholders:
        if id(ph) == title_id:
            continue
        t = _ph_type(ph)
        if t in BODY_TYPES:
            return ph
    # Fall back to any non-title, non-chrome, non-subtitle placeholder.
    for ph in slide.placeholders:
        if id(ph) == title_id:
            continue
        t = _ph_type(ph)
        if t in CHROME_TYPES or t in SUBTITLE_TYPES:
            continue
        return ph
    return None


def _title_size(text: str):
    """Point size for a title long enough to overflow, or None to keep the
    template's. python-pptx never recomputes PowerPoint's autofit, so without
    this a 200-character title renders at full size across the whole slide."""
    for limit, size in ((50, None), (80, 32), (120, 26), (200, 20)):
        if len(text) <= limit:
            return size
    return 16


def _body_size(bullets) -> int:
    """Point size for a bullet list, shrinking as it gets denser. Long bullets
    wrap, so count the lines they'll take rather than the bullets."""
    lines = sum(1 + len(text) // 60 for _, text in bullets)
    for limit, size in ((5, 20), (8, 16), (12, 14)):
        if lines <= limit:
            return size
    return 12


def _fit(text_frame) -> None:
    """Wrap, and let PowerPoint shrink further than our estimate if it must."""
    text_frame.word_wrap = True
    text_frame.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE


def _set_title(slide, text: str) -> None:
    ph = _title_placeholder(slide)
    if ph is None or not ph.has_text_frame:
        return
    ph.text_frame.text = text
    _fit(ph.text_frame)
    size = _title_size(text)
    if size:
        for p in ph.text_frame.paragraphs:
            p.font.size = Pt(size)


def _fill_bullets(text_frame, bullets) -> None:
    _fit(text_frame)
    size = Pt(_body_size(bullets))
    for j, (level, text) in enumerate(bullets):
        p = text_frame.paragraphs[0] if j == 0 else text_frame.add_paragraph()
        p.text = text
        p.level = level
        p.font.size = size


def build_pptx(slides: list, template=None) -> bytes:
    """Build a .pptx from parsed slides.

    template: optional .pptx/.potx as bytes, a stream, a path, or a
    Presentation. When given, the deck inherits that template's theme,
    fonts, and layouts, and each slide's 'layout' name is matched against the
    template's layouts. When omitted, the python-pptx default template is used.
    """
    prs = _as_presentation(template)
    _clear_slides(prs)

    for i, s in enumerate(slides):
        text_bullets = [(lvl, t) for lvl, t in s["bullets"] if isinstance(t, str)]
        image_bullets = [(lvl, t) for lvl, t in s["bullets"] if isinstance(t, dict)]
        image_only = not text_bullets and image_bullets
        is_title = i == 0 and not s["bullets"]

        layout = _find_layout_by_name(prs, s.get("layout", ""))
        if layout is None:
            role = "title" if is_title else ("title_only" if image_only else "content")
            layout = _pick_layout(prs, role)

        ps = prs.slides.add_slide(layout)
        _set_title(ps, s["title"])

        if s["subtitle"]:
            sub = _subtitle_placeholder(ps) or _body_placeholder(ps)
            if sub is not None:
                sub.text = s["subtitle"]

        if text_bullets:
            body = _body_placeholder(ps)
            if body is not None and body.has_text_frame:
                _fill_bullets(body.text_frame, text_bullets)

        for _, img in image_bullets:
            data = _fetch_image(img["image"])
            if data is not None:
                ps.shapes.add_picture(data, Pt(60), Pt(120), height=Pt(340))

        if s["notes"]:
            ps.notes_slide.notes_text_frame.text = s["notes"]

    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()


def _pick_outline(messages: list, target=None) -> str:
    """Content to export: the clicked assistant message, else the newest one
    that actually holds an outline.

    Open WebUI sometimes hands the action a message whose content is empty, and
    a strict clicked-message-only read then reports "no outline" at a chat that
    plainly has one. Returns the clicked/newest text when nothing parses, so
    the caller can echo what it actually read.
    """
    candidates = []
    for msg in reversed(messages or []):
        if msg.get("role") != "assistant":
            continue
        # Drop our own appended output (download link, diagnostic) so a
        # re-click echoes back what the model wrote, not what we last posted.
        # The blank line right before the marker is ours too (see the emits
        # below), so trim it along with everything after — but only when the
        # marker is actually present, or an ordinary message's own trailing
        # whitespace would get eaten.
        text = msg.get("content") or ""
        before, marker, _ = text.partition(APPENDIX_MARKER)
        if marker:
            text = before.rstrip()
        if target is not None and msg.get("id") == target:
            candidates.insert(0, text)
        else:
            candidates.append(text)
    for text in candidates:
        if parse_outline(text):
            return text
    return candidates[0] if candidates else ""


class Action:
    async def action(
        self,
        body: dict,
        __user__: dict = None,
        __event_emitter__=None,
        __request__=None,
    ):
        async def emit(event):
            if __event_emitter__:
                await __event_emitter__(event)

        async def status(desc, done=False):
            await emit({"type": "status", "data": {"description": desc, "done": done}})

        async def notify(content, kind="info"):
            # A model preset can hide the status line via its status_updates
            # capability, so every outcome also toasts and hits the server log.
            print(f"[openPPT] {kind}: {content}", flush=True)
            await emit({"type": "notification", "data": {"type": kind, "content": content}})

        try:
            messages = body.get("messages", [])
            target = body.get("id")
            content = _pick_outline(messages, target)
            slides = parse_outline(content)
            if not slides:
                help_text = (
                    "openPPT: nothing slide-shaped in that message. Ask the model for "
                    "an outline — a '# Slide Title' heading per slide, with '-' "
                    "bullets under it."
                )
                await notify(help_text, "warning")
                await status(help_text, done=True)
                # Toasts and the status line get truncated, so the detail goes in
                # the chat. Written so parse_outline finds nothing in it: no '---'
                # rule, no bullet list, no lone '**bold**' line — otherwise the
                # next click exports this diagnostic as a deck (v0.3.5 did).
                seen = content[:400] if content else "(nothing — the message was empty)"
                await emit(
                    {
                        "type": "message",
                        "data": {
                            "content": (
                                f"\n\n{APPENDIX_MARKER}\n"
                                f"**openPPT {VERSION} diagnostic** — {help_text} "
                                f"(messages in request: {len(messages)}, clicked id: "
                                f"`{target}`, read {len(content)} characters)\n\n"
                                f"What openPPT read from that message:\n\n"
                                f"```\n{seen}\n```\n"
                            )
                        },
                    }
                )
                return

            template = await self._find_template(body, __user__)
            wants_layout = any(s.get("layout") for s in slides)
            if wants_layout and template is None:
                await status(
                    "This deck names layouts but no .pptx/.potx template is "
                    "attached — using the default theme.",
                )

            note = " on your template" if template is not None else ""
            await status(f"Building PowerPoint ({len(slides)} slides{note})…")
            data = build_pptx(slides, template=template)

            # ponytail: internal Files/Storage API — swap to REST POST /api/v1/files/
            # with the user's token if these internals change across Open WebUI versions.
            from open_webui.models.files import FileForm, Files
            from open_webui.storage.provider import Storage

            slug = "".join(
                c if c.isalnum() else "-" for c in slides[0]["title"].lower()
            ).strip("-") or "deck"
            name = f"{slug}.pptx"
            file_id = str(uuid.uuid4())
            try:
                _, file_path = Storage.upload_file(
                    io.BytesIO(data), f"{file_id}_{name}", tags={}
                )
            except TypeError:  # older Open WebUI: no tags param
                _, file_path = Storage.upload_file(io.BytesIO(data), f"{file_id}_{name}")
            record = await _maybe_await(
                Files.insert_new_file(
                    __user__["id"],
                    FileForm(
                        id=file_id,
                        filename=name,
                        path=file_path,
                        data={},
                        meta={"name": name, "content_type": PPTX_MIME, "size": len(data)},
                    ),
                )
            )
            # insert_new_file returns None on failure instead of raising — without
            # this the link is posted and 404s ("download ready", nothing downloads).
            if record is None:
                raise RuntimeError("Open WebUI rejected the file record")

            # Two delivery paths on purpose. The 'files' event is how Open WebUI
            # returns its own generated files: it lands in the message's files
            # column and renders as an attachment chip (with a .pptx preview and
            # a download). The markdown link covers versions whose event handler
            # ignores 'files'. Emitted separately so one failing can't lose both.
            url = f"/api/v1/files/{file_id}/content"
            full_url = url
            if __request__ is not None:
                try:
                    full_url = str(__request__.base_url).rstrip("/") + url
                except Exception:
                    pass
            try:
                await emit(
                    {
                        "type": "files",
                        "data": {
                            "files": [
                                {
                                    "type": "file",
                                    "id": file_id,
                                    "url": file_id,  # FileItem prefixes /api/v1/files/
                                    "name": name,
                                    "collection_name": "",
                                    "status": "uploaded",
                                    "size": len(data),
                                    "content_type": PPTX_MIME,
                                    "meta": {
                                        "name": name,
                                        "content_type": PPTX_MIME,
                                        "size": len(data),
                                    },
                                }
                            ]
                        },
                    }
                )
            except Exception as e:
                print(f"[openPPT] files event rejected: {type(e).__name__}: {e}", flush=True)

            # The marker line lets a future export ignore everything below it
            # (see parse_outline) — otherwise re-clicking Export on this same
            # message would parse this link into a bogus trailing bullet. The
            # full URL repeated on its own line is copy-pasteable even where
            # neither the chip nor the markdown link renders.
            await emit(
                {
                    "type": "message",
                    "data": {
                        "content": (
                            f"\n\n{APPENDIX_MARKER}\n📊 [Download {name}]({full_url})"
                            f"\n\n{full_url}"
                        )
                    },
                }
            )
            await notify(f"{name} ready — {full_url}", "success")
            await status(f"{name} ready — {full_url}", done=True)
        except Exception as e:
            msg = f"openPPT export failed: {type(e).__name__}: {e}"
            await notify(msg, "error")
            await status(msg, done=True)

    async def _find_template(self, body: dict, __user__: dict):
        """Return template bytes from the most recent .pptx/.potx attached to
        the chat, or None. Files can be attached to any message or ride along
        in body['files']; we scan newest-first and load the first template."""
        try:
            from open_webui.models.files import Files
            from open_webui.storage.provider import Storage
        except Exception:
            return None

        def file_ids():
            for msg in reversed(body.get("messages", [])):
                for f in msg.get("files", []) or []:
                    fid = (f.get("file") or {}).get("id") or f.get("id")
                    if fid:
                        yield fid, (f.get("name") or (f.get("file") or {}).get("filename") or "")
            for f in body.get("files", []) or []:
                fid = (f.get("file") or {}).get("id") or f.get("id")
                if fid:
                    yield fid, (f.get("name") or (f.get("file") or {}).get("filename") or "")

        for fid, hint in file_ids():
            try:
                rec = await _maybe_await(Files.get_file_by_id(fid))
                if rec is None:
                    continue
                fname = (rec.filename or hint or "").lower()
                if not fname.endswith((".pptx", ".potx")):
                    continue
                with open(Storage.get_file(rec.path), "rb") as fh:
                    return fh.read()
            except Exception:
                continue
        return None
