"""
title: Export to PowerPoint
author: openPPT
version: 0.3.0
requirements: python-pptx
description: Adds an "Export to PowerPoint" button that converts a Markdown outline in the assistant message into a downloadable .pptx, with speaker notes, images, and custom templates. Attach a .pptx/.potx to the chat and the deck inherits its theme, fonts, and layouts — the model just dumps content as an outline and picks layouts by name. Works with any model (no tool calling needed).
"""

import io
import re
import urllib.request
import uuid

from pptx import Presentation
from pptx.enum.shapes import PP_PLACEHOLDER
from pptx.util import Pt

IMAGE_RE = re.compile(r"^!\[(.*?)\]\((.*?)\)$")
NOTES_RE = re.compile(r"^notes?:\s*(.*)$", re.IGNORECASE)
# '# Title @layout: Section Header' or '# Title {layout: Section Header}'
LAYOUT_RE = re.compile(r"\s*(?:@layout:|\{layout:)\s*([^}]+?)\}?\s*$", re.IGNORECASE)
LAYOUT_LINE_RE = re.compile(r"^layout:\s*(.*)$", re.IGNORECASE)

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

    Text before the first '#' (chat prose) is ignored. Returns [] if no
    slides found.
    """
    slides = []
    slide = None
    for raw in markdown.splitlines():
        line = raw.strip()
        if not line or line.startswith("```"):
            continue
        if line.startswith("# "):
            title = line[2:].strip()
            layout = ""
            layout_match = LAYOUT_RE.search(title)
            if layout_match:
                layout = layout_match.group(1).strip()
                title = title[: layout_match.start()].strip()
            slide = {
                "title": title,
                "subtitle": "",
                "layout": layout,
                "notes": "",
                "bullets": [],
            }
            slides.append(slide)
            continue
        if slide is None:
            continue
        notes_match = NOTES_RE.match(line)
        if notes_match:
            slide["notes"] = (slide["notes"] + "\n" + notes_match.group(1)).strip()
            continue
        layout_line = LAYOUT_LINE_RE.match(line)
        if layout_line:
            slide["layout"] = layout_line.group(1).strip()
            continue
        indent = len(raw) - len(raw.lstrip())
        level = min(indent // 2, 4)
        if line.startswith("## ") and len(slides) == 1 and not slide["bullets"]:
            slide["subtitle"] = line[3:].strip()
            continue
        if line[:2] in ("- ", "* "):
            text = line[2:].strip()
        elif line.split(".", 1)[0].isdigit() and "." in line:
            text = line.split(".", 1)[1].strip()
        else:
            text = line
        image_match = IMAGE_RE.match(text)
        if image_match:
            slide["bullets"].append((level, {"alt": image_match.group(1), "image": image_match.group(2)}))
        else:
            slide["bullets"].append((level, text))
    return slides


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


def _set_title(slide, text: str) -> None:
    ph = _title_placeholder(slide)
    if ph is not None:
        ph.text = text


def _fill_bullets(text_frame, bullets) -> None:
    for j, (level, text) in enumerate(bullets):
        p = text_frame.paragraphs[0] if j == 0 else text_frame.add_paragraph()
        p.text = text
        p.level = level
        p.font.size = Pt(20)


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


class Action:
    async def action(
        self,
        body: dict,
        __user__: dict = None,
        __event_emitter__=None,
        __request__=None,
    ):
        async def status(desc, done=False):
            if __event_emitter__:
                await __event_emitter__(
                    {"type": "status", "data": {"description": desc, "done": done}}
                )

        try:
            content = ""
            for msg in reversed(body.get("messages", [])):
                if msg.get("role") == "assistant":
                    content = msg.get("content", "")
                    break

            slides = parse_outline(content)
            if not slides:
                await status(
                    "No slides found — ask the model to write the deck as "
                    "'# Slide Title' headings with '-' bullets.",
                    done=True,
                )
                return

            template = self._find_template(body, __user__)
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
            Files.insert_new_file(
                __user__["id"],
                FileForm(
                    id=file_id,
                    filename=name,
                    path=file_path,
                    data={},
                    meta={
                        "name": name,
                        "content_type": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
                        "size": len(data),
                    },
                ),
            )

            if __event_emitter__:
                await __event_emitter__(
                    {
                        "type": "message",
                        "data": {
                            "content": f"\n\n📊 [Download {name}](/api/v1/files/{file_id}/content)"
                        },
                    }
                )
            await status(f"{name} ready", done=True)
        except Exception as e:
            await status(f"PowerPoint export failed: {e}", done=True)

    def _find_template(self, body: dict, __user__: dict):
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
                rec = Files.get_file_by_id(fid)
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
