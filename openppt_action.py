"""
title: Export to PowerPoint
author: openPPT
version: 0.7.3
requirements: python-pptx
description: Adds an "Export to PowerPoint" button that converts an HTML outline in the assistant message into a downloadable .pptx, with tables, code blocks, speaker notes, images, and custom templates. Attach a .pptx/.potx to the chat and the deck inherits its theme, fonts, and layouts — the model just dumps content as an outline and picks layouts by name. Works with any model (no tool calling needed).
"""

import asyncio
import base64
import html
import inspect
import io
import ipaddress
import json
import math
import os
import re
import socket
import urllib.parse
import urllib.request
import uuid
from html.parser import HTMLParser

try:
    from pydantic import BaseModel, Field
except ImportError:  # outline_to_pptx.py run without Open WebUI's pydantic

    def Field(default=None, **_):
        return default

    class BaseModel:
        def __init__(self, **values):
            for name in getattr(type(self), "__annotations__", {}):
                setattr(self, name, values.get(name, getattr(type(self), name)))

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import PP_PLACEHOLDER
from pptx.enum.text import MSO_AUTO_SIZE
from pptx.oxml.ns import qn
from pptx.util import Emu, Pt

PPTX_MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation"

# Precedes anything openPPT appends to a chat message — the download link, the
# no-outline diagnostic. An HTML comment renders invisibly in chat but survives
# a round trip back in, and the parser stops dead at this comment (see
# _OutlineParser.handle_comment), so re-clicking Export can't parse openPPT's
# own output back into slides (v0.3.5 exported its own error report as a deck).
APPENDIX_MARKER = "<!-- openppt -->"

# Kept equal to the docstring's 'version:' line — a stale paste in Open WebUI's
# Functions list is otherwise invisible, and the diagnostic is where it shows.
VERSION = "0.7.3"

# How long the browser gets to start the auto-download before the export
# carries on without it (the chip and the link still deliver the file).
AUTO_DOWNLOAD_TIMEOUT = 15


def _new_slide(title: str, layout: str = "") -> dict:
    return {
        "title": title,
        "subtitle": "",
        "layout": layout,
        "notes": "",
        "bullets": [],
        "runs": [],
        "table": [],
        "code": "",
        "columns": [],
    }


def _attr(attrs: dict, name: str) -> str:
    return (attrs.get(name) or "").strip()


# Inline tags kept as formatting on the bullet's text runs.
INLINE_FORMATS = {"b": "b", "strong": "b", "i": "i", "em": "i", "u": "u", "code": "code"}
COLUMN_TAGS = ("column", "col")


def _normalize_runs(pieces) -> list:
    """[(text, formats)] → whitespace collapsed (HTML semantics: a line break
    inside an '<li>' is just a space), ends trimmed, same-format neighbours
    merged, empties dropped."""
    out = []
    for text, fmt in pieces:
        text = re.sub(r"\s+", " ", text)
        if out and out[-1][0].endswith(" ") and text.startswith(" "):
            text = text[1:]
        if not text:
            continue
        if out and out[-1][1] == fmt:
            out[-1] = (out[-1][0] + text, fmt)
        else:
            out.append((text, fmt))
    while out and not out[0][0].strip():
        out.pop(0)
    while out and not out[-1][0].strip():
        out.pop()
    if out:
        out[0] = (out[0][0].lstrip(), out[0][1])
        out[-1] = (out[-1][0].rstrip(), out[-1][1])
    return [(t, f) for t, f in out if t]


class _OutlineParser(HTMLParser):
    """Turns a stream of tag events into slide dicts for build_pptx.

    Deliberately tolerant rather than a strict-HTML validator, because local
    models don't emit strict HTML any more reliably than they emitted strict
    Markdown: unclosed '<slide>'/'<li>'/'<p>'/'<column>' tags are auto-closed,
    stray text directly inside '<slide>' (not wrapped in '<p>'/'<li>') still
    becomes bullets — one per line — instead of being dropped, and
    '<blockquote>' is accepted as a synonym for '<notes>'. HTMLParser itself
    is already lenient about malformed markup (mismatched tags, missing
    quotes, bad nesting) and never raises on it, which is most of why HTML
    beats hand-rolled regex here.
    """

    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.slides = []
        self.current = None
        self.in_slide = False
        self.stopped = False
        self._reset_slide_state()

    def _reset_slide_state(self):
        self.list_depth = 0
        self.li_stack = []  # [{"level", "buf": [(str, fmt)], "image": dict|None, "emitted": bool}]
        self.pre_depth = 0
        self.code_buf = []
        self.notes_depth = 0
        self.notes_buf = []
        self.p_depth = 0
        self.p_buf = []
        self.bare_buf = []
        self.table = None
        self.row = None
        self.cell_depth = 0
        self.cell_buf = []
        self.column = None
        self.fmt = {"b": 0, "i": 0, "u": 0, "code": 0}

    # -- emitting ---------------------------------------------------------
    def _target(self):
        """Where bullets go: the open '<column>', else the slide itself."""
        return self.column if self.column is not None else self.current

    def _emit(self, level, pieces):
        runs = _normalize_runs(pieces)
        text = "".join(t for t, _ in runs)
        if not text:
            return
        target = self._target()
        target["bullets"].append((level, text))
        # Formatting rides alongside, so (level, text) bullets stay plain.
        target["runs"].append(runs if any(f for _, f in runs) else None)

    def _emit_image(self, level, image):
        target = self._target()
        target["bullets"].append((level, image))
        target["runs"].append(None)

    def _emit_frame(self, frame):
        if frame["emitted"]:
            return
        if frame["image"] is not None:
            self._emit_image(frame["level"], frame["image"])
        self._emit(frame["level"], frame["buf"])
        frame["buf"], frame["emitted"] = [], True

    def _flush_bare(self):
        """Bare text directly inside '<slide>' — models forget the '<li>'
        wrapper constantly — becomes one bullet per line."""
        line = []
        for text, fmt in self.bare_buf:
            parts = text.split("\n")
            for k, part in enumerate(parts):
                if k:
                    self._emit(0, line)
                    line = []
                line.append((part, fmt))
        self._emit(0, line)
        self.bare_buf = []

    def _flush_blocks(self):
        """Close whatever list item / paragraph / bare text is still open —
        before a column starts or ends, and when the slide does."""
        self._flush_bare()
        while self.li_stack:
            frame = self.li_stack.pop()
            if frame["image"] is not None and not frame["emitted"]:
                self._emit_image(frame["level"], frame["image"])
            if not frame["emitted"] or frame["buf"]:
                self._emit(frame["level"], frame["buf"])
        self.list_depth = 0
        if self.p_depth > 0:
            self._emit(0, self.p_buf)
            self.p_depth, self.p_buf = 0, []

    def _close_slide(self):
        """Flush whatever was still open (unclosed li/p/column/notes/table/
        pre) and end the current slide — called on '</slide>', on a new
        '<slide>' opening without a previous close, and at end-of-document."""
        self._flush_blocks()
        if self.notes_depth > 0:
            text = "".join(self.notes_buf).strip()
            if text:
                self.current["notes"] = (self.current["notes"] + "\n" + text).strip()
        if self.table is not None:
            if self.row is not None:
                self.table.append(self.row)
            self.current["table"].extend(self.table)
        if self.pre_depth > 0:
            self.current["code"] += "".join(self.code_buf)
        self._reset_slide_state()
        self.in_slide = False
        self.current = None

    # -- HTMLParser callbacks ---------------------------------------------
    def handle_comment(self, data):
        if data.strip() == "openppt":
            self.stopped = True  # hard stop: never parse openPPT's own appendix

    def handle_starttag(self, tag, attrs):
        if self.stopped:
            return
        a = dict(attrs)
        if self.pre_depth > 0 and tag != "pre":
            if tag != "code":  # transparent — doesn't count as nesting depth
                self.code_buf.append(self.get_starttag_text())
            return
        if tag == "pre":
            self._flush_bare()
            self.pre_depth += 1
            return
        if tag == "slide":
            if self.in_slide:
                self._close_slide()  # tolerate a missing '</slide>'
            self.current = _new_slide(_attr(a, "title"), _attr(a, "layout"))
            self.current["subtitle"] = _attr(a, "subtitle")
            self.slides.append(self.current)
            self.in_slide = True
            return
        if not self.in_slide:
            return  # stray tag before the first '<slide>' — ignore
        if tag in INLINE_FORMATS:
            self.fmt[INLINE_FORMATS[tag]] += 1
            return
        if tag == "br":
            self.handle_data("\n")  # a line break: a space in a bullet, a new bare line
            return
        self._flush_bare()
        if tag in COLUMN_TAGS and self.table is None:  # '<col>' in a table is HTML's own
            self._flush_blocks()
            self.column = {"heading": _attr(a, "heading") or _attr(a, "title"), "bullets": [], "runs": []}
            self.current["columns"].append(self.column)
        elif tag in ("ul", "ol"):
            if self.li_stack:
                # A nested list inside a still-open '<li>' — flush that li's
                # own text now, in document order, rather than waiting for its
                # '</li>' (which comes after the nested list closes and would
                # put the parent's text after its children's).
                self._emit_frame(self.li_stack[-1])
            self.list_depth += 1
        elif tag == "li":
            level = min(self.list_depth - 1, 4) if self.list_depth > 0 else 0
            self.li_stack.append(
                {"level": max(level, 0), "buf": [], "image": None, "emitted": False}
            )
        elif tag == "img":
            image = {"alt": _attr(a, "alt"), "image": _attr(a, "src")}
            if self.li_stack:
                self.li_stack[-1]["image"] = image
            else:
                self._emit_image(0, image)
        elif tag in ("notes", "blockquote"):
            self.notes_depth += 1
        elif tag == "p":
            self.p_depth += 1
        elif tag == "table":
            self.table = []
        elif tag == "tr":
            self.row = []
        elif tag in ("td", "th"):
            self.cell_depth += 1

    def handle_endtag(self, tag):
        if self.stopped:
            return
        if self.pre_depth > 0 and tag != "pre":
            if tag != "code":
                self.code_buf.append(f"</{tag}>")
            return
        if tag == "pre":
            self.pre_depth = max(0, self.pre_depth - 1)
            if self.pre_depth == 0 and self.in_slide:
                self.current["code"] += "".join(self.code_buf)
                self.code_buf = []
            return
        if tag == "slide":
            if self.in_slide:
                self._close_slide()
            return
        if not self.in_slide:
            return
        if tag in INLINE_FORMATS:
            key = INLINE_FORMATS[tag]
            self.fmt[key] = max(0, self.fmt[key] - 1)
            return
        self._flush_bare()
        if tag in COLUMN_TAGS and self.table is None:
            self._flush_blocks()
            self.column = None
        elif tag in ("ul", "ol"):
            self.list_depth = max(0, self.list_depth - 1)
        elif tag == "li":
            if self.li_stack:
                frame = self.li_stack.pop()
                if frame["image"] is not None and not frame["emitted"]:
                    self._emit_image(frame["level"], frame["image"])
                self._emit(frame["level"], frame["buf"])
        elif tag in ("notes", "blockquote"):
            self.notes_depth = max(0, self.notes_depth - 1)
            if self.notes_depth == 0:
                text = "".join(self.notes_buf).strip()
                self.notes_buf = []
                if text:
                    self.current["notes"] = (self.current["notes"] + "\n" + text).strip()
        elif tag == "p":
            self.p_depth = max(0, self.p_depth - 1)
            if self.p_depth == 0:
                self._emit(0, self.p_buf)
                self.p_buf = []
        elif tag == "table":
            if self.table is not None:
                self.current["table"].extend(self.table)
                self.table = None
        elif tag == "tr":
            if self.row is not None and self.table is not None:
                self.table.append(self.row)
            self.row = None
        elif tag in ("td", "th"):
            if self.cell_depth > 0:
                self.cell_depth -= 1
                text = re.sub(r"\s+", " ", "".join(self.cell_buf)).strip()
                self.cell_buf = []
                if self.row is not None:
                    self.row.append(text)

    def handle_data(self, data):
        if self.stopped or not self.in_slide:
            return
        fmt = tuple(sorted(k for k, v in self.fmt.items() if v > 0))
        if self.pre_depth > 0:
            self.code_buf.append(data)
        elif self.cell_depth > 0:
            self.cell_buf.append(data)
        elif self.li_stack:
            self.li_stack[-1]["buf"].append((data, fmt))
        elif self.p_depth > 0:
            self.p_buf.append((data, fmt))
        elif self.notes_depth > 0:
            self.notes_buf.append(data)
        else:
            self.bare_buf.append((data, fmt))


def _extract_fences(text: str) -> str:
    """Handle ``` fences the model adds despite being asked for plain HTML.

    A fence opened before the first '<slide' is a wrapper around the whole
    answer (models do this constantly) — unwrapped and parsed through. A
    fence opened after is a real code sample; its content is HTML-escaped and
    spliced in as a '<pre>' block so the outline parser needs only one code
    path ('<pre>', wherever the text came from), and code containing '<'/'>'
    can't be mistaken for markup.
    """
    seen_slide = False
    in_fence = False
    fence_is_code = False
    code_lines = []
    out = []
    for raw in text.splitlines():
        if raw.strip().startswith("```"):
            if not in_fence:
                in_fence, fence_is_code, code_lines = True, seen_slide, []
            else:
                in_fence = False
                if fence_is_code:
                    out.append(f"<pre>{html.escape(chr(10).join(code_lines))}</pre>")
                else:
                    out.extend(code_lines)
            continue
        if in_fence:
            code_lines.append(raw)
            continue
        if "<slide" in raw.lower():
            seen_slide = True
        out.append(raw)
    if in_fence:  # unclosed fence at EOF — flush what we have instead of losing it
        if fence_is_code:
            out.append(f"<pre>{html.escape(chr(10).join(code_lines))}</pre>")
        else:
            out.extend(code_lines)
    return "\n".join(out)

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


# Reasoning that a model emits inline, in every wrapper Open WebUI itself
# recognises (its DEFAULT_REASONING_TAGS), plus the <details type="..."> blocks
# it serialises reasoning and tool calls into in a message's legacy 'content'.
_REASONING_BLOCK_RE = re.compile(
    r"<(think|thinking|reason|reasoning|thought)\b[^>]*>.*?</\1\s*>"
    r"|<details\s[^>]*\btype=[\"'][^\"']*[\"'][^>]*>.*?</details\s*>"
    r"|<\|begin_of_thought\|>.*?<\|end_of_thought\|>"
    r"|◁think▷.*?◁/think▷",
    re.S | re.I,
)
# A closing tag with no opener: chat templates (Nemotron's, DeepSeek-R1's)
# pre-fill '<think>' in the prompt, so the reply starts mid-reasoning.
_ORPHAN_CLOSE_RE = re.compile(r"</(think|thinking|reasoning)\s*>", re.I)


def _strip_reasoning(text: str) -> str:
    """Remove a reasoning model's thinking before parsing. Drafts of the deck
    written while thinking would otherwise be exported as slides — and
    whichever comes first, draft or final, would win slide one."""
    text = _REASONING_BLOCK_RE.sub("", text or "")
    last = None
    for last in _ORPHAN_CLOSE_RE.finditer(text):
        pass
    if last is not None:
        text = text[last.end():]
    return text.replace("<|begin_of_solution|>", "").replace("<|end_of_solution|>", "")


def parse_outline(markup: str) -> list:
    """Parse an HTML outline into slides.

    Each slide is {"title", "subtitle", "layout", "notes", "table", "code",
    "bullets": [(level, text) | (level, {"image": url, "alt": alt})],
    "runs": [None | [(text, formats)]] aligned with bullets,
    "columns": [{"heading", "bullets", "runs"}]}.

    Each '<slide title="..." subtitle="..." layout="...">' element is one
    slide; 'subtitle' and 'layout' are optional, and 'layout' names a
    template layout matched against the template's layouts in build_pptx.
    Inside a slide: '<ul>'/'<ol>' of '<li>' are bullets, with a nested
    '<ul>'/'<ol>' inside an '<li>' nesting the bullet under it; '<p>' is a
    single bullet; '<notes>' (or '<blockquote>') is appended to the slide's
    speaker notes; '<img src="..." alt="...">' — bare or inside an '<li>' —
    embeds an image instead of text; '<table>' of '<tr>' of '<td>'/'<th>'
    becomes a real table; and '<pre>' (with or without a nested '<code>')
    becomes a code block rendered in a monospace box. A ``` fence is also
    accepted as a code block for models that fall back to it out of habit —
    escaped and folded into an equivalent '<pre>' before parsing.
    '<column heading="...">' (or '<col>') splits a slide into side-by-side
    columns holding their own bullets/images, and '<b>'/'<strong>',
    '<i>'/'<em>', '<u>', '<code>' are kept as run formatting ('runs').
    A reasoning model's inline thinking ('<think>…</think>' and the other
    wrappers Open WebUI recognises, or everything up to a lone '</think>')
    is dropped first, so slides it drafted while thinking aren't exported.

    The parser (html.parser.HTMLParser, stdlib) is deliberately tolerant
    rather than a strict validator: an unclosed '<slide>'/'<li>'/'<p>' is
    auto-closed by whatever ends it (the next '<slide>', or end of input),
    and bare text sitting directly inside '<slide>' — not wrapped in '<p>' or
    '<li>' — still becomes a bullet instead of being silently dropped. Text
    outside any '<slide>' (chat preamble, unknown tags) is ignored. A ```
    fence that opens before the first '<slide>' is a wrapper around the whole
    answer, not a code sample, and is unwrapped rather than swallowing the
    deck — models wrap their answer in one whether or not you ask them to.

    Returns [] if no slides found. Parsing stops dead at the first
    `APPENDIX_MARKER` HTML comment, so re-exporting a message openPPT already
    appended to (a download link, a diagnostic) never parses that text back
    into slides — even if the model's own last '<slide>' was left unclosed.
    """
    parser = _OutlineParser()
    try:
        parser.feed(_extract_fences(_strip_reasoning(markup)))
        parser.close()
    except Exception:
        pass  # never raise — whatever slides were parsed before the failure still count
    if parser.in_slide:
        parser._close_slide()
    return parser.slides


async def _maybe_await(value):
    """Open WebUI's Files API turned async; a pasted Function must fit both.

    Calling an async Files method without awaiting returns a coroutine that
    never runs — the insert silently doesn't happen and the download 404s.
    """
    return await value if inspect.isawaitable(value) else value


MAX_IMAGE_BYTES = 10 * 1024 * 1024

# An Open WebUI file reference, as its own generated images are linked.
OWUI_FILE_RE = re.compile(r"/api/v1/files/([0-9A-Za-z-]+)(?:/content)?/?(?:\?.*)?$")


def _public_host(host: str) -> bool:
    """True only if every address the host resolves to is publicly routable.

    The URL comes from model output, and a model can be steered by whatever
    document it was given — so the server must not fetch loopback, private,
    or link-local addresses (cloud metadata endpoints live there) on its say.
    """
    try:
        infos = socket.getaddrinfo(host, None)
    except Exception:
        return False
    for info in infos:
        try:
            ip = ipaddress.ip_address(info[4][0].split("%")[0])
        except ValueError:
            return False
        if not ip.is_global:
            return False
    return bool(infos)


class _PublicRedirects(urllib.request.HTTPRedirectHandler):
    """Re-check the host on every redirect, or a public URL could bounce the
    fetch into the private network."""

    def redirect_request(self, req, fp, code, msg, headers, newurl):
        parsed = urllib.parse.urlparse(newurl)
        if parsed.scheme not in ("http", "https") or not _public_host(parsed.hostname or ""):
            return None
        return super().redirect_request(req, fp, code, msg, headers, newurl)


def _fetch_remote_image(url: str):
    parsed = urllib.parse.urlparse(url)
    if not _public_host(parsed.hostname or ""):
        return None
    opener = urllib.request.build_opener(_PublicRedirects)
    request = urllib.request.Request(url, headers={"User-Agent": "openPPT"})
    with opener.open(request, timeout=10) as resp:
        if not (resp.headers.get_content_type() or "").startswith("image/"):
            return None
        data = resp.read(MAX_IMAGE_BYTES + 1)
    return data if len(data) <= MAX_IMAGE_BYTES else None


def _load_image(src: str, images=None, allow_local=False):
    """Image bytes plus pixel size for an <img src>, or None — never raises.

    Accepted sources: anything the caller pre-resolved into `images` (the
    Action resolves Open WebUI file links that way, since that lookup is async
    and user-scoped), a data:image/...;base64 URI, a public http(s) URL, and —
    only with allow_local, i.e. the command-line converter — a local path.
    Whatever comes back must decode as an image: one bad link (an HTML page,
    a text file) used to abort the whole export.
    """
    try:
        src = (src or "").strip()
        data = (images or {}).get(src)
        if data is None:
            if src.startswith("data:image/") and ";base64," in src:
                data = base64.b64decode(src.split(",", 1)[1], validate=False)
            elif src.startswith(("http://", "https://")):
                data = _fetch_remote_image(src)
            elif allow_local and src and not src.startswith("data:"):
                with open(src, "rb") as f:
                    data = f.read(MAX_IMAGE_BYTES + 1)
        if not data or len(data) > MAX_IMAGE_BYTES:
            return None
        from PIL import Image  # python-pptx's own dependency

        with Image.open(io.BytesIO(data)) as im:
            size = im.size
        return data, size
    except Exception:
        return None


def _file_entries(files):
    """(id, filename) for each attachment entry. Open WebUI has shipped both
    a flat {'id','name'} and a nested {'file': {'id','filename'}} shape."""
    for f in files or []:
        if not isinstance(f, dict):
            continue
        inner = f.get("file") if isinstance(f.get("file"), dict) else {}
        fid = inner.get("id") or f.get("id")
        if fid:
            yield fid, (f.get("name") or inner.get("filename") or "")


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
    if role == "section":
        # A divider: the template's own section layout when it has one (its
        # name is the only reliable signal — it looks like a title slide by
        # placeholders), else a title with nothing else to fill.
        for lay, s in scored:
            if "section" in lay.name.lower() and s["title"]:
                return lay
        role = "title_only"
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
    """First fillable body placeholder that isn't the title/subtitle/chrome.

    Matched by placeholder type, never by Python identity: python-pptx hands
    out a fresh proxy object on every access, so an id() comparison never
    matches — and the title used to come back as the "body" on layouts with
    no body of their own (Title Only), which then got dropped as empty.
    """
    for ph in slide.placeholders:
        if _ph_type(ph) in BODY_TYPES:
            return ph
    # Fall back to any non-title, non-chrome, non-subtitle placeholder.
    for ph in slide.placeholders:
        t = _ph_type(ph)
        if t in TITLE_TYPES or t in CHROME_TYPES or t in SUBTITLE_TYPES:
            continue
        return ph
    return None


EMU_PER_PT = 12700
# Default text-frame insets (0.1in each side, 0.05in top and bottom), in
# points summed over both sides — python-pptx's and PowerPoint's default.
INSET_X, INSET_Y = 14.4, 7.2

TITLE_SIZES = (40, 36, 32, 28, 24, 20, 18, 16, 14)
BODY_SIZES = (28, 24, 22, 20, 18, 16, 14, 12, 11, 10)
TABLE_SIZES = (18, 16, 14, 12, 11, 10, 9)
CODE_SIZES = (16, 14, 12, 11, 10, 9, 8)


def _pt(emu) -> float:
    return (emu or 0) / EMU_PER_PT


def _emu(points: float) -> Emu:
    return Emu(int(points * EMU_PER_PT))


def _wrapped_lines(text: str, width_pt: float, size: float, char_em: float = 0.5) -> int:
    """How many lines `text` wraps to in a box `width_pt` wide at `size` pt.

    python-pptx can't measure text, and PowerPoint only recomputes autofit
    when a slide is edited, so the size has to be right on first open. This
    estimates from an average character width — 0.5em errs wide for
    proportional fonts, 0.6em for monospace — so text ends up a little smaller
    than it could be rather than spilling out of its box.
    """
    per_line = max(4, int(max(width_pt, 1) / (size * char_em)))
    return sum(max(1, math.ceil(len(part) / per_line)) for part in text.split("\n"))


def _level_size(size: float, level: int) -> float:
    return max(size - 2 * min(level, 2), 9)


def _bullets_height(bullets, width_pt: float, size: float) -> float:
    total = 0.0
    for level, text in bullets:
        s = _level_size(size, level)
        indent = 18 + 22 * level  # hanging bullet indent + per-level margin
        lines = _wrapped_lines(text, width_pt - indent, s)
        total += lines * s * 1.2 + s * 0.2  # line height + space before
    return total


def _body_size(bullets, rect, max_size: float = 28) -> float:
    """Largest size from BODY_SIZES at which `bullets` fit `rect` (EMU box) —
    a sparse slide keeps big readable text instead of a fixed small size, a
    dense one shrinks until it fits."""
    # 10% slack: renderers that do recompute autofit (LibreOffice, Google
    # Slides) shrink hard on the slightest overflow, and font metrics vary.
    width, height = _pt(rect[2]) - INSET_X, (_pt(rect[3]) - INSET_Y) * 0.9
    for size in BODY_SIZES:
        if size <= max_size and _bullets_height(bullets, width, size) <= height:
            return size
    return BODY_SIZES[-1]


def _template_title_size(prs) -> float:
    """The master's title point size (PowerPoint's default is 44)."""
    try:
        sz = prs.slide_master.element.xpath(
            "./p:txStyles/p:titleStyle/a:lvl1pPr/a:defRPr/@sz"
        )
        return int(sz[0]) / 100 if sz else 44.0
    except Exception:
        return 44.0


def _title_size(text: str, rect, template_size: float = 44.0):
    """None when the title fits its placeholder at the template's own size
    (so the template's styling is left alone), else the largest size that
    fits. Without this, a long title renders at ~44pt across the bullets."""
    width, height = _pt(rect[2]) - INSET_X, _pt(rect[3]) - INSET_Y

    def fits(size):
        return _wrapped_lines(text, width, size, 0.52) * size * 1.15 <= height

    if not rect[2] or fits(template_size):
        return None
    for size in TITLE_SIZES:
        if size < template_size and fits(size):
            return size
    return TITLE_SIZES[-1]


def _fit(text_frame) -> None:
    """Wrap, and set autofit so tools that recompute it (LibreOffice, Google
    Slides) can still shrink text further. PowerPoint itself only recomputes
    autofit on edit, so on first open our explicit Pt() sizes are what fit
    the text — this just avoids overflow markers on the placeholder."""
    text_frame.word_wrap = True
    text_frame.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE


def _size_paragraph(p, size) -> None:
    """Set a paragraph's size on its runs as well as its defaults: PowerPoint
    renders a run's own properties, and doesn't apply a paragraph-level
    default to runs that already exist."""
    p.font.size = Pt(size)
    for run in p.runs:
        run.font.size = Pt(size)


def _rect(shape):
    return (shape.left, shape.top, shape.width, shape.height)


def _set_rect(shape, rect) -> None:
    """Give a shape an explicit box. All four sides at once: a placeholder
    inherits its geometry from the layout until it gets an xfrm of its own,
    and setting just one side writes zeros for the rest — which is how bullets
    used to vanish (width 0) on slides that also had a table or code."""
    left, top, width, height = (Emu(int(v)) for v in rect)
    shape.left, shape.top, shape.width, shape.height = left, top, width, height


def _set_title(slide, text: str, template_size: float = 44.0) -> None:
    ph = _title_placeholder(slide)
    if ph is None or not ph.has_text_frame:
        return
    rect = _rect(ph)
    ph.text_frame.text = text
    _fit(ph.text_frame)
    size = _title_size(text, rect, template_size) if rect[2] else None
    if size:
        for p in ph.text_frame.paragraphs:
            _size_paragraph(p, size)


def _insert_bullet_props(paragraph, element) -> None:
    """Insert a bullet element into a:pPr in schema order (before tabLst /
    defRPr / extLst) — out-of-order children make PowerPoint offer a repair."""
    pPr = paragraph._p.get_or_add_pPr()
    for tag in ("a:buNone", "a:buChar", "a:buAutoNum", "a:buFont"):
        for old in pPr.findall(qn(tag)):
            pPr.remove(old)
    for i, child in enumerate(pPr):
        if child.tag in (qn("a:tabLst"), qn("a:defRPr"), qn("a:extLst")):
            pPr.insert(i, element)
            return
    pPr.append(element)


def _no_bullet(paragraph) -> None:
    pPr = paragraph._p.get_or_add_pPr()
    pPr.set("marL", "0")
    pPr.set("indent", "0")
    _insert_bullet_props(paragraph, pPr.makeelement(qn("a:buNone"), {}))


def _bullet_glyph(paragraph, level: int) -> None:
    """A real bullet on a plain text box paragraph (placeholders get theirs
    from the layout; a text box has none)."""
    pPr = paragraph._p.get_or_add_pPr()
    pPr.set("marL", str(int(Pt(18 + 22 * level))))
    pPr.set("indent", str(-int(Pt(18))))
    _insert_bullet_props(paragraph, pPr.makeelement(qn("a:buChar"), {"char": "•"}))


def _text_items(bullets, runs=None):
    """The text bullets of a slide or column, with their formatting runs
    aligned (image bullets are placed separately)."""
    runs = list(runs or [])
    runs += [None] * (len(bullets) - len(runs))
    pairs = [(b, r) for b, r in zip(bullets, runs) if isinstance(b[1], str)]
    return [b for b, _ in pairs], [r for _, r in pairs]


def _write_runs(p, text, runs) -> None:
    """Paragraph text, keeping <b>/<i>/<u>/<code> as run formatting."""
    if not runs:
        p.text = text
        return
    for chunk, fmt in runs:
        run = p.add_run()
        run.text = chunk
        if "b" in fmt:
            run.font.bold = True
        if "i" in fmt:
            run.font.italic = True
        if "u" in fmt:
            run.font.underline = True
        if "code" in fmt:
            run.font.name = "Consolas"


def _fill_bullets(text_frame, bullets, rect, glyphs=False, runs=None, heading="") -> float:
    """Write bullets into a text frame at the largest size that fits `rect`.
    `heading` (a column's) goes first as a bold, bullet-less line."""
    _fit(text_frame)
    sized = ([(0, heading)] if heading else []) + list(bullets)
    size = _body_size(sized, rect)
    runs = list(runs or []) + [None] * len(bullets)
    j = 0
    if heading:
        p = text_frame.paragraphs[0]
        p.text = heading
        _no_bullet(p)
        _size_paragraph(p, size + 2)
        for run in p.runs:
            run.font.bold = True
        j = 1
    for k, (level, text) in enumerate(bullets):
        p = text_frame.paragraphs[0] if j == 0 else text_frame.add_paragraph()
        j += 1
        _write_runs(p, text, runs[k])
        p.level = level
        _size_paragraph(p, _level_size(size, level))
        if glyphs:
            _bullet_glyph(p, level)
    return size


def _free_rect(prs, slide):
    """Content box for a layout with no body placeholder: under the title,
    aligned with it, down to a bottom margin."""
    sw, sh = prs.slide_width, prs.slide_height
    title = _title_placeholder(slide)
    margin = int(sw * 0.05)
    if title is not None and title.width and title.top is not None:
        left, width = title.left, title.width
        top = title.top + title.height + int(Pt(10))
    else:
        left, width, top = margin, sw - 2 * margin, int(sh * 0.08)
    bottom = int(sh * 0.93)
    return (left, top, width, max(bottom - top, int(sh * 0.3)))


def _content_area(prs, slide):
    """(body placeholder or None, its box) — the box comes from the layout."""
    body = _body_placeholder(slide)
    if body is not None and body.width:
        return body, _rect(body)
    return body, _free_rect(prs, slide)


def _drop_shape(shape) -> None:
    """Remove a shape (an unfilled placeholder renders as an empty prompt box)."""
    shape._element.getparent().remove(shape._element)


def _drop_empty_placeholders(slide) -> None:
    """Remove placeholders nothing was put in. PowerPoint shows each one as a
    'Click to add text' box — e.g. the second column of a 'Two Content'
    layout, or a subtitle box the outline gave no subtitle for."""
    for ph in list(slide.placeholders):
        if _ph_type(ph) in CHROME_TYPES:
            continue
        if ph.has_text_frame and not ph.text_frame.text.strip():
            _drop_shape(ph)


def _add_lead(slide, text: str, area):
    """A content slide's subtitle, as a lead line above its content when the
    layout has no subtitle placeholder (it used to be written into the body
    and then overwritten by the first bullet). Returns the remaining area."""
    left, top, width, height = area
    size = 18
    lines = min(_wrapped_lines(text, _pt(width) - INSET_X, size), 3)
    box_h = int(_emu(lines * size * 1.25 + INSET_Y))
    box = slide.shapes.add_textbox(left, top, width, box_h)
    box.text_frame.word_wrap = True
    p = box.text_frame.paragraphs[0]
    p.text = text
    _size_paragraph(p, size)
    for run in p.runs:
        run.font.italic = True
    gap = int(Pt(6))
    return (left, top + box_h + gap, width, max(height - box_h - gap, int(height * 0.3)))


def _table_layout(rows, width_pt: float, size: float):
    """Column widths (by content length, so a long column isn't squeezed to
    the width of a one-word one) and row heights (by wrapped lines), in pt."""
    cols = max(len(r) for r in rows)
    lens = [max(len(r[c]) if c < len(r) else 0 for r in rows) for c in range(cols)]
    weights = [max(n, 4) ** 0.8 for n in lens]
    widths = [width_pt * w / sum(weights) for w in weights]
    heights = []
    for r in rows:
        lines = max(
            _wrapped_lines(r[c] if c < len(r) else "", widths[c] - INSET_X, size)
            for c in range(cols)
        )
        heights.append(lines * size * 1.2 + INSET_Y + 2)
    return widths, heights


def _table_height(rows, width_pt: float, size: float) -> float:
    return sum(_table_layout(rows, width_pt, size)[1])


def _add_table(slide, rows, rect) -> None:
    left, top, width, height = rect
    width_pt, height_pt = _pt(width), _pt(height)
    for size in TABLE_SIZES:
        widths, heights = _table_layout(rows, width_pt, size)
        if sum(heights) <= height_pt:
            break
    cols = len(widths)
    shape = slide.shapes.add_table(len(rows), cols, left, top, width, _emu(sum(heights)))
    table = shape.table
    col_emu = [int(_emu(w)) for w in widths]
    col_emu[-1] += int(width) - sum(col_emu)  # rounding goes to the last column
    for c, w in enumerate(col_emu):
        table.columns[c].width = Emu(w)
    for r, h in enumerate(heights):
        table.rows[r].height = _emu(h)
    for r, row in enumerate(rows):
        for c in range(cols):
            cell = table.cell(r, c)
            cell.text = row[c] if c < len(row) else ""
            cell.text_frame.word_wrap = True
            for p in cell.text_frame.paragraphs:
                _size_paragraph(p, size)


def _code_size(code: str, rect):
    lines = code.rstrip().splitlines() or [""]
    width, height = _pt(rect[2]) - INSET_X, _pt(rect[3]) - INSET_Y
    for size in CODE_SIZES:
        wrapped = sum(_wrapped_lines(line, width, size, 0.6) for line in lines)
        if wrapped * size * 1.2 <= height:
            break
    return size, wrapped * size * 1.2 + INSET_Y + 4


def _add_code(slide, code, rect) -> None:
    left, top, width, height = rect
    size, needed = _code_size(code, rect)
    box = slide.shapes.add_textbox(left, top, width, min(int(height), int(_emu(needed))))
    box.fill.solid()
    box.fill.fore_color.rgb = RGBColor(0xF4, 0xF4, 0xF4)
    tf = box.text_frame
    tf.word_wrap = True
    for j, text in enumerate(code.rstrip().splitlines() or [""]):
        p = tf.paragraphs[0] if j == 0 else tf.add_paragraph()
        p.text = text
        p.font.name = "Consolas"
        _size_paragraph(p, size)
        for run in p.runs:
            run.font.name = "Consolas"


def _add_pictures(slide, pictures, area) -> None:
    """Fit images into `area` in a grid (up to 3 across), each scaled to keep
    its aspect ratio and centered in its cell — not dropped at a fixed spot
    where they overlapped the bullets and each other."""
    left, top, width, height = (int(v) for v in area)
    cols = min(len(pictures), 3)
    rows = math.ceil(len(pictures) / cols)
    gap = int(Pt(8))
    cell_w = (width - gap * (cols - 1)) / cols
    cell_h = (height - gap * (rows - 1)) / rows
    for k, (data, (px_w, px_h), alt) in enumerate(pictures):
        if not px_w or not px_h:
            continue
        r, c = divmod(k, cols)
        scale = min(cell_w / px_w, cell_h / px_h)
        w, h = int(px_w * scale), int(px_h * scale)
        x = int(left + c * (cell_w + gap) + (cell_w - w) / 2)
        y = int(top + r * (cell_h + gap) + (cell_h - h) / 2)
        try:
            pic = slide.shapes.add_picture(io.BytesIO(data), Emu(x), Emu(y), Emu(w), Emu(h))
            if alt:
                pic._element.nvPicPr.cNvPr.set("descr", alt)
        except Exception:
            continue  # one bad image never costs the deck


def _split_h(rect, fraction):
    """Left `fraction` of a box and the remainder to its right, with a gutter."""
    left, top, width, height = rect
    gutter = int(Pt(16))
    first = int(width * fraction)
    return (left, top, first, height), (left + first + gutter, top, width - first - gutter, height)


def _columns_height(columns, width_pt: float, size: float) -> float:
    n = max(len(columns), 1)
    col_pt = (width_pt - 18 * (n - 1)) / n - INSET_X
    return max(
        _bullets_height(
            ([(0, c["heading"])] if c["heading"] else []) + c["text"], col_pt, size
        ) + (120 if c["pictures"] else 0)
        for c in columns
    ) + INSET_Y


def _stack(area, bullets, table, code, columns=()):
    """Split `area` top-to-bottom among whichever of bullets/columns/table/
    code the slide has, in proportion to how much room each needs at a typical
    size, so a two-line intro doesn't get half the slide above a long table."""
    left, top, width, height = area
    width_pt = _pt(width)
    wanted = []
    # Text is budgeted at a comfortable 22pt (tables at 18pt): when there's
    # room, each block's own fit then lands near that, not at a cramped size.
    if bullets:
        wanted.append(("bullets", _bullets_height(bullets, width_pt - INSET_X, 22) + INSET_Y))
    if columns:
        wanted.append(("columns", _columns_height(columns, width_pt, 22)))
    if table:
        wanted.append(("table", _table_height(table, width_pt, 18)))
    if code:
        wanted.append(("code", _code_size(code, (0, 0, width, height))[1]))
    gutter = int(Pt(10))
    usable = height - gutter * (len(wanted) - 1)
    total = sum(_emu(w) for _, w in wanted) or 1
    rects, y = {}, top
    for key, w in wanted:
        if len(wanted) == 1:
            h = usable
        elif total <= usable:
            # Everything fits: each block gets what it needs plus some slack,
            # top-aligned, rather than being stretched apart down the slide.
            h = int(_emu(w) * min(usable / total, 1.4))
        else:
            h = int(usable * max(_emu(w) / total, 0.2))
        rects[key] = (left, y, width, h)
        y += h + gutter
    # min-share bumps can overshoot; squeeze the last block back inside
    key = wanted[-1][0]
    l, t, w, h = rects[key]
    rects[key] = (l, t, w, max(min(h, top + height - t), int(height * 0.15)))
    return rects


def _columns_layout(prs, n: int):
    """A layout with `n` body placeholders side by side (e.g. 'Two Content'),
    by placeholder count rather than name; else the ordinary content layout."""
    scored = [(lay, _layout_stats(lay)) for lay in prs.slide_layouts]
    for exact in (True, False):
        for lay, s in scored:
            if s["title"] and not s["subtitle"] and (
                s["body"] == n if exact else s["body"] >= n
            ):
                return lay
    return _pick_layout(prs, "content")


def _body_union(prs, slide):
    """The box spanning every body placeholder — for a two-column layout,
    both columns — or the free area under the title when there are none."""
    bodies = [ph for ph in slide.placeholders if _ph_type(ph) in BODY_TYPES and ph.width]
    if not bodies:
        return _free_rect(prs, slide)
    left = min(b.left for b in bodies)
    top = min(b.top for b in bodies)
    right = max(b.left + b.width for b in bodies)
    bottom = max(b.top + b.height for b in bodies)
    return (left, top, right - left, bottom - top)


def _place_columns(slide, columns, area, placeholders) -> None:
    """Side-by-side columns, filled into the layout's own body placeholders
    (so a template's column styling applies) where it has enough, else into
    text boxes. Each column: heading, bullets, then any images below."""
    left, top, width, height = (int(v) for v in area)
    n = len(columns)
    gutter = int(Pt(18))
    col_w = (width - gutter * (n - 1)) // n
    placeholders = sorted(placeholders, key=lambda ph: ph.left or 0)
    for k, col in enumerate(columns):
        rect = (left + k * (col_w + gutter), top, col_w, height)
        text_rect, pic_rect = rect, None
        has_text = bool(col["text"] or col["heading"])
        if col["pictures"] and has_text:
            upper = int(height * 0.55)
            text_rect = (rect[0], top, col_w, upper)
            pic_rect = (rect[0], top + upper + int(Pt(8)), col_w, height - upper - int(Pt(8)))
        elif col["pictures"]:
            pic_rect = rect
        if has_text:
            ph = placeholders[k] if k < len(placeholders) else None
            if ph is not None and ph.has_text_frame:
                _set_rect(ph, text_rect)
                frame, glyphs = ph.text_frame, False
            else:
                frame, glyphs = slide.shapes.add_textbox(*text_rect).text_frame, True
            _fill_bullets(frame, col["text"], text_rect, glyphs, col["runs"], col["heading"])
        if pic_rect is not None:
            _add_pictures(slide, col["pictures"], pic_rect)


def _place_content(prs, slide, body, area, bullets, table, code, pictures,
                   runs=None, columns=()) -> None:
    main, pic_area = area, None
    if pictures and (bullets or table or code or columns):
        main, pic_area = _split_h(area, 0.58)
    elif pictures:
        pic_area = area
    has_main = bool(bullets or table or code or columns)
    rects = _stack(main, bullets, table, code, columns) if has_main else {}
    # A column slide's body placeholders belong to its columns.
    col_phs = [ph for ph in slide.placeholders if _ph_type(ph) in BODY_TYPES] if columns else []

    if bullets:
        if body is not None and body.has_text_frame and not columns:
            _set_rect(body, rects["bullets"])
            _fill_bullets(body.text_frame, bullets, rects["bullets"], runs=runs)
        else:  # layout has no free body placeholder (e.g. picked by name)
            box = slide.shapes.add_textbox(*rects["bullets"])
            _fill_bullets(box.text_frame, bullets, rects["bullets"], glyphs=True, runs=runs)
    elif body is not None and not columns:
        _drop_shape(body)  # nothing to put in it; don't render an empty box
    if columns:
        _place_columns(slide, columns, rects["columns"], col_phs)
    if table:
        _add_table(slide, table, rects["table"])
    if code:
        _add_code(slide, code, rects["code"])
    if pic_area is not None:
        _add_pictures(slide, pictures, pic_area)


DEFAULT_ACCENT = "1F4E79"
WIDE_SLIDE_WIDTH = 12192000  # 13.333in: 16:9 at PowerPoint's 7.5in height


def _hex_color(value, default=DEFAULT_ACCENT) -> str:
    value = (value or "").strip().lstrip("#")
    return value.upper() if re.fullmatch(r"[0-9A-Fa-f]{6}", value) else default


def _mix(hex6: str, toward: str, amount: float) -> str:
    """Blend a color `amount` (0..1) of the way toward another — shades and
    tints derived from the one accent color the admin picks."""
    a = [int(hex6[i:i + 2], 16) for i in (0, 2, 4)]
    b = [int(toward[i:i + 2], 16) for i in (0, 2, 4)]
    return "".join(f"{round(x + (y - x) * amount):02X}" for x, y in zip(a, b))


def _insert_before(parent, child, *later_tags) -> None:
    """Insert `child` before the first existing child named in `later_tags`
    (OOXML children are order-sensitive), else append it."""
    later = {qn(t) for t in later_tags}
    for i, existing in enumerate(parent):
        if existing.tag in later:
            parent.insert(i, child)
            return
    parent.append(child)


def _solid_fill(defRPr, color: str, scheme: bool = False) -> None:
    """Replace a run-properties element's fill with one solid color."""
    for tag in ("a:noFill", "a:solidFill", "a:gradFill", "a:blipFill", "a:pattFill", "a:grpFill"):
        for old in defRPr.findall(qn(tag)):
            defRPr.remove(old)
    fill = defRPr.makeelement(qn("a:solidFill"), {})
    clr = fill.makeelement(qn("a:schemeClr" if scheme else "a:srgbClr"), {"val": color})
    fill.append(clr)
    # a:ln is the only child that may precede the fill
    ln = defRPr.find(qn("a:ln"))
    defRPr.insert(list(defRPr).index(ln) + 1 if ln is not None else 0, fill)


def _level1(parent, tag="a:lvl1pPr"):
    """Get-or-create parent/lvl1pPr and its defRPr."""
    lvl = parent.find(qn(tag))
    if lvl is None:
        lvl = parent.makeelement(qn(tag), {})
        _insert_before(parent, lvl, "a:lvl2pPr", "a:lvl3pPr", "a:extLst")
    rpr = lvl.find(qn("a:defRPr"))
    if rpr is None:
        rpr = lvl.makeelement(qn("a:defRPr"), {})
        _insert_before(lvl, rpr, "a:extLst")
    return lvl, rpr


def _widen(prs) -> None:
    """Make python-pptx's 4:3 default template 16:9. Changing the slide size
    alone leaves every master/layout placeholder where it was, bunched on the
    left, so their x positions and widths are scaled along with it."""
    old = int(prs.slide_width)
    if old >= WIDE_SLIDE_WIDTH * 0.95:
        return
    fx = WIDE_SLIDE_WIDTH / old
    prs.slide_width = Emu(WIDE_SLIDE_WIDTH)
    for el in [prs.slide_master.element] + [lay.element for lay in prs.slide_layouts]:
        for off in el.xpath(".//p:cSld//a:off"):
            off.set("x", str(int(int(off.get("x", "0")) * fx)))
        for ext in el.xpath(".//p:cSld//a:ext[@cx]"):
            ext.set("cx", str(int(int(ext.get("cx")) * fx)))


def _theme_colors(prs, accent: str) -> None:
    """Point the theme's accent1/dk2 at the accent. Tables, charts and
    anything else PowerPoint styles 'Accent 1' pick it up from here."""
    from lxml import etree

    from pptx.opc.constants import RELATIONSHIP_TYPE as RT

    part = prs.slide_master.part.part_related_by(RT.THEME)
    root = etree.fromstring(part.blob)
    scheme = root.find(".//" + qn("a:clrScheme"))
    for name, color in (("accent1", accent), ("dk2", _mix(accent, "000000", 0.35))):
        slot = scheme.find(qn("a:" + name))
        for child in list(slot):
            slot.remove(child)
        slot.append(slot.makeelement(qn("a:srgbClr"), {"val": color}))
    part._blob = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def _accent_bar(spTree, width: int, height: int, color: str, shape_id: int) -> None:
    """A full-width accent strip across the top of every content slide, put
    behind the placeholders on the master."""
    nsdecl = (
        'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
        'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
    )
    from pptx.oxml import parse_xml

    sp = parse_xml(
        f'<p:sp {nsdecl}><p:nvSpPr><p:cNvPr id="{shape_id}" name="openPPT accent bar"/>'
        "<p:cNvSpPr/><p:nvPr userDrawn=\"1\"/></p:nvSpPr><p:spPr>"
        f'<a:xfrm><a:off x="0" y="0"/><a:ext cx="{width}" cy="{height}"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        f'<a:solidFill><a:srgbClr val="{color}"/></a:solidFill><a:ln><a:noFill/></a:ln>'
        "</p:spPr></p:sp>"
    )
    spTree.insert(2, sp)  # after nvGrpSpPr/grpSpPr: drawn first, i.e. behind


def _style_cover_layout(layout, accent: str) -> None:
    """Title-slide / section-header look: accent background, white
    left-aligned title, a light tint of the accent for the subtitle."""
    layout.background.fill.solid()
    layout.background.fill.fore_color.rgb = RGBColor.from_string(accent)
    layout.element.set("showMasterSp", "0")  # no accent bar on an accent page
    # Line every text block up on the title's left edge (python-pptx's
    # default indents the subtitle), so the left-aligned page reads as one.
    title = next((ph for ph in layout.placeholders if _ph_type(ph) in TITLE_TYPES), None)
    for ph in layout.placeholders:
        if title is None or ph is title or _ph_type(ph) in CHROME_TYPES + TITLE_TYPES:
            continue
        if ph.width is not None and title.width is not None:
            ph.left, ph.width = title.left, title.width
    for ph in layout.placeholders:
        t = _ph_type(ph)
        if t in CHROME_TYPES:
            continue
        txBody = ph.element.find(qn("p:txBody"))
        if txBody is None:
            continue
        lst = txBody.find(qn("a:lstStyle"))
        if lst is None:
            lst = txBody.makeelement(qn("a:lstStyle"), {})
            _insert_before(txBody, lst, "a:p")
        lvl, rpr = _level1(lst)
        lvl.set("algn", "l")
        color = "FFFFFF" if t in TITLE_TYPES else _mix(accent, "FFFFFF", 0.8)
        _solid_fill(rpr, color)


def _apply_builtin_theme(prs, accent=None) -> None:
    """The look used when no template is in play — python-pptx's default is
    a plain 4:3 white deck. 16:9, an accent color (theme accent1, which also
    styles tables), left-aligned dark titles, accent bullets, an accent strip
    on content slides, and accent title/section pages. Best effort: a failure
    leaves the plain default, never a failed export."""
    accent = _hex_color(accent)
    try:
        _widen(prs)
        _theme_colors(prs, accent)
        master = prs.slide_master.element
        title_lvl, title_rpr = _level1(master.find(qn("p:txStyles")).find(qn("p:titleStyle")))
        title_lvl.set("algn", "l")
        _solid_fill(title_rpr, "tx2", scheme=True)
        title_rpr.set("b", "1")
        body_style = master.find(qn("p:txStyles")).find(qn("p:bodyStyle"))
        for lvl in body_style:
            if lvl.tag.endswith("pPr") and lvl.find(qn("a:buClr")) is None:
                clr = lvl.makeelement(qn("a:buClr"), {})
                clr.append(clr.makeelement(qn("a:schemeClr"), {"val": "accent1"}))
                _insert_before(
                    lvl, clr, "a:buSzTx", "a:buSzPct", "a:buSzPts", "a:buFontTx",
                    "a:buFont", "a:buNone", "a:buAutoNum", "a:buChar", "a:buBlip",
                    "a:tabLst", "a:defRPr", "a:extLst",
                )
        ids = [int(i) for i in master.xpath(".//p:cNvPr/@id")] or [0]
        _accent_bar(
            master.find(qn("p:cSld")).find(qn("p:spTree")),
            int(prs.slide_width), int(Pt(6)), accent, max(ids) + 1,
        )
        for layout in prs.slide_layouts:
            stats = _layout_stats(layout)
            if (stats["title"] and stats["subtitle"] and not stats["body"]) or (
                "section" in layout.name.lower()
            ):
                _style_cover_layout(layout, accent)
    except Exception as e:
        print(f"[openPPT] built-in theme skipped: {type(e).__name__}: {e}", flush=True)


def _add_slide_number(slide, number: int) -> None:
    """Fill the layout's slide-number placeholder, where the template put it.

    python-pptx's add_slide never copies the date/footer/slide-number
    placeholders, so without this no deck ever had page numbers. The field
    renumbers itself in PowerPoint; the literal is for viewers that don't.
    """
    for layout_ph in slide.slide_layout.placeholders:
        if _ph_type(layout_ph) != PP_PLACEHOLDER.SLIDE_NUMBER:
            continue
        idx = layout_ph.placeholder_format.idx
        slide.shapes.clone_placeholder(layout_ph)
        for ph in slide.placeholders:
            if ph.placeholder_format.idx == idx:
                p = ph.text_frame.paragraphs[0]._p
                fld = p.makeelement(
                    qn("a:fld"),
                    {"id": "{B6F15528-21DE-4FAA-801E-634DDDAF4B2B}", "type": "slidenum"},
                )
                text = fld.makeelement(qn("a:t"), {})
                text.text = str(number)
                fld.append(text)
                _insert_before(p, fld, "a:endParaRPr")
        return


def build_pptx(
    slides: list,
    template=None,
    *,
    images=None,
    allow_local_images=False,
    accent=None,
    slide_numbers=True,
) -> bytes:
    """Build a .pptx from parsed slides.

    template: optional .pptx/.potx as bytes, a stream, a path, or a
    Presentation. When given, the deck inherits that template's theme,
    fonts, and layouts, and each slide's 'layout' name is matched against the
    template's layouts. When omitted, openPPT's built-in 16:9 theme is used,
    colored by `accent` (a hex color; DEFAULT_ACCENT if unset or invalid).

    slide_numbers: number every slide after the first, in the layout's own
    slide-number position (template or built-in).

    images: optional {src: bytes} of image sources the caller already
    resolved (the Action resolves Open WebUI file links this way).
    allow_local_images: let <img src> name a local file — for the command-line
    converter only; in Open WebUI that would read the server's filesystem.
    """
    prs = _as_presentation(template)
    _clear_slides(prs)
    if template is None:
        _apply_builtin_theme(prs, accent)
    title_pt = _template_title_size(prs)

    for i, s in enumerate(slides):
        bullets, runs = _text_items(s["bullets"], s.get("runs"))
        image_refs = [t for _, t in s["bullets"] if isinstance(t, dict)]
        table, code = s.get("table") or [], (s.get("code") or "").strip()

        def pictures_of(refs):
            loaded = [(_load_image(r.get("image"), images, allow_local_images), r) for r in refs]
            return [(data, size, r.get("alt", "")) for (data, size), r in
                    ((l, r) for l, r in loaded if l)]

        columns = []
        for col in s.get("columns") or []:
            text, col_runs = _text_items(col.get("bullets", []), col.get("runs"))
            refs = [t for _, t in col.get("bullets", []) if isinstance(t, dict)]
            if text or refs or col.get("heading"):
                columns.append({
                    "heading": col.get("heading", ""), "text": text,
                    "runs": col_runs, "pictures": pictures_of(refs),
                })
        if len(columns) == 1:  # one "column" is just more bullets
            col = columns.pop()
            if col["heading"]:
                bullets, runs = bullets + [(0, col["heading"])], runs + [[(col["heading"], ("b",))]]
            bullets, runs = bullets + col["text"], runs + col["runs"]
            extra_pictures = col["pictures"]
        else:
            extra_pictures = []

        has_body = bool(bullets or table or code or image_refs or columns or extra_pictures)
        if not has_body:
            role = "title" if i == 0 else "section"
        elif not (bullets or table or code or columns):
            role = "title_only"  # images only: they get the whole content area
        else:
            role = "content"

        layout = _find_layout_by_name(prs, s.get("layout", "")) or (
            _columns_layout(prs, len(columns)) if columns else _pick_layout(prs, role)
        )
        ps = prs.slides.add_slide(layout)
        _set_title(ps, s["title"], title_pt)

        subtitle_ph = _subtitle_placeholder(ps)
        if role in ("title", "section"):
            target = subtitle_ph or _body_placeholder(ps)
            if s["subtitle"] and target is not None and target.has_text_frame:
                target.text_frame.text = s["subtitle"]
        else:
            body, area = _content_area(prs, ps)
            if columns:
                area = _body_union(prs, ps)
            if s["subtitle"]:
                if subtitle_ph is not None:
                    subtitle_ph.text_frame.text = s["subtitle"]
                else:
                    area = _add_lead(ps, s["subtitle"], area)
            pictures = pictures_of(image_refs) + extra_pictures
            _place_content(prs, ps, body, area, bullets, table, code, pictures, runs, columns)

        if s["notes"]:
            ps.notes_slide.notes_text_frame.text = s["notes"]
        _drop_empty_placeholders(ps)
        if slide_numbers and role not in ("title", "section"):
            # not on cover pages: a page number is noise on a title or divider
            _add_slide_number(ps, i + 1)

    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()


def _chat_messages(chat) -> list:
    """Messages out of a saved chat dict, oldest-first.

    Open WebUI keeps the same conversation twice — a flat 'messages' list and
    a 'history.messages' id->message map — and which one is populated varies
    by version, so read whichever holds more. Not "flat if it's non-empty":
    the flat list can lag a turn behind the history map (it's rebuilt from the
    current branch on save), and a flat list holding just the user's message
    would then shadow a history that has the assistant's reply in it.
    """
    if not isinstance(chat, dict):
        return []
    flat = chat.get("messages")
    flat = flat if isinstance(flat, list) else []
    history = chat.get("history")
    hist = history.get("messages") if isinstance(history, dict) else None
    if isinstance(hist, dict):
        hist = list(hist.values())
    elif not isinstance(hist, list):
        hist = []
    return hist if len(hist) > len(flat) else flat


def _parts_text(parts) -> str:
    """Join the 'text' of a list of content parts ({'type','text'} dicts)."""
    texts = []
    for part in parts if isinstance(parts, list) else []:
        if isinstance(part, str):
            texts.append(part)
        elif isinstance(part, dict) and isinstance(part.get("text"), str):
            texts.append(part["text"])
    return "\n".join(t for t in texts if t.strip())


def _message_text(msg) -> str:
    """The text of a chat message, wherever this Open WebUI version keeps it.

    Since v0.10, Open WebUI stores an assistant reply as structured 'output'
    items and renders those instead of 'content', which can be stale or empty.
    The frontend reads 'output' first ('getOutputText(output) || content'), so
    this does the same: 'output' message items, then string 'content', then
    list-shaped (multimodal) 'content' parts.
    """
    if not isinstance(msg, dict):
        return ""
    output = msg.get("output")
    if isinstance(output, list):
        texts = [
            _parts_text(item.get("content"))
            for item in output
            if isinstance(item, dict) and item.get("type") == "message"
        ]
        joined = "\n".join(t for t in texts if t.strip())
        if joined:
            return joined
    content = msg.get("content")
    if isinstance(content, str):
        return content
    return _parts_text(content)


def _strip_appendix(text: str) -> str:
    """Drop whatever openPPT appended to a message on an earlier click.

    The blank line right before the marker is ours too, so trim it along with
    everything after — but only when the marker is actually present, or an
    ordinary message's own trailing whitespace would get eaten.
    """
    before, marker, _ = (text or "").partition(APPENDIX_MARKER)
    return before.rstrip() if marker else (text or "")


def _clicked_text(messages: list, target) -> str:
    """Text of the message the button was clicked on, minus our appendix."""
    for msg in messages or []:
        if isinstance(msg, dict) and target is not None and msg.get("id") == target:
            return _strip_appendix(_message_text(msg))
    return ""


def _shape(messages: list) -> str:
    """'user:str(21), assistant:list(2)' — role and content shape per message.

    Read only when no outline was found, and the four ways that happens need
    different fixes: content arriving as a list (multimodal), as an empty
    string, under no assistant role at all, or already consumed by the
    appendix strip. The character count alone can't tell them apart.
    """
    parts = []
    for msg in (messages or [])[-4:]:
        if not isinstance(msg, dict):
            parts.append(f"{type(msg).__name__}!")
            continue
        body = msg.get("content")
        size = len(body) if hasattr(body, "__len__") else "?"
        shape = f"{msg.get('role', '?')}:{type(body).__name__}({size})"
        if isinstance(msg.get("output"), list):
            shape += f"+output({len(msg['output'])})"
        parts.append(shape)
    return ", ".join(parts) or "(none)"


def _saved_shape(chat) -> str:
    """'flat=1 hist=3' — how many messages each of the saved chat's two message
    stores holds, or why there's no saved copy at all.

    Read only when the request carried no outline and the saved-chat fallback
    didn't rescue it, where three very different failures look identical from
    the outside: the chat lookup failed (Open WebUI API drift, temporary chat),
    the chat is saved but the assistant turn isn't in it yet, or both stores
    disagree and _chat_messages read the emptier one.
    """
    if chat is None:
        return "none"
    if not isinstance(chat, dict):
        return f"{type(chat).__name__}!"

    def size(v):
        return len(v) if hasattr(v, "__len__") else "?"

    history = chat.get("history")
    hist = history.get("messages") if isinstance(history, dict) else None
    return f"flat={size(chat.get('messages'))} hist={size(hist)}"


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
        if not isinstance(msg, dict) or msg.get("role") != "assistant":
            continue
        # Drop our own appended output (download link, diagnostic) so a
        # re-click echoes back what the model wrote, not what we last posted.
        text = _strip_appendix(_message_text(msg))
        if target is not None and msg.get("id") == target:
            candidates.insert(0, text)
        else:
            candidates.append(text)
    for text in candidates:
        if parse_outline(text):
            return text
    return candidates[0] if candidates else ""


def _download_js(urls: list, name: str) -> str:
    """Browser-side code that downloads the deck through the user's session.

    Run via Open WebUI's 'execute' event. Tries each URL in turn (relative
    first, since the server's idea of its own base URL is wrong behind a
    reverse proxy), and returns a short result string instead of throwing —
    a thrown error never answers the event call, which would stall the export
    until AUTO_DOWNLOAD_TIMEOUT.
    """
    return f"""
const urls = {json.dumps(urls)};
const name = {json.dumps(name)};
for (const u of urls) {{
  try {{
    const r = await fetch(u, {{
      headers: {{ Authorization: 'Bearer ' + (localStorage.token || '') }},
      credentials: 'include'
    }});
    if (!r.ok) continue;
    const blob = await r.blob();
    const a = document.createElement('a');
    a.href = window.URL.createObjectURL(blob);
    a.download = name;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => window.URL.revokeObjectURL(a.href), 30000);
    return 'downloaded';
  }} catch (e) {{}}
}}
return 'failed';
"""


def _image_refs(slides) -> set:
    """Every <img src> in the deck, columns included."""
    lists = [s.get("bullets", []) for s in slides]
    lists += [c.get("bullets", []) for s in slides for c in s.get("columns") or []]
    return {ref.get("image", "") for bullets in lists for _, ref in bullets if isinstance(ref, dict)}


class Action:
    class Valves(BaseModel):
        default_template: str = Field(
            default="",
            description="Template for decks when no .pptx/.potx is attached to the chat: an absolute path on the Open WebUI server, or the id of a file uploaded to Open WebUI. Empty = openPPT's built-in 16:9 theme. A template attached to the chat always wins.",
        )
        accent_color: str = Field(
            default=DEFAULT_ACCENT,
            description="Hex color (e.g. 1F4E79) for the built-in theme: title pages, titles, bullets, tables. Ignored when a template is used.",
        )
        slide_numbers: bool = Field(
            default=True,
            description="Number every slide except title and section pages.",
        )
        auto_download: bool = Field(
            default=True,
            description="Download the .pptx in the browser as soon as it's built, in addition to the attachment chip and link under the message.",
        )

    def __init__(self):
        self.valves = self.Valves()

    async def action(
        self,
        body: dict,
        __user__: dict = None,
        __event_emitter__=None,
        __event_call__=None,
        __request__=None,
    ):
        async def emit(event):
            if __event_emitter__:
                await __event_emitter__(event)

        clicked = _clicked_text(body.get("messages", []), body.get("id"))

        async def append_to_message(appendix):
            """Add openPPT's text under the clicked message, visibly.

            Returning {'messages': [...]} from the action is Open WebUI's own
            contract for editing the message: the frontend merges it into the
            chat and, since v0.12, syncs it into the structured 'output' the
            reply is rendered from. A 'message' event only appends to the
            legacy 'content', which v0.10+ doesn't render — the download link
            was being written somewhere nobody could see it.

            Only done when the clicked text is known: returning just the
            appendix would replace the model's reply with it. Otherwise fall
            back to the 'message' event, which can only ever append.
            """
            if clicked.strip() and body.get("id"):
                return {"messages": [{"id": body["id"], "content": clicked + appendix}]}
            await emit({"type": "message", "data": {"content": appendix}})
            return None

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
            saved_chat = None
            if not slides:
                # The body's copy of the conversation is unreliable: some Open
                # WebUI versions POST the action assistant messages whose
                # 'content' is an empty string, so there is nothing to parse and
                # every export fails while the chat plainly holds an outline.
                # Fall back to the stored chat, which is also where the template
                # has had to come from since v0.5.1. Only on failure, so a body
                # that does carry the outline costs no extra lookup.
                saved_chat = await self._saved_chat(body)
                stored = _chat_messages(saved_chat)
                if stored:
                    from_chat = _pick_outline(stored, target)
                    slides = parse_outline(from_chat)
                    if slides or not content:
                        messages, content = stored, from_chat
            if not slides:
                # Temporary Chats are never written to the Chats table, so the
                # fallback above can never find one — and the live request body
                # is the only copy left, which different Open WebUI versions
                # populate with the assistant's text inconsistently (missing
                # message, empty string) since it's unsaved.
                # No stored copy at all — whether the lookup missed or there was
                # no chat_id to look up with. A missing chat_id is the strongest
                # unsaved-chat signal there is, so it must not suppress the hint.
                temp_chat = saved_chat is None
                # "Fix your outline" is the wrong advice when there is no reply
                # to have an outline in — and that's a distinct failure: neither
                # the request nor the saved chat holds an assistant message, so
                # the model's text never reached openPPT at all. Blaming the
                # outline sends people off rewriting a reply that was fine.
                no_reply = not any(
                    isinstance(m, dict) and m.get("role") == "assistant"
                    for m in list(messages or []) + _chat_messages(saved_chat)
                )
                if no_reply:
                    help_text = (
                        "openPPT never received the model's reply. Neither this "
                        "export request nor Open WebUI's saved copy of the chat "
                        "contains an assistant message, so there was no text to "
                        "read — this isn't about the outline's format. Reload the "
                        "page: if the reply disappears, Open WebUI isn't saving "
                        "it, so check that this isn't a Temporary Chat and that "
                        "chat history saving is on."
                    )
                else:
                    # Deliberately describes the grammar without writing a real
                    # tag: this text is posted back into the chat, and a literal
                    # slide element here would make the diagnostic itself parse
                    # as a deck on the next click (see the v0.3.5 guard in
                    # test_no_outline_diagnostic_is_not_itself_slide_shaped).
                    help_text = (
                        "openPPT: nothing slide-shaped in that message. Ask the model "
                        "for an HTML outline — one 'slide' element per slide, each "
                        "carrying a title attribute, with 'ul'/'li' bullets inside it."
                    )
                if temp_chat:
                    help_text += (
                        " This looks like a Temporary Chat — Open WebUI doesn't "
                        "save those, so openPPT has no reliable copy of the reply "
                        "to read. Turn off Temporary Chat and try again."
                    )
                # The toast/status is the only channel some deployments show —
                # the appended chat block below can go unseen — so lead with the
                # evidence, not the advice: the version proves the paste is
                # current, and 0 chars read (vs. N) separates "the action was
                # handed no text" from "the parser rejected real text". The
                # Temporary Chat note is the one exception worth repeating here.
                head = re.sub(r"\s+", " ", content[:80]).strip()
                temp_chat_note = (
                    "; looks like a Temporary Chat — turn it off and retry"
                    if temp_chat else ""
                )
                lead = (
                    "no assistant reply reached openPPT (not the outline's fault)"
                    if no_reply else "nothing slide-shaped here"
                )
                toast = (
                    f"openPPT {VERSION}: {lead} — read "
                    f"{len(content)} chars from {len(messages)} messages "
                    f"[{_shape(messages)}]; chat_id "
                    f"{'present' if body.get('chat_id') else 'MISSING'}; "
                    f"saved {_saved_shape(saved_chat)}; "
                    f"starts: {head or '(empty)'}{temp_chat_note}"
                )
                await notify(toast, "warning")
                await status(toast, done=True)
                # Toasts and the status line get truncated, so the detail goes in
                # the chat. Written so parse_outline finds nothing in it: no '---'
                # rule, no bullet list, no lone '**bold**' line — otherwise the
                # next click exports this diagnostic as a deck (v0.3.5 did).
                seen = content[:400] if content else "(nothing — the message was empty)"
                return await append_to_message(
                    f"\n\n{APPENDIX_MARKER}\n"
                    f"**openPPT {VERSION} diagnostic** — {help_text} "
                    f"(messages in request: {len(messages)}, clicked id: "
                    f"`{target}`, read {len(content)} characters)\n\n"
                    f"What openPPT read from that message:\n\n"
                    f"```\n{seen}\n```\n"
                )

            # Precedence: a template attached to the chat, then the admin's
            # default_template Valve, then the built-in theme.
            template = await self._find_template(body, __user__)
            note = " on your template" if template is not None else ""
            if template is None:
                template = await self._default_template()
                note = " on the default template" if template is not None else ""
            wants_layout = any(s.get("layout") for s in slides)
            if wants_layout and template is None:
                await status(
                    "This deck names layouts but no .pptx/.potx template is "
                    "attached — using the built-in theme.",
                )

            await status(f"Building PowerPoint ({len(slides)} slides{note})…")
            data = build_pptx(
                slides,
                template=template,
                images=await self._owui_images(slides, __user__),
                accent=self.valves.accent_color,
                slide_numbers=self.valves.slide_numbers,
            )

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

            # Three delivery paths on purpose, because every Open WebUI version
            # drops at least one: v0.10+ renders the 'files' chip but not text
            # appended to 'content'; v0.9 and older render appended text but
            # only image chips. The browser auto-download works on both (v0.6.5
            # onward). Each is isolated so one failing can't lose the others.
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

            downloaded = False
            if self.valves.auto_download and __event_call__ is not None:
                urls = [url] + ([full_url] if full_url != url else [])
                try:
                    answer = await asyncio.wait_for(
                        __event_call__(
                            {"type": "execute", "data": {"code": _download_js(urls, name)}}
                        ),
                        AUTO_DOWNLOAD_TIMEOUT,
                    )
                    downloaded = answer == "downloaded"
                except Exception as e:  # incl. timeout: the chip and link remain
                    print(f"[openPPT] auto-download skipped: {type(e).__name__}: {e}", flush=True)

            done = f"{name} {'downloaded' if downloaded else 'ready'} — {full_url}"
            await notify(done, "success")
            await status(done, done=True)
            # The marker line lets a future export ignore everything below it
            # (see parse_outline), and lets a re-export replace this link rather
            # than stack another under it. The full URL repeated on its own line
            # is copy-pasteable even where neither the chip nor the link renders.
            return await append_to_message(
                f"\n\n{APPENDIX_MARKER}\n📊 [Download {name}]({full_url})"
                f"\n\n{full_url}"
            )
        except Exception as e:
            msg = f"openPPT export failed: {type(e).__name__}: {e}"
            await notify(msg, "error")
            await status(msg, done=True)

    async def _saved_chat(self, body: dict):
        """The stored chat dict behind this action, or None.

        Open WebUI strips the action request body down hard before POSTing to
        /api/chat/actions — it carries no 'files' at all, and on some versions
        the assistant message's 'content' comes through empty too. The stored
        chat is the only copy that reliably has both.
        """
        if not body.get("chat_id"):
            return None
        try:
            from open_webui.models.chats import Chats

            chat = await _maybe_await(Chats.get_chat_by_id(body["chat_id"]))
            return getattr(chat, "chat", None)
        except Exception as e:
            print(f"[openPPT] chat lookup failed: {type(e).__name__}: {e}", flush=True)
            return None

    async def _default_template(self):
        """Bytes of the admin's default_template Valve (a server path or an
        Open WebUI file id), or None. A bad setting is logged, not fatal."""
        ref = (self.valves.default_template or "").strip()
        if not ref:
            return None
        try:
            if os.path.isfile(ref):
                with open(ref, "rb") as fh:
                    return fh.read()
            from open_webui.models.files import Files
            from open_webui.storage.provider import Storage

            rec = await _maybe_await(Files.get_file_by_id(ref))
            if rec is not None:
                with open(Storage.get_file(rec.path), "rb") as fh:
                    return fh.read()
        except Exception as e:
            print(f"[openPPT] default_template unusable: {type(e).__name__}: {e}", flush=True)
        else:
            print(f"[openPPT] default_template not found: {ref}", flush=True)
        return None

    async def _owui_images(self, slides, user) -> dict:
        """{src: bytes} for <img> sources that are Open WebUI file links —
        how its own generated and uploaded images are referenced. Read
        straight from storage (fetching them over HTTP would need the user's
        token, and the server's own address is private). Scoped to the
        requesting user's files; admins can reach any file, as in the UI."""
        found = {}
        user = user or {}
        for src in _image_refs(slides):
            match = OWUI_FILE_RE.search(src or "")
            if not match:
                continue
            try:
                from open_webui.models.files import Files
                from open_webui.storage.provider import Storage

                fid, rec = match.group(1), None
                scoped = getattr(Files, "get_file_by_id_and_user_id", None)
                if scoped is not None and user.get("id"):
                    rec = await _maybe_await(scoped(fid, user["id"]))
                if rec is None and user.get("role") == "admin":
                    rec = await _maybe_await(Files.get_file_by_id(fid))
                if rec is None:
                    continue
                with open(Storage.get_file(rec.path), "rb") as fh:
                    found[src] = fh.read(MAX_IMAGE_BYTES + 1)
            except Exception as e:
                print(f"[openPPT] image {src} skipped: {type(e).__name__}: {e}", flush=True)
        return found

    async def _find_template(self, body: dict, __user__: dict):
        """Return template bytes from the most recent .pptx/.potx attached to
        the chat, or None.

        The action request body has no attachments in it: Open WebUI maps every
        message down to id/role/content/info/timestamp before POSTing to
        /api/chat/actions, and sends no top-level 'files'. So the attachment is
        read from the *saved* chat (body['chat_id']) — which keeps both the
        per-message 'files' and the chat-level 'files' list. The body is still
        scanned first, for versions or callers that do pass files along.
        """
        try:
            from open_webui.models.files import Files
            from open_webui.storage.provider import Storage
        except Exception:
            return None

        def file_ids(container):
            """(id, name) for each file in a {'messages': [...], 'files': [...]}
            dict — message attachments newest-first, then chat-level ones."""
            if not isinstance(container, dict):
                return
            for msg in reversed(_chat_messages(container)):
                if isinstance(msg, dict):
                    yield from _file_entries(msg.get("files"))
            yield from _file_entries(container.get("files"))

        sources = [body, await self._saved_chat(body)]

        for fid, hint in (e for src in sources for e in file_ids(src)):
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
