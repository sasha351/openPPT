"""
title: PowerPoint Template Primer
author: openPPT
version: 0.3.0
requirements: python-pptx
description: When you attach a .pptx/.potx template to the chat, this primes the model — before it answers — with your template's layout names and the openPPT outline format, and tells it to organize the content you provided into a logically flowing deck that fills those layouts. Without a template, it still primes deck/presentation requests with the same outline format and a content-quality bar. Pair it with the "Export to PowerPoint" action: dump your content, and the model formats a deck you can export in one click. Works with any model (no tool calling needed).
"""

import io
import re

from pptx import Presentation
from pptx.enum.shapes import PP_PLACEHOLDER
from pydantic import BaseModel, Field

# Marker so we never inject the primer twice into the same request.
SENTINEL = "[openPPT template primer]"

# Loose signal that the user is asking for a deck, used only when no
# template is attached (template presence is signal enough on its own).
_DECK_KEYWORDS_RE = re.compile(
    r"\b(deck|presentation|slides?|slideshow|powerpoint|pptx)\b", re.IGNORECASE
)

_QUALITY_BAR = (
    "Make every bullet concrete: one specific number, name, date, or action "
    'verb — never a topic label. If a bullet could describe any project '
    'unchanged ("Improved efficiency", "Key considerations"), replace it '
    "with the real fact from the content, or cut it."
)

TITLE_TYPES = (PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.CENTER_TITLE)
SUBTITLE_TYPES = (PP_PLACEHOLDER.SUBTITLE,)
BODY_TYPES = (PP_PLACEHOLDER.BODY, PP_PLACEHOLDER.OBJECT)


def _ph_counts(layout):
    title = subtitle = body = 0
    for ph in layout.placeholders:
        try:
            t = ph.placeholder_format.type
        except Exception:
            t = None
        if t in TITLE_TYPES:
            title += 1
        elif t in SUBTITLE_TYPES:
            subtitle += 1
        elif t in BODY_TYPES:
            body += 1
    return title, subtitle, body


def _describe_layout(layout) -> str:
    """A short 'Name — what it's good for' hint from the layout's placeholders."""
    title, subtitle, body = _ph_counts(layout)
    if title and subtitle and not body:
        hint = "title slide (title + subtitle)"
    elif title and not subtitle and not body:
        hint = "divider / section break — title only, no bullets"
    elif title and body >= 3:
        hint = "comparison / multi-column"
    elif title and body == 2:
        hint = "two columns of bullets"
    elif title and body == 1:
        hint = "title + bullet list"
    elif not title and not body:
        hint = "blank"
    else:
        hint = "title + content"
    return f"- {layout.name} — {hint}"


def _layouts_from_bytes(data: bytes) -> list:
    prs = Presentation(io.BytesIO(data))
    return [_describe_layout(lay) for lay in prs.slide_layouts]


def _template_primer(layout_lines: list) -> str:
    layouts = "\n".join(layout_lines)
    return f"""{SENTINEL}
The user attached a PowerPoint template and provided content to turn into a
presentation. Build the deck FROM the content they gave you — reorganize,
group, and summarize it into slides that flow logically. Do not invent facts
that aren't in their content; you may condense and rephrase.

Respond with ONLY a Markdown outline in this exact format (no prose around it):

# Deck Title
## One-line subtitle

# First Slide Title @layout: <layout name>
- Short bullet point
- Another bullet point
  - Sub-point (indent two spaces)
Notes: extra detail for the speaker (optional, kept off the slide)

Formatting rules:
- Start every slide with '# '. The first slide (title + subtitle, no bullets)
  is the title slide.
- 3-6 bullets per content slide, each under ~12 words. Push detail into 'Notes:'.
- Open a new topic with a divider slide, and lay out the deck so it flows:
  title → agenda/overview → grouped sections → summary/next steps.
- Embed an image with a bullet '![alt](url)' when the content includes one.
- Put tabular data in a Markdown '| col | col |' table instead of bullets.
- Put commands or code in a ``` fence; it renders as a monospace box.
- {_QUALITY_BAR}

Choose a layout for each slide by adding '@layout: <name>' to its heading,
using ONLY these layouts from the attached template:
{layouts}

When the user later asks for changes, reply with the full revised outline."""


def _no_template_primer() -> str:
    return f"""{SENTINEL}
The user asked for a deck or presentation. Build it FROM the content already
in this conversation — reorganize, group, and condense it into slides that
flow logically. Do not invent facts that aren't there.

Respond with ONLY a Markdown outline in this exact format (no prose around it):

# Deck Title
## One-line subtitle

# First Slide Title
- Short bullet point
- Another bullet point
  - Sub-point (indent two spaces)
Notes: extra detail for the speaker (optional, kept off the slide)

Formatting rules:
- Start every slide with '# '. The first slide (title + subtitle, no bullets)
  is the title slide.
- 3-6 bullets per content slide, each under ~12 words. Push detail into 'Notes:'.
- Flow: title -> agenda/overview -> grouped sections -> summary/next steps.
- Embed an image with a bullet '![alt](url)' when the content includes one.
- Put tabular data in a Markdown '| col | col |' table instead of bullets.
- Put commands or code in a ``` fence; it renders as a monospace box.
- {_QUALITY_BAR}

If this isn't a deck/presentation request, ignore this and answer normally.
When the user later asks for changes, reply with the full revised outline."""


def _primer(layout_lines: list = None) -> str:
    return _template_primer(layout_lines) if layout_lines else _no_template_primer()


def _looks_like_deck_request(messages: list) -> bool:
    for m in reversed(messages or []):
        if m.get("role") == "user":
            return bool(_DECK_KEYWORDS_RE.search(m.get("content") or ""))
    return False


class Filter:
    class Valves(BaseModel):
        enabled: bool = Field(
            default=True,
            description="Prime the model whenever a .pptx/.potx template is attached.",
        )
        prime_without_template: bool = Field(
            default=True,
            description="Also prime deck/presentation requests when no template is attached, using a template-less primer.",
        )
        priority: int = Field(
            default=0, description="Filter execution priority (lower runs first)."
        )

    def __init__(self):
        self.valves = self.Valves()

    def inlet(self, body: dict, __user__: dict = None) -> dict:
        """Before the model answers: if a template is attached, inject a system
        message listing its layouts and the outline spec; otherwise, for a
        deck/presentation request, inject the template-less version of the
        same spec. Never raises — on any problem it returns the request
        untouched so the chat is unaffected."""
        try:
            if not self.valves.enabled:
                return body
            messages = body.get("messages", []) or []
            if any(
                m.get("role") == "system" and SENTINEL in (m.get("content") or "")
                for m in messages
            ):
                return body  # already primed this request

            data = self._find_template_bytes(body)
            if data is not None:
                layout_lines = _layouts_from_bytes(data)
                if not layout_lines:
                    return body
                content = _primer(layout_lines)
            elif self.valves.prime_without_template and _looks_like_deck_request(
                messages
            ):
                content = _primer()
            else:
                return body

            primer = {"role": "system", "content": content}
            # Sit right after an existing system prompt, else lead the messages.
            idx = 0
            for i, m in enumerate(messages):
                if m.get("role") == "system":
                    idx = i + 1
                else:
                    break
            messages.insert(idx, primer)
            body["messages"] = messages
        except Exception:
            return body
        return body

    def _find_template_bytes(self, body: dict):
        """Newest-first, return bytes of the first attached .pptx/.potx, or None.

        Duplicated from the export action on purpose: Open WebUI functions are
        pasted in individually and can't import one another.
        """
        try:
            from open_webui.models.files import Files
            from open_webui.storage.provider import Storage
        except Exception:
            return None

        def file_ids():
            for msg in reversed(body.get("messages", []) or []):
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
