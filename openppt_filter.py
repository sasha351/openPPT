"""
title: PowerPoint Template Primer
author: openPPT
version: 0.1.0
requirements: python-pptx
description: When you attach a .pptx/.potx template to the chat, this primes the model — before it answers — with your template's layout names and the openPPT outline format, and tells it to organize the content you provided into a logically flowing deck that fills those layouts. Pair it with the "Export to PowerPoint" action: attach a template, dump your content, and the model formats a deck you can export in one click. Works with any model (no tool calling needed).
"""

import io

from pptx import Presentation
from pptx.enum.shapes import PP_PLACEHOLDER
from pydantic import BaseModel, Field

# Marker so we never inject the primer twice into the same request.
SENTINEL = "[openPPT template primer]"

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


def _primer(layout_lines: list) -> str:
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

Choose a layout for each slide by adding '@layout: <name>' to its heading,
using ONLY these layouts from the attached template:
{layouts}

When the user later asks for changes, reply with the full revised outline."""


class Filter:
    class Valves(BaseModel):
        enabled: bool = Field(
            default=True,
            description="Prime the model whenever a .pptx/.potx template is attached.",
        )
        priority: int = Field(
            default=0, description="Filter execution priority (lower runs first)."
        )

    def __init__(self):
        self.valves = self.Valves()

    def inlet(self, body: dict, __user__: dict = None) -> dict:
        """Before the model answers: if a template is attached, inject a system
        message listing its layouts and the outline spec. Never raises — on any
        problem it returns the request untouched so the chat is unaffected."""
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
            if data is None:
                return body

            layout_lines = _layouts_from_bytes(data)
            if not layout_lines:
                return body

            primer = {"role": "system", "content": _primer(layout_lines)}
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
