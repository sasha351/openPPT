"""
title: PowerPoint Template Primer
author: openPPT
version: 0.5.0
requirements: python-pptx
description: When you attach a .pptx/.potx template to the chat, this primes the model — before it answers — with your template's layout names and the openPPT HTML outline format, and tells it to distill the content you provided into a logically flowing deck that fills those layouts. Without a template, it still primes deck/presentation requests — and follow-up edits to a deck already in the chat — with the same outline format and a content-quality bar. Pair it with the "Export to PowerPoint" action: dump your content, and the model formats a deck you can export in one click. Works with any model (no tool calling needed).
"""

import inspect
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

_QUALITY_BAR = """- Make every bullet concrete: one specific number, name, date, or action verb
  — never a topic label. If a bullet could describe any project unchanged
  ("Improved efficiency", "Key considerations"), replace it with the real
  fact from the content, or cut it.
- Give each slide exactly one point a viewer takes away, not a paragraph
  chopped into fragments. If two bullets are making the same point, merge
  them or cut one.
- Group by theme, not by the source's paragraph order — pull related facts
  together even if they were scattered across the input, and drop anything
  restating a point already made on an earlier slide.
- Size the deck to how much the content actually supports: a short input
  makes a short deck. Padding it out with filler slides ("Overview",
  "Conclusion") that add no new information is worse than a shorter deck.
- Push supporting detail, caveats, and sourcing into that slide's notes
  instead of cramming them onto the slide — the slide is the headline, the
  notes are the explanation."""

TITLE_TYPES = (PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.CENTER_TITLE)
SUBTITLE_TYPES = (PP_PLACEHOLDER.SUBTITLE,)
BODY_TYPES = (PP_PLACEHOLDER.BODY, PP_PLACEHOLDER.OBJECT)


async def _maybe_await(value):
    """Open WebUI's Files API is async since its async-DB move (v0.12 included)
    and sync before it; a pasted Function must fit both. Calling the async one
    without awaiting returns a coroutine, so the template was never found and
    the layout-aware primer never fired. Duplicated from the export action on
    purpose: Functions can't import one another."""
    return await value if inspect.isawaitable(value) else value


def _text(content) -> str:
    """A message's text, whether 'content' is a string or a list of parts
    (multimodal requests — e.g. Gemma with an image attached)."""
    if isinstance(content, str):
        return content
    if isinstance(content, list):
        return "\n".join(
            p.get("text", "") if isinstance(p, dict) else str(p)
            for p in content
            if isinstance(p, (dict, str))
        )
    return ""


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
        hint = "comparison / multi-column — use <column> elements"
    elif title and body == 2:
        hint = "two columns of bullets — use two <column> elements"
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


_FORMAT_RULES = """- Wrap every slide in <slide title="...">...</slide>. The first slide has
  only title and subtitle attributes, no body — that's the title slide. A
  later slide with no body is a section divider.
- 3-6 <li> bullets per content slide, each under ~12 words. Nest a <ul> inside
  an <li> for a sub-point. <b>, <i> and <code> inside a bullet are kept.
- Side-by-side comparison (pros/cons, before/after, options): put each side
  in <column heading="...">...</column> with its own <ul> of <li> inside.
- Embed an image in place of a bullet: <img src="url" alt="description">.
- Put tabular data in <table><tr><th>...</th></tr><tr><td>...</td></tr></table>,
  not bullets.
- Put commands or code in <pre><code>...</code></pre>; it renders as a
  monospace box. Escape literal <, > and & in any text (as &lt; &gt; &amp;)
  so it can't be mistaken for markup.
- Respond with ONLY the <slide> elements — no prose before, after, or between
  them, and no ```html fence wrapping them. If you think before answering,
  the outline still goes in full in your final answer."""


def _template_primer(layout_lines: list) -> str:
    layouts = "\n".join(layout_lines)
    return f"""{SENTINEL}
The user attached a PowerPoint template and provided content to turn into a
presentation. Distill the content they gave you into a deck — pull out the
substance, group it by theme, and cut anything that doesn't earn its place.
Do not invent facts that aren't in their content; you may condense and
rephrase.

Respond with ONLY an HTML outline in this exact format (no prose around it):

<slide title="Deck Title" subtitle="One-line subtitle"></slide>

<slide title="First Slide Title" layout="<layout name>">
<ul>
<li>Short bullet point</li>
<li>Another bullet point
<ul><li>Sub-point</li></ul>
</li>
</ul>
<notes>extra detail for the speaker (optional, kept off the slide)</notes>
</slide>

Formatting rules:
{_FORMAT_RULES}
- Open a new topic with a divider slide (layout with a title and nothing
  else), and lay out the deck so it flows: title → agenda/overview → grouped
  sections → summary/next steps.
{_QUALITY_BAR}

Choose a layout for each slide with the layout="<name>" attribute, using ONLY
these layouts from the attached template:
{layouts}

If the user's latest message isn't about the deck, ignore this and answer
normally. When the user later asks for changes, reply with the full revised
outline."""


def _no_template_primer() -> str:
    return f"""{SENTINEL}
The user asked for a deck or presentation. Distill the content already in
this conversation into a deck — pull out the substance, group it by theme,
and cut anything that doesn't earn its place. Do not invent facts that
aren't there.

Respond with ONLY an HTML outline in this exact format (no prose around it):

<slide title="Deck Title" subtitle="One-line subtitle"></slide>

<slide title="First Slide Title">
<ul>
<li>Short bullet point</li>
<li>Another bullet point
<ul><li>Sub-point</li></ul>
</li>
</ul>
<notes>extra detail for the speaker (optional, kept off the slide)</notes>
</slide>

Formatting rules:
{_FORMAT_RULES}
- Flow: title -> agenda/overview -> grouped sections -> summary/next steps.
{_QUALITY_BAR}

If this isn't a deck/presentation request, ignore this and answer normally.
When the user later asks for changes, reply with the full revised outline."""


def _primer(layout_lines: list = None) -> str:
    return _template_primer(layout_lines) if layout_lines else _no_template_primer()


def _looks_like_deck_request(messages: list) -> bool:
    for m in reversed(messages or []):
        if isinstance(m, dict) and m.get("role") == "user":
            return bool(_DECK_KEYWORDS_RE.search(_text(m.get("content"))))
    return False


def _deck_in_progress(messages: list) -> bool:
    """An earlier reply already holds an outline, so this turn is most likely
    an edit ("make slide 3 shorter") — which names no deck keyword, and
    without the primer the model drifts back to prose or Markdown."""
    return any(
        isinstance(m, dict)
        and m.get("role") == "assistant"
        and "<slide" in _text(m.get("content")).lower()
        for m in messages or []
    )


def _add_primer(messages: list, primer: str) -> list:
    """Append the primer to the leading system message, or start one.

    Never a second system message: Gemma's and Nemotron's chat templates (and
    many others) only accept a system message at position 0, and one later
    in the list is rejected or silently dropped by the serving stack.
    """
    if messages and isinstance(messages[0], dict) and messages[0].get("role") == "system":
        first = dict(messages[0])
        content = first.get("content")
        if isinstance(content, list):
            first["content"] = content + [{"type": "text", "text": primer}]
        else:
            existing = _text(content).rstrip()
            first["content"] = f"{existing}\n\n{primer}" if existing else primer
        return [first] + list(messages[1:])
    return [{"role": "system", "content": primer}] + list(messages)


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


def _chat_messages(chat) -> list:
    """Messages of a saved chat — the fuller of its flat list and history map."""
    if not isinstance(chat, dict):
        return []
    flat = chat.get("messages") if isinstance(chat.get("messages"), list) else []
    history = chat.get("history")
    hist = history.get("messages") if isinstance(history, dict) else None
    hist = list(hist.values()) if isinstance(hist, dict) else (hist if isinstance(hist, list) else [])
    return hist if len(hist) > len(flat) else flat


class Filter:
    class Valves(BaseModel):
        enabled: bool = Field(
            default=True,
            description="Prime the model whenever a .pptx/.potx template is attached.",
        )
        prime_without_template: bool = Field(
            default=True,
            description="Also prime deck/presentation requests (and follow-up edits to a deck already in the chat) when no template is attached, using a template-less primer.",
        )
        priority: int = Field(
            default=0, description="Filter execution priority (lower runs first)."
        )

    def __init__(self):
        self.valves = self.Valves()

    async def inlet(self, body: dict, __user__: dict = None, __chat_id__: str = None) -> dict:
        """Before the model answers: if a template is attached, add the outline
        spec with its layouts to the system prompt; otherwise do the same,
        without layouts, for a deck request or an edit to a deck already in
        the chat. Never raises — on any problem it returns the request
        untouched so the chat is unaffected."""
        try:
            if not self.valves.enabled:
                return body
            messages = body.get("messages", []) or []
            if any(
                isinstance(m, dict) and m.get("role") == "system" and SENTINEL in _text(m.get("content"))
                for m in messages
            ):
                return body  # already primed this request

            data = await _maybe_await(
                self._find_template_bytes(body, __chat_id__ or body.get("chat_id"))
            )
            if data is not None:
                layout_lines = _layouts_from_bytes(data)
                if not layout_lines:
                    return body
                content = _primer(layout_lines)
            elif self.valves.prime_without_template and (
                _looks_like_deck_request(messages) or _deck_in_progress(messages)
            ):
                content = _primer()
            else:
                return body

            body["messages"] = _add_primer(messages, content)
        except Exception:
            return body
        return body

    async def _find_template_bytes(self, body: dict, chat_id: str = None):
        """Newest-first, return bytes of the first attached .pptx/.potx, or None.

        Looks at the request first, then the saved chat: a template attached
        a few turns back isn't necessarily in this request's 'files' any more.
        Duplicated from the export action on purpose: Open WebUI functions are
        pasted in individually and can't import one another.
        """
        try:
            from open_webui.models.files import Files
            from open_webui.storage.provider import Storage
        except Exception:
            return None

        def file_ids(container):
            if not isinstance(container, dict):
                return
            for msg in reversed(_chat_messages(container)):
                if isinstance(msg, dict):
                    yield from _file_entries(msg.get("files"))
            yield from _file_entries(container.get("files"))

        async def saved_chat():
            if not chat_id:
                return None
            try:
                from open_webui.models.chats import Chats

                chat = await _maybe_await(Chats.get_chat_by_id(chat_id))
                return getattr(chat, "chat", None)
            except Exception:
                return None

        seen = set()
        for source in (body, "saved"):
            container = await saved_chat() if source == "saved" else source
            for fid, hint in file_ids(container):
                if fid in seen:
                    continue
                seen.add(fid)
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
