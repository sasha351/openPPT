"""
title: Export to PowerPoint
author: openPPT
version: 0.2.0
requirements: python-pptx
description: Adds an "Export to PowerPoint" button that converts a Markdown outline in the assistant message into a downloadable .pptx, with support for speaker notes and images. Works with any model (no tool calling needed).
"""

import io
import re
import urllib.request
import uuid

from pptx import Presentation
from pptx.util import Pt

IMAGE_RE = re.compile(r"^!\[(.*?)\]\((.*?)\)$")
NOTES_RE = re.compile(r"^notes?:\s*(.*)$", re.IGNORECASE)


def parse_outline(markdown: str) -> list:
    """Parse a Markdown outline into slides.

    Each slide is {"title", "subtitle", "notes", "bullets": [(level, text) | (level, {"image": url, "alt": alt})]}.

    '# Heading' starts a slide, '-'/'*'/'1.' lines are bullets (indent = nesting),
    '## ' on a bullet-less first slide is its subtitle. A bullet written as
    '![alt](url)' embeds an image instead of text. A line starting with
    'Notes:' (anywhere after the slide heading) is appended to that slide's
    speaker notes. Text before the first '#' (chat prose) is ignored.
    Returns [] if no slides found.
    """
    slides = []
    slide = None
    for raw in markdown.splitlines():
        line = raw.strip()
        if not line or line.startswith("```"):
            continue
        if line.startswith("# "):
            slide = {"title": line[2:].strip(), "subtitle": "", "notes": "", "bullets": []}
            slides.append(slide)
            continue
        if slide is None:
            continue
        notes_match = NOTES_RE.match(line)
        if notes_match:
            slide["notes"] = (slide["notes"] + "\n" + notes_match.group(1)).strip()
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


def build_pptx(slides: list) -> bytes:
    """Build a .pptx from parsed slides using the python-pptx default template."""
    prs = Presentation()
    for i, s in enumerate(slides):
        if i == 0 and not s["bullets"]:
            layout = prs.slide_layouts[0]  # Title slide
            ps = prs.slides.add_slide(layout)
            ps.shapes.title.text = s["title"]
            if s["subtitle"]:
                ps.placeholders[1].text = s["subtitle"]
        else:
            text_bullets = [(level, text) for level, text in s["bullets"] if isinstance(text, str)]
            image_bullets = [(level, text) for level, text in s["bullets"] if isinstance(text, dict)]

            image_only = not text_bullets and image_bullets
            layout = prs.slide_layouts[5] if image_only else prs.slide_layouts[1]
            ps = prs.slides.add_slide(layout)
            ps.shapes.title.text = s["title"]
            if not image_only:
                tf = ps.placeholders[1].text_frame
                for j, (level, text) in enumerate(text_bullets):
                    p = tf.paragraphs[0] if j == 0 else tf.add_paragraph()
                    p.text = text
                    p.level = level
                    p.font.size = Pt(20)

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

            await status(f"Building PowerPoint ({len(slides)} slides)…")
            data = build_pptx(slides)

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
