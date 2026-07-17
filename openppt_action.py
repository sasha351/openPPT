"""
title: Export to PowerPoint
author: openPPT
version: 0.1.0
requirements: python-pptx
description: Adds an "Export to PowerPoint" button that converts a Markdown outline in the assistant message into a downloadable .pptx. Works with any model (no tool calling needed).
"""

import io
import uuid

from pptx import Presentation
from pptx.util import Pt


def parse_outline(markdown: str) -> list:
    """Parse a Markdown outline into slides: [{"title", "subtitle", "bullets": [(level, text)]}].

    '# Heading' starts a slide, '-'/'*'/'1.' lines are bullets (indent = nesting),
    '## ' on a bullet-less first slide is its subtitle. Text before the first '#'
    (chat prose) is ignored. Returns [] if no slides found.
    """
    slides = []
    slide = None
    for raw in markdown.splitlines():
        line = raw.strip()
        if not line or line.startswith("```"):
            continue
        if line.startswith("# "):
            slide = {"title": line[2:].strip(), "subtitle": "", "bullets": []}
            slides.append(slide)
            continue
        if slide is None:
            continue
        indent = len(raw) - len(raw.lstrip())
        level = min(indent // 2, 4)
        if line.startswith("## ") and len(slides) == 1 and not slide["bullets"]:
            slide["subtitle"] = line[3:].strip()
        elif line[:2] in ("- ", "* "):
            slide["bullets"].append((level, line[2:].strip()))
        elif line.split(".", 1)[0].isdigit() and "." in line:
            slide["bullets"].append((level, line.split(".", 1)[1].strip()))
        else:
            slide["bullets"].append((level, line))
    return slides


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
            continue
        ps = prs.slides.add_slide(prs.slide_layouts[1])  # Title and Content
        ps.shapes.title.text = s["title"]
        tf = ps.placeholders[1].text_frame
        for j, (level, text) in enumerate(s["bullets"]):
            p = tf.paragraphs[0] if j == 0 else tf.add_paragraph()
            p.text = text
            p.level = level
            p.font.size = Pt(20)
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
