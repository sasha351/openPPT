# openPPT

Let any LLM in [Open WebUI](https://openwebui.com) create PowerPoint decks — including models with **no function calling** (Gemma, etc.). The model just writes a Markdown outline in chat; an **"Export to PowerPoint"** button under the message converts it to a downloadable `.pptx`.

## The workflow: template in, deck out

**Attach a template + dump your content → get a formatted deck.**

1. Attach your own `.pptx`/`.potx` template to the chat.
2. Paste in the raw content you want the deck to be about (notes, a doc, bullet dumps — anything).
3. Ask for a presentation.

The model reads your template's layouts, organizes your content into slides that flow logically (title → overview → grouped sections → summary), and writes the outline. Click **Export to PowerPoint** and the deck is built on your template — inheriting its theme, fonts, colors, and layouts.

Two pieces make this work, and you install both:

| Piece | File | What it does |
|-------|------|--------------|
| **Filter** | [`openppt_filter.py`](openppt_filter.py) | Runs *before* the model. Detects your attached template, reads its layouts, and primes the model to format your content into a flowing deck that fills those layouts. |
| **Action** | [`openppt_action.py`](openppt_action.py) | The **Export to PowerPoint** button. Turns the outline into a `.pptx` built on your template. |

The Action works on its own (write an outline, click export). Add the Filter and you get the hands-off "template in, deck out" flow above.

## Install

1. Open WebUI → **Admin Panel → Functions → + New Function**.
2. Paste the contents of [`openppt_action.py`](openppt_action.py), save. Repeat for [`openppt_filter.py`](openppt_filter.py).
3. Enable both, and toggle them on for the models you want (or globally). `python-pptx` is installed automatically from each function's `requirements` frontmatter.

An "Export to PowerPoint" button now appears under every assistant message, and attaching a template primes the model automatically.

> **No button showing?** Open WebUI only renders Action buttons for models **saved in its database** — models pulled live from Ollama, or reached through a **direct OpenAI-API connection**, don't get them, even with the function enabled globally and no load errors. Fix: **Workspace → Models → ➕**, set the **Base Model** to the model you chat with, save, then use that custom model. The button (and the Template Primer filter) attach to the custom model. This is the most common reason the button never appears.

## Outline format

The button parses the assistant message it's clicked on:

```markdown
# Q3 Results            ← first slide with no bullets = title slide
## Revenue up 12% YoY   ← its subtitle

# Key Wins              ← '#' starts a slide
- Launched EU region    ← '-', '*', or '1.' bullets
  - Best quarter ever   ← indent = nesting
```

Prose before the first `#` (e.g. "Here's your deck:") is ignored.

**Speaker notes:** a line starting with `Notes:` (anywhere in a slide, after its heading) becomes that slide's speaker notes.

```markdown
# Key Wins
- Launched EU region
Notes: mention the EU launch timeline if asked
```

**Images:** a bullet written as `![alt text](url or path)` embeds that image on the slide instead of a text bullet. A slide made up entirely of image bullets keeps its title and drops the bullet list in favor of the image(s); images that fail to load (bad URL, network error) are silently skipped so the export never fails because of one broken link.

```markdown
# Team Photo
- ![Team offsite](https://example.com/team.jpg)
```

## Templates: bring your own theme

Attach a `.pptx` or `.potx` file to the chat and click **Export to PowerPoint**. Instead of the plain default theme, the deck is built **on your template** — inheriting its slide master, layouts, fonts, and colors. If you attach several, the most recent one wins. Any sample slides in the template are dropped; only its design is kept.

With the [Filter](#the-workflow-template-in-deck-out) installed, attaching the template also primes the model automatically: it's told your template's layout names and the outline format, then asked to organize whatever content you provided into a deck that flows. You don't have to describe the format or the layouts yourself — just drop in the template and your content.

**Pick a layout per slide.** Name any layout from your template so the model controls the design of each slide — inline on the heading or on its own `Layout:` line (the Filter tells the model to do this for you):

```markdown
# Product Roadmap @layout: Section Header

# Q4 Priorities
Layout: Two Content
- Ship billing v2
- Migrate to EU region
```

Names are matched case-insensitively (exact first, then a loose substring match), so `@layout: two content` finds **Two Content**. If no template is attached, or a name doesn't match, openPPT falls back to sensible defaults (a title-slide layout for the opening slide, a title-and-content layout for the rest) chosen by inspecting each layout's placeholders — so this works with any template, not just PowerPoint's built-in themes.

Not sure what your template offers? `list_layouts()` returns every layout name (and its placeholder counts):

```python
from openppt_action import list_layouts
for name, ph in list_layouts(open("brand.potx", "rb").read()):
    print(name, ph)   # e.g. 'Section Header' {'title': 1, 'subtitle': 0, 'body': 1}
```

Drop those names into your model's system prompt so it picks layouts that actually exist in your template. (The Filter does exactly this on your behalf whenever a template is attached, so you rarely need to.)

## No template? A "Presentation Builder" preset

The Filter only kicks in when a template is attached. For deck-building **without** a template, create a model preset (**Workspace → Models → + New**) based on your model (e.g. Gemma) with this system prompt:

```
You are a presentation builder. When the user asks for a presentation,
respond ONLY with a Markdown outline in exactly this format:

# Deck Title
## One-line subtitle

# First Slide Title
- Short bullet point
- Another bullet point
  - Sub-point (indent two spaces)

Rules: start every slide with '# ', use 3-6 bullets per slide, keep
bullets under 12 words, no prose outside the outline. When the user
requests changes, reply with the full revised outline.

If the user attaches a template and lists its layout names, choose one per
slide by adding '@layout: <name>' to the heading (e.g.
'# Overview @layout: Section Header'). Use only names the user provided.
```

Chat until the outline looks right, then click **Export to PowerPoint**. Attach a `.pptx`/`.potx` template to the chat to give the deck your theme.

## Development

```
pip install python-pptx pydantic && python test_parser.py && python test_filter.py
```

v0.3 scope: title + bullet slides, speaker notes, images, **custom `.pptx`/`.potx` templates**, per-slide layout selection by name, and a **template primer Filter** that turns "attach a template + dump content" into a formatted deck automatically. Placeholder mapping is generic (resolved by placeholder type), so it works on any template's layouts.
