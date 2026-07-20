# openPPT

Let any LLM in [Open WebUI](https://openwebui.com) create PowerPoint decks — including models with **no function calling** (Gemma, etc.). The model just writes a Markdown outline in chat; an **"Export to PowerPoint"** button under the message converts it to a downloadable `.pptx`.

**Dump content + a template → the model fills the slides.** Attach your own `.pptx`/`.potx` to the chat, paste in raw content, and ask for a deck: the model writes the outline and the export inherits your template's theme, fonts, colors, and layouts. See [Templates](#templates-bring-your-own-theme).

## Install

1. Open WebUI → **Admin Panel → Functions → + New Function**.
2. Paste the contents of [`openppt_action.py`](openppt_action.py), save.
3. Enable the function, and toggle it on for the models you want (or globally). `python-pptx` is installed automatically from the function's `requirements` frontmatter.

An "Export to PowerPoint" button now appears under every assistant message.

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

This is the "dump content + a template" flow: give the model a pile of raw content and your branded template, let it write the outline, and the export drops that content into your design.

**Pick a layout per slide.** Name any layout from your template so the model controls the design of each slide — inline on the heading or on its own `Layout:` line:

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

Drop those names into your model's system prompt so it picks layouts that actually exist in your template.

## Recommended: a "Presentation Builder" preset

For the smoothest flow, create a model preset (**Workspace → Models → + New**) based on your model (e.g. Gemma) with this system prompt:

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
pip install python-pptx && python test_parser.py
```

v0.3 scope: title + bullet slides, speaker notes, images, **custom `.pptx`/`.potx` templates**, and per-slide layout selection by name. Placeholder mapping is generic (resolved by placeholder type), so it works on any template's layouts.
