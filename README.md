# openPPT

Let any LLM in [Open WebUI](https://openwebui.com) create PowerPoint decks — including models with **no function calling** (Gemma, etc.). The model just writes a Markdown outline in chat; an **"Export to PowerPoint"** button under the message converts it to a downloadable `.pptx`.

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
```

Chat until the outline looks right, then click **Export to PowerPoint**.

## Development

```
pip install python-pptx && python test_parser.py
```

v0.2 scope: title + bullet slides on the default template, plus speaker notes and images. No custom themes (yet).
