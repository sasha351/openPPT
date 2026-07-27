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
| **Filter** | [`openppt_filter.py`](openppt_filter.py) | Runs *before* the model. Detects your attached template, reads its layouts, and primes the model to format your content into a flowing deck that fills those layouts. No template? It still primes plain deck/presentation requests with the outline format. |
| **Action** | [`openppt_action.py`](openppt_action.py) | The **Export to PowerPoint** button. Turns the outline into a `.pptx` built on your template. |

The Action works on its own (write an outline, click export). Add the Filter and you get the hands-off "template in, deck out" flow above.

## Install

1. Open WebUI → **Admin Panel → Functions → + New Function**.
2. Paste the contents of [`openppt_action.py`](openppt_action.py), save. Repeat for [`openppt_filter.py`](openppt_filter.py).
3. Enable both, and toggle them on for the models you want (or globally). `python-pptx` is installed automatically from each function's `requirements` frontmatter.

An "Export to PowerPoint" button now appears under every assistant message, and attaching a template primes the model automatically.

> **No button showing?** Open WebUI only renders Action buttons for models **saved in its database** — models pulled live from Ollama, or reached through a **direct OpenAI-API connection**, don't get them, even with the function enabled globally and no load errors. Fix: **Workspace → Models → ➕**, set the **Base Model** to the model you chat with, save, then use that custom model. The button (and the Template Primer filter) attach to the custom model. This is the most common reason the button never appears.

> **Where does the deck show up?** Three places, so a version that ignores one still delivers: as an **attachment chip** under the message (click it to preview or download), as a **📊 Download link** appended to the message text, and as that same full URL repeated as plain text right after it — copy-pasteable even where neither the chip nor the markdown link renders. The same URL also goes to the success toast and the server log. It's a full `scheme://host/api/v1/files/<id>/content` URL when Open WebUI passes the request through to the Action; otherwise it falls back to the relative path, which you'd paste onto your Open WebUI host yourself.

> **Button shows but clicking it seems to do nothing?** It always reports back two ways: a **toast** in the browser, and a line in the Open WebUI **server log** prefixed `[openPPT]`. The usual cause is a message with no slide markers, and the toast says so. Note that a model preset can suppress the status line under a message via its `status_updates` capability — that hides statuses only, never the toast or the log. (Before v0.3.1, openPPT reported *only* through statuses, so a suppressed status looked exactly like a dead button.) If openPPT reports no outline, it posts a short diagnostic under the message saying what it actually read; that diagnostic is deliberately not slide-shaped, so clicking export again re-reads your outline rather than exporting the error report (which is what v0.3.5 did).

## Updating an existing install

**Edit the function you already have — don't add a second one**, or you'll get two buttons. **Admin Panel → Functions** → pencil/**Edit** on *Export to PowerPoint* → select all, paste the current [`openppt_action.py`](openppt_action.py) → **Save**. Open WebUI reloads the module and refreshes its cache on save, so it's live immediately: no container restart, no re-enabling, no re-toggling per model, no settings to migrate. Confirm the version reads **0.5.3** afterwards.

## Outline format

The button parses the assistant message it's clicked on:

```markdown
# Q3 Results            ← first slide with no bullets = title slide
## Revenue up 12% YoY   ← its subtitle

# Key Wins              ← '#' starts a slide
- Launched EU region    ← '-', '*', '+', '1.', '1)' bullets
  - Best quarter ever   ← indent = nesting
```

Prose before the first `#` (e.g. "Here's your deck:") is ignored.

**The parser is tolerant**, because local models rarely stick to one format:

- Slides start at whichever heading level is **most frequent**. A deck written with `## ` per slide works, and a lone `# ` above it becomes the deck title rather than a stray bullet.
- With **no headings at all**, slides start at `Slide 3: Title` prefixes, lone `**Bold Title**` lines, or `---` separators.
- Inline `**bold**`, `` `code` `` and `[links](url)` are flattened to plain text.

Known gap: a deck using **only** `---` separators loses its first section, since treating leading prose as a slide would turn every chat message into a deck. Give the first slide a heading or a `Slide 1:` prefix.

Clicking **Export to PowerPoint** appends a download link to the message. If you click it again on that same message, the parser stops at that appended link and ignores everything after it — so re-exporting never turns your own download link into a bogus trailing bullet.

**Speaker notes:** a line starting with `Notes:` — or a `> ` blockquote — anywhere in a slide after its heading becomes that slide's speaker notes.

```markdown
# Key Wins
- Launched EU region
Notes: mention the EU launch timeline if asked
> or write the note as a blockquote
```

**Tables:** consecutive `|`-delimited lines become a real PowerPoint table on that slide (the `|---|` alignment row is dropped). Bullets on the same slide keep the top of the content area and the table sits below them.

```markdown
# Q3 Numbers
- Revenue held through the reorg

| Region | Revenue | Growth |
|--------|---------|--------|
| EMEA   | 4.2     | 12%    |
| APAC   | 3.1     | 30%    |
```

**Code blocks:** a ` ``` ` fence inside a slide is rendered in a monospace box, so `# comments` in the code aren't mistaken for slide headings. A fence that opens *before* the first slide is treated as a wrapper around the whole outline and parsed through — models wrap their answer in one whether or not you ask them to.

````markdown
# Deploy
```bash
make ship
```
````

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

## No template? The Filter primes deck requests anyway

With the Filter installed, attaching a template isn't required: if the last message looks like a deck/presentation request (mentions "deck", "presentation", "slides", etc.), it injects the same outline format and content-quality bar without any layout list. Turn this off with the Filter's `prime_without_template` valve if you'd rather it only fire when a template is attached.

If you'd rather not rely on the Filter (or it's disabled), create a model preset (**Workspace → Models → + New**) based on your model (e.g. Gemma) with this system prompt instead:

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
bullets under 12 words, no prose outside the outline. Make every bullet
concrete: one specific number, name, date, or action verb — never a topic
label. When the user requests changes, reply with the full revised outline.

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

v0.3.1 adds a tolerant parser (most-frequent heading level, heading-less fallbacks, inline-markdown stripping, `> ` notes) and makes every outcome report via a toast and the server log instead of the suppressible status line.

v0.5.0 renders **tables** and **code blocks**: `|`-delimited rows become a real PowerPoint table, a ``` fence becomes a monospace box, and a fence wrapping the whole outline (or a `# comment` inside one) no longer confuses the slide parser.

v0.4.0 fixes exports that came out garbled: openPPT now ignores text it appended to a message itself (its own download link or diagnostic) instead of parsing it back into slides, falls back to the newest message that actually holds an outline when the clicked one arrives empty, and shrinks long titles and dense bullet lists to fit their placeholders instead of letting them overflow across the slide.
