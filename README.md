# openPPT

Let any LLM in [Open WebUI](https://openwebui.com) create PowerPoint decks — including models with **no function calling** (Gemma, etc.). The model just writes an HTML outline in chat; an **"Export to PowerPoint"** button under the message converts it to a downloadable `.pptx`.

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

> **Button shows but clicking it seems to do nothing?** It always reports back two ways: a **toast** in the browser, and a line in the Open WebUI **server log** prefixed `[openPPT]`. The usual cause is a message with no slide markers, and the toast says so. Since v0.5.2 the toast leads with the evidence rather than advice — `read 0 chars from 2 messages [user:str(12), assistant:str(0)]` — so a report from a machine you can't reach still says which paste ran and how much text the action was handed. `0 chars` means the outline never reached the parser: some Open WebUI versions POST the action an assistant message whose `content` is empty, which v0.6.0 handles by reading the outline from the saved chat instead. If you're in a **Temporary Chat**, that fallback can't help — Temporary Chats are never written to Open WebUI's chat storage, so there is nothing to fall back to; v0.6.1 detects this (a `chat_id` present but no stored chat found) and says so directly in the toast — turn off Temporary Chat and retry. On some deployments the button can still fail to read the chat reliably even in a normal saved chat (an Open WebUI version/config quirk in how it POSTs the action request, not the outline format); if the toast keeps reporting `0 chars` with no clear cause, skip the button and use the [command-line converter](#no-working-export-button-convert-from-the-command-line) instead. Note that a model preset can suppress the status line under a message via its `status_updates` capability — that hides statuses only, never the toast or the log. (Before v0.3.1, openPPT reported *only* through statuses, so a suppressed status looked exactly like a dead button.) If openPPT reports no outline, it posts a short diagnostic under the message saying what it actually read; that diagnostic is deliberately not slide-shaped, so clicking export again re-reads your outline rather than exporting the error report (which is what v0.3.5 did).

## Updating an existing install

**Edit the function you already have — don't add a second one**, or you'll get two buttons. **Admin Panel → Functions** → pencil/**Edit** on *Export to PowerPoint* → select all, paste the current [`openppt_action.py`](openppt_action.py) → **Save**. Open WebUI reloads the module and refreshes its cache on save, so it's live immediately: no container restart, no re-enabling, no re-toggling per model, no settings to migrate. Confirm the version reads **0.7.1** afterwards.

## Outline format

The button parses the assistant message it's clicked on. Each slide is a `<slide>` element:

```html
<slide title="Q3 Results" subtitle="Revenue up 12% YoY"></slide>
<!-- no bullets, title + subtitle attributes = title slide -->

<slide title="Key Wins">           <!-- each <slide> starts a slide -->
<ul>
<li>Launched EU region</li>        <!-- <li> bullets -->
<li>Churn down to 2.1%
<ul><li>Best quarter ever</li></ul>  <!-- nested <ul> = nested bullet -->
</li>
</ul>
</slide>
```

Prose before the first `<slide>` (e.g. "Here's your deck:") is ignored.

**The parser is deliberately tolerant**, using Python's standard `html.parser` rather than a strict validator, because local models don't emit clean markup any more reliably than they emitted clean Markdown:

- An unclosed `<slide>`, `<li>`, or `<p>` is auto-closed by whatever ends it (the next `<slide>`, or the end of the message) instead of losing the rest of the deck.
- Bare text sitting directly inside `<slide>` — not wrapped in `<p>` or `<li>` — still becomes a bullet instead of being silently dropped.
- Inline tags like `<b>`, `<code>`, and `<a href="...">` are flattened to plain text.
- A ` ``` ` fence around the whole reply is unwrapped rather than swallowing the deck — models wrap their answer in one whether or not you ask them to.

Clicking **Export to PowerPoint** appends a download link to the message, right after an `<!-- openppt -->` comment. Re-clicking Export on that same message stops parsing dead at that comment — even if the model's own last `<slide>` was left unclosed — so your own download link never becomes a bogus trailing bullet.

**Speaker notes:** a `<notes>` (or `<blockquote>`) element anywhere in a slide becomes that slide's speaker notes.

```html
<slide title="Key Wins">
<li>Launched EU region</li>
<notes>mention the EU launch timeline if asked</notes>
</slide>
```

**Tables:** a `<table>` of `<tr>` rows of `<td>`/`<th>` cells becomes a real PowerPoint table on that slide. Bullets on the same slide keep the top of the content area and the table sits below them.

```html
<slide title="Q3 Numbers">
<p>Revenue held through the reorg</p>
<table>
<tr><th>Region</th><th>Revenue</th><th>Growth</th></tr>
<tr><td>EMEA</td><td>4.2</td><td>12%</td></tr>
<tr><td>APAC</td><td>3.1</td><td>30%</td></tr>
</table>
</slide>
```

**Code blocks:** `<pre><code>...</code></pre>` inside a slide is rendered in a monospace box. Escape literal `<`, `>`, and `&` in the code as `&lt;` `&gt;` `&amp;` so they aren't mistaken for markup — or use a ` ``` ` fence instead, which openPPT HTML-escapes for you automatically before parsing.

````html
<slide title="Deploy">
<pre><code>make ship</code></pre>
</slide>
````

**Images:** `<img src="url-or-path" alt="description">` — bare inside a slide, or inside an `<li>` — embeds that image instead of a text bullet. A slide made up entirely of image bullets keeps its title and drops the bullet list in favor of the image(s); images that fail to load (bad URL, network error) are silently skipped so the export never fails because of one broken link.

```html
<slide title="Team Photo">
<img src="https://example.com/team.jpg" alt="Team offsite">
</slide>
```

## Templates: bring your own theme

Attach a `.pptx` or `.potx` file to the chat and click **Export to PowerPoint**. Instead of the plain default theme, the deck is built **on your template** — inheriting its slide master, layouts, fonts, and colors. If you attach several, the most recent one wins. Any sample slides in the template are dropped; only its design is kept.

With the [Filter](#the-workflow-template-in-deck-out) installed, attaching the template also primes the model automatically: it's told your template's layout names and the outline format, then asked to organize whatever content you provided into a deck that flows. You don't have to describe the format or the layouts yourself — just drop in the template and your content.

**Pick a layout per slide.** Name any layout from your template with the `layout="..."` attribute so the model controls the design of each slide (the Filter tells the model to do this for you):

```html
<slide title="Product Roadmap" layout="Section Header"></slide>

<slide title="Q4 Priorities" layout="Two Content">
<li>Ship billing v2</li>
<li>Migrate to EU region</li>
</slide>
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
respond ONLY with an HTML outline in exactly this format:

<slide title="Deck Title" subtitle="One-line subtitle"></slide>

<slide title="First Slide Title">
<ul>
<li>Short bullet point</li>
<li>Another bullet point
<ul><li>Sub-point</li></ul>
</li>
</ul>
</slide>

Rules: every slide is a <slide title="..."> element, use 3-6 <li> bullets
per slide, keep bullets under 12 words, no prose outside the outline. Make
every bullet concrete: one specific number, name, date, or action verb —
never a topic label. Distill the source material by theme rather than
transcribing it in order, and size the deck to what the content actually
supports. When the user requests changes, reply with the full revised
outline.

If the user attaches a template and lists its layout names, choose one per
slide with a layout="<name>" attribute (e.g.
'<slide title="Overview" layout="Section Header">'). Use only names the
user provided.
```

Chat until the outline looks right, then click **Export to PowerPoint**. Attach a `.pptx`/`.potx` template to the chat to give the deck your theme.

## No working Export button? Convert from the command line

[`outline_to_pptx.py`](outline_to_pptx.py) builds the same `.pptx` as the **Export to PowerPoint** button, but reads the outline from a file instead of the chat — so it sidesteps Open WebUI entirely and works even where the button can't reliably read the chat (locked-down deployments, Temporary Chats, or any of the other causes above). It reuses the same `parse_outline`/`build_pptx` code as the button, so it follows the [outline format](#outline-format) exactly.

Requires only `python-pptx` (`pip install python-pptx`), and `openppt_action.py` in the same directory (or on your `PYTHONPATH`).

Copy the model's outline out of the chat into a text file, then:

```bash
python outline_to_pptx.py outline.html                     # writes outline.pptx
python outline_to_pptx.py outline.html -o deck.pptx         # choose the output name
python outline_to_pptx.py outline.html -t brand.pptx        # build on a template, same as attaching one to the chat
pbpaste | python outline_to_pptx.py -                       # read the outline from the clipboard/stdin instead of a file
```

If the outline doesn't parse (no `<slide title="...">` elements found), it exits with an error instead of writing an empty deck.

## Development

```
pip install python-pptx pydantic && python test_parser.py && python test_filter.py
```

v0.7.1 fixes the no-outline diagnostic: it now reports whether the request carried a `chat_id` at all, and the "looks like a Temporary Chat" hint fires whenever no saved copy of the chat was found — including when the request has no `chat_id`, which is exactly the unsaved-chat case the old condition suppressed.

v0.7.0 replaces the Markdown outline grammar with an HTML one: `<slide title="..." subtitle="..." layout="...">` per slide, `<ul>`/`<li>` bullets (nesting = a nested `<ul>`), `<notes>`, `<table>`/`<tr>`/`<td>`, `<pre><code>`, and `<img>`. It's parsed with Python's standard `html.parser` instead of regex heading-level guessing, so slide boundaries are unambiguous and malformed/unclosed tags degrade gracefully instead of losing content. The template primer's content-quality guidance is also expanded — beyond "make bullets concrete," it now asks the model to distill by theme, cap the deck to what the content supports, and push supporting detail into notes.

v0.3 scope: title + bullet slides, speaker notes, images, **custom `.pptx`/`.potx` templates**, per-slide layout selection by name, and a **template primer Filter** that turns "attach a template + dump content" into a formatted deck automatically. Placeholder mapping is generic (resolved by placeholder type), so it works on any template's layouts.

v0.3.1 adds a tolerant parser (most-frequent heading level, heading-less fallbacks, inline-markdown stripping, `> ` notes) and makes every outcome report via a toast and the server log instead of the suppressible status line.

v0.5.0 renders **tables** and **code blocks**: `|`-delimited rows become a real PowerPoint table, a ``` fence becomes a monospace box, and a fence wrapping the whole outline (or a `# comment` inside one) no longer confuses the slide parser.

v0.4.0 fixes exports that came out garbled: openPPT now ignores text it appended to a message itself (its own download link or diagnostic) instead of parsing it back into slides, falls back to the newest message that actually holds an outline when the clicked one arrives empty, and shrinks long titles and dense bullet lists to fit their placeholders instead of letting them overflow across the slide.

v0.6.1 names Temporary Chat as a likely cause when the stored-chat fallback finds nothing (a `chat_id` present but no chat in storage), in both the toast and the appended diagnostic. Also adds [`outline_to_pptx.py`](outline_to_pptx.py), a standalone command-line converter for deployments where the button can't read the chat reliably at all.
