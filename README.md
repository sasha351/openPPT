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

**Action Valves** (the gear icon next to *Export to PowerPoint* in Admin Panel → Functions):

| Valve | Default | What it does |
|-------|---------|--------------|
| `default_template` | *(empty)* | The template every user's decks are built on: the **filename** of a `.pptx`/`.potx` an admin uploaded to Open WebUI (e.g. `brand.potx` — see [Default template for everyone](#default-template-for-everyone)). Also accepts an absolute server path or an Open WebUI file id. Empty = the built-in theme. A template attached to a chat wins for that chat. |
| `accent_color` | `1F4E79` | Hex color of the built-in theme: title and section pages, titles, bullets, tables. Ignored when a template is used. |
| `slide_numbers` | on | Number every slide except title and section pages, in the layout's own slide-number position. |
| `auto_download` | on | Download the `.pptx` in the browser as soon as it's built. |

> **No button showing?** Open WebUI only renders Action buttons for models **saved in its database** — models pulled live from Ollama, or reached through a **direct OpenAI-API connection**, don't get them, even with the function enabled globally and no load errors. Fix: **Workspace → Models → ➕**, set the **Base Model** to the model you chat with, save, then use that custom model. The button (and the Template Primer filter) attach to the custom model. This is the most common reason the button never appears.

> **Where does the deck show up?** It **downloads in your browser** the moment it's built (the `auto_download` Valve), and it also stays under the message as an **attachment chip** and a **📊 Download link**, with the full URL repeated as plain text — so you can fetch it again later. Every Open WebUI version drops at least one of these: v0.10+ renders a reply from its structured `output` rather than its `content`, which is why links appended the old way (a `message` event) never appeared — the "it says the download is ready but nothing shows up" bug fixed in v0.8.0. The link is now written by returning the edited message from the action, which Open WebUI syncs into `output` (v0.12+); v0.9 and older render the link but not file chips. Re-exporting replaces the previous link rather than stacking another. The same URL also goes to the success toast and the server log.

> **Button shows but clicking it seems to do nothing?** It always reports back two ways: a **toast** in the browser, and a line in the Open WebUI **server log** prefixed `[openPPT]`. The usual cause is a message with no slide markers, and the toast says so. Since v0.5.2 the toast leads with the evidence rather than advice — `read 0 chars from 2 messages [user:str(12), assistant:str(0)]` — so a report from a machine you can't reach still says which paste ran and how much text the action was handed. `0 chars` means the outline never reached the parser: some Open WebUI versions POST the action an assistant message whose `content` is empty, which v0.6.0 handles by reading the outline from the saved chat instead. If you're in a **Temporary Chat**, that fallback can't help — Temporary Chats are never written to Open WebUI's chat storage, so there is nothing to fall back to; v0.6.1 detects this (a `chat_id` present but no stored chat found) and says so directly in the toast — turn off Temporary Chat and retry. On some deployments the button can still fail to read the chat reliably even in a normal saved chat (an Open WebUI version/config quirk in how it POSTs the action request, not the outline format); if the toast keeps reporting `0 chars` with no clear cause, skip the button and use the [command-line converter](#no-working-export-button-convert-from-the-command-line) instead. Note that a model preset can suppress the status line under a message via its `status_updates` capability — that hides statuses only, never the toast or the log. (Before v0.3.1, openPPT reported *only* through statuses, so a suppressed status looked exactly like a dead button.) If openPPT reports no outline, it posts a short diagnostic under the message saying what it actually read; that diagnostic is deliberately not slide-shaped, so clicking export again re-reads your outline rather than exporting the error report (which is what v0.3.5 did).

## Updating an existing install

**Edit the function you already have — don't add a second one**, or you'll get two buttons. **Admin Panel → Functions** → pencil/**Edit** on *Export to PowerPoint* → select all, paste the current [`openppt_action.py`](openppt_action.py) → **Save**. Open WebUI reloads the module and refreshes its cache on save, so it's live immediately: no container restart, no re-enabling, no re-toggling per model, no settings to migrate. Confirm the version reads **0.8.1** afterwards (and **0.6.0** for the Template Primer filter).

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

Prose before the first `<slide>` (e.g. "Here's your deck:") is ignored. A later `<slide>` with no body becomes a **section divider**, on the template's section layout.

**Columns:** for a side-by-side comparison, wrap each side in `<column heading="...">` (or `<col>`). Two columns use a two-column layout ("Two Content", or any layout with two body placeholders); with no such layout they become side-by-side text boxes. Bullets outside the columns sit above them.

```html
<slide title="Build vs buy">
<p>Decision needed by <b>Sep 1</b></p>
<column heading="Build"><ul><li>Full control</li><li>~6 months</li></ul></column>
<column heading="Buy"><ul><li><b>Live in 2 weeks</b></li><li>$40k / year</li></ul></column>
</slide>
```

**Text formatting:** `<b>`/`<strong>`, `<i>`/`<em>`, `<u>`, and `<code>` inside a bullet are kept as bold, italic, underline, and monospace.

**The parser is deliberately tolerant**, using Python's standard `html.parser` rather than a strict validator, because local models don't emit clean markup any more reliably than they emitted clean Markdown:

- An unclosed `<slide>`, `<li>`, or `<p>` is auto-closed by whatever ends it (the next `<slide>`, or the end of the message) instead of losing the rest of the deck.
- Bare text sitting directly inside `<slide>` — not wrapped in `<p>` or `<li>` — still becomes a bullet instead of being silently dropped.
- Bare lines of text inside a `<slide>` become one bullet each; a line break inside an `<li>` is just a space, as in HTML.
- Other inline tags, like `<a href="...">`, are flattened to plain text.
- A reasoning model's thinking is ignored: `<think>…</think>` (and the other reasoning wrappers Open WebUI recognises), Open WebUI's own `<details type="reasoning">` blocks, and everything before a lone `</think>` (what Nemotron-style templates emit) are dropped before parsing, so slides the model drafted while thinking aren't exported.
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

**Images:** `<img src="..." alt="description">` — bare inside a slide, or inside an `<li>` — embeds that image instead of a text bullet. Images keep their aspect ratio: beside the bullets when a slide has both, filling the content area when a slide has only images, in a grid when there are several. Accepted sources: Open WebUI file links (`/api/v1/files/<id>/content`, e.g. its generated images — only your own files), `data:image/...;base64` URIs, and public `http(s)` URLs (never private, loopback, or link-local addresses; 10 MB max). Local file paths work only in the [command-line converter](#no-working-export-button-convert-from-the-command-line). An image that fails to load or isn't really an image is skipped, so one broken link never fails the export.

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

Names are matched case-insensitively (exact first, then a loose substring match), so `layout="two content"` finds **Two Content**. If no template is attached, or a name doesn't match, openPPT falls back to sensible defaults (a title-slide layout for the opening slide, a section layout for dividers, a two-column layout for `<column>` slides, a title-and-content layout for the rest) chosen by inspecting each layout's placeholders — so this works with any template, not just PowerPoint's built-in themes.

**Which look a deck gets:** a template attached to the chat, else the admin's [default template](#default-template-for-everyone), else openPPT's built-in theme — 16:9, in the `accent_color` Valve's color, with accent title and section pages, dark left-aligned titles, and slide numbers. Text is sized to the box it lands in (large on sparse slides, smaller on dense ones) instead of overflowing, and placeholders left empty are removed so PowerPoint never shows "Click to add text".

Not sure what your template offers? `list_layouts()` returns every layout name (and its placeholder counts):

```python
from openppt_action import list_layouts
for name, ph in list_layouts(open("brand.potx", "rb").read()):
    print(name, ph)   # e.g. 'Section Header' {'title': 1, 'subtitle': 0, 'body': 1}
```

Drop those names into your model's system prompt so it picks layouts that actually exist in your template. (The Filter does exactly this on your behalf whenever a template is attached — including one attached a few turns earlier — so you rarely need to.)

## Default template for everyone

To have every user's decks built on your company template, without anyone attaching it:

1. **Upload the template as an admin.** Attach `brand.potx` (or `.pptx`) to any chat once, or add it to a Knowledge base. Open WebUI keeps the file. Only uploads by **admin** accounts count, so a user can't replace everyone's template by uploading a file with the same name.
2. **Admin Panel → Functions →** gear icon on **Export to PowerPoint** → set `default_template` to the filename, `brand.potx`. Save.

That's it. Both parts use the setting:

- **Export:** any deck exported without an attached template is built on it.
- **The Template Primer filter** reads the same setting from the Export action. So when any user asks for a deck, or edits one already in the chat, the model is told the template's layout names and picks one per slide (`layout="..."`), the same as when a template is attached. Other messages aren't touched.

**To update the template**, upload the new version with the same filename. The newest admin upload wins. A template attached to a specific chat still overrides the default for that chat.

`.potx` files work as well as `.pptx`. Before v0.8.1, a real `.potx` was rejected by python-pptx ("not a PowerPoint file").

**If it isn't picked up:**
- Exports still show the built-in theme, and the server log (`docker logs open-webui 2>&1 | grep openPPT`) says `default_template 'brand.potx': no such path, file id, or admin-uploaded file`. Check that the upload was made by an admin account, and that the filename matches exactly (case doesn't matter).
- Admin Panel → Settings → Documents has an **Allowed File Extensions** list: if it's set, add `pptx` / `potx` to it, or the upload is refused.
- Open WebUI may report that it couldn't extract text from the template. That doesn't matter; the file is still stored.

## No template? The Filter primes deck requests anyway

With the Filter installed, attaching a template isn't required: if the last message looks like a deck/presentation request (mentions "deck", "presentation", "slides", etc.) — or the chat already holds an outline, so this turn is probably an edit like "make slide 3 shorter" — it adds the same outline format and content-quality bar without any layout list. Turn this off with the Filter's `prime_without_template` valve if you'd rather it only fire when a template is attached. The primer is appended to the existing system prompt rather than sent as a second system message, because many chat templates (Gemma's and Nemotron's among them) only accept one system message, at the start.

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
per slide, keep bullets under 12 words, no prose outside the outline. For a
side-by-side comparison, put each side in <column heading="...">. If you
think before answering, the full outline still goes in your final answer. Make
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

## Recommended setup: Open WebUI 0.12.3 with Gemma 4 31B or Nemotron 3 Ultra

openPPT v0.8.0 is checked against Open WebUI 0.12.3's source: the returned-message link, the file chip, the browser download, and the async Files API the Filter needs all apply there. Settings worth making on each **Workspace → Models** entry you export from (Advanced Params):

- **Max tokens: 8192 or more**, especially for **Nemotron 3 Ultra** with reasoning on — its thinking spends the same budget, and a reply cut off mid-deck exports without its last slides. openPPT warns when the last slide was left open, which is what a cut-off reply looks like.
- **Context length: 16k or more** if the model is served through Ollama (its default is far smaller): the primer is about 1,200 tokens, on top of whatever documents you paste or attach.
- **Temperature around 0.3–0.5** keeps the HTML format steady. Higher settings mostly add stray prose.
- **Reasoning:** keep it on for Nemotron 3 Ultra if you can afford the tokens — the decks are better organised. openPPT ignores the thinking either way, including drafts of the deck the model writes while thinking.
- If either model reaches Open WebUI through a **direct OpenAI-compatible connection** (vLLM, NIM, llama.cpp server), remember the custom-model step under [Install](#install), or the button won't show.

## No working Export button? Convert from the command line

[`outline_to_pptx.py`](outline_to_pptx.py) builds the same `.pptx` as the **Export to PowerPoint** button, but reads the outline from a file instead of the chat — so it sidesteps Open WebUI entirely and works even where the button can't reliably read the chat (locked-down deployments, Temporary Chats, or any of the other causes above). It reuses the same `parse_outline`/`build_pptx` code as the button, so it follows the [outline format](#outline-format) exactly.

Requires only `python-pptx` (`pip install python-pptx`), and `openppt_action.py` in the same directory (or on your `PYTHONPATH`).

Copy the model's outline out of the chat into a text file, then:

```bash
python outline_to_pptx.py outline.html                     # writes outline.pptx
python outline_to_pptx.py outline.html -o deck.pptx         # choose the output name
python outline_to_pptx.py outline.html -t brand.pptx        # build on a template, same as attaching one to the chat
python outline_to_pptx.py outline.html --accent 0F766E      # built-in theme in another color
pbpaste | python outline_to_pptx.py -                       # read the outline from the clipboard/stdin instead of a file
```

If the outline doesn't parse (no `<slide title="...">` elements found), it exits with an error instead of writing an empty deck. Unlike the button, the converter lets `<img src>` name a local file — it's your own machine.

## Development

```
pip install python-pptx pydantic && python test_parser.py && python test_filter.py
```

v0.8.1 makes a default template work for every user:

- `default_template` can be the filename of a template an admin uploaded to Open WebUI. The newest admin upload with that name is used.
- The Template Primer filter reads that setting from the Export action, so deck requests get the template's layout names without anything attached.
- Real `.potx` files no longer fail with "not a PowerPoint file".

v0.8.0 fixes the download that never appeared, and a round of deck quality:

- **Delivery:** the deck downloads in the browser on export, and the link under the message is written in a way Open WebUI v0.10+ actually renders. The old way appended it to `content`, which v0.10+ doesn't display. Re-exporting replaces the old link.
- **Rendering bugs:**
  - Bullets no longer vanish on slides that also have a table or code block (the body was resized to width 0).
  - Image-only slides keep their title.
  - A content slide's subtitle is kept.
  - A non-image `<img src>` no longer aborts the export.
  - Empty placeholders are removed.
- **Look:**
  - Built-in 16:9 theme with an accent color, slide numbers, and section dividers.
  - Text sized to its box, and set where PowerPoint reads it.
  - Images keep their aspect ratio beside the bullets.
  - Table rows and code boxes size to their text.
- **Features:**
  - `<column>` slides.
  - Bold/italic/underline/code kept.
  - `default_template` / `accent_color` / `slide_numbers` / `auto_download` Valves.
  - Open WebUI image links embed.
- **Models:**
  - Reasoning models' thinking is ignored when parsing.
  - The filter awaits Open WebUI 0.12's async Files API (it never found an attached template before), keeps a single system message, and primes follow-up edits.
  - Cut-off replies are flagged.

v0.7.3 stops blaming the outline for a missing reply: when neither the export request nor the saved chat holds an assistant message at all, the failure says so and points at chat persistence (Temporary Chat, chat-history saving) instead of telling you to reformat an outline that was never read.

v0.7.2 makes the saved-chat fallback read whichever of Open WebUI's two message stores holds more — the flat `messages` list can lag a turn behind the `history.messages` map, and preferring a non-empty flat list would shadow a history that has the assistant's reply. The no-outline diagnostic also reports both stores' sizes (`saved flat=1 hist=3`).

v0.7.1 fixes the no-outline diagnostic: it now reports whether the request carried a `chat_id` at all, and the "looks like a Temporary Chat" hint fires whenever no saved copy of the chat was found — including when the request has no `chat_id`, which is exactly the unsaved-chat case the old condition suppressed.

v0.7.0 replaces the Markdown outline grammar with an HTML one: `<slide title="..." subtitle="..." layout="...">` per slide, `<ul>`/`<li>` bullets (nesting = a nested `<ul>`), `<notes>`, `<table>`/`<tr>`/`<td>`, `<pre><code>`, and `<img>`. It's parsed with Python's standard `html.parser` instead of regex heading-level guessing, so slide boundaries are unambiguous and malformed/unclosed tags degrade gracefully instead of losing content. The template primer's content-quality guidance is also expanded — beyond "make bullets concrete," it now asks the model to distill by theme, cap the deck to what the content supports, and push supporting detail into notes.

v0.3 scope: title + bullet slides, speaker notes, images, **custom `.pptx`/`.potx` templates**, per-slide layout selection by name, and a **template primer Filter** that turns "attach a template + dump content" into a formatted deck automatically. Placeholder mapping is generic (resolved by placeholder type), so it works on any template's layouts.

v0.3.1 adds a tolerant parser (most-frequent heading level, heading-less fallbacks, inline-markdown stripping, `> ` notes) and makes every outcome report via a toast and the server log instead of the suppressible status line.

v0.5.0 renders **tables** and **code blocks**: `|`-delimited rows become a real PowerPoint table, a ``` fence becomes a monospace box, and a fence wrapping the whole outline (or a `# comment` inside one) no longer confuses the slide parser.

v0.4.0 fixes exports that came out garbled: openPPT now ignores text it appended to a message itself (its own download link or diagnostic) instead of parsing it back into slides, falls back to the newest message that actually holds an outline when the clicked one arrives empty, and shrinks long titles and dense bullet lists to fit their placeholders instead of letting them overflow across the slide.

v0.6.1 names Temporary Chat as a likely cause when the stored-chat fallback finds nothing (a `chat_id` present but no chat in storage), in both the toast and the appended diagnostic. Also adds [`outline_to_pptx.py`](outline_to_pptx.py), a standalone command-line converter for deployments where the button can't read the chat reliably at all.
