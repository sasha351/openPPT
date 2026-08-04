# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

openPPT is a pair of **Open WebUI Functions** (not a standalone app/package) that let any local LLM — including ones with no function calling — produce a PowerPoint deck. The model just writes an HTML outline in chat; a button converts it to a `.pptx`.

- `openppt_action.py` — the "Export to PowerPoint" **Action**. Parses the outline in an assistant message and builds the `.pptx` (`parse_outline` → `build_pptx`).
- `openppt_filter.py` — the **Filter**. Runs before the model (`inlet`), and when a `.pptx`/`.potx` template is attached to the chat, injects a system message listing the template's layout names and the outline format so the model writes slides that fit the template.

Both are installed by **pasting the raw file contents** into Open WebUI (Admin Panel → Functions). This has real constraints:

- **Single file each.** Functions can't import one another, so helpers that exist in both files (e.g. `_find_template_bytes`/`_find_layout_by_name`-style logic) are intentionally duplicated, not shared.
- Third-party deps come only from the `requirements:` line in each file's module docstring frontmatter (currently `python-pptx`); Open WebUI installs them automatically. Don't add a dependency without adding it there.
- Every code path must degrade gracefully instead of throwing — a raised exception in a Function surfaces badly in Open WebUI. `Filter.inlet` always returns `body` even on error; `Action.action` wraps its body in try/except and reports failures via `notify`/`status` rather than raising.
- Open WebUI's internal APIs (`open_webui.models.files.Files`, `open_webui.storage.provider.Storage`) are imported lazily inside methods (not at module level) and drift across Open WebUI versions — keep compat shims (see the `TypeError` fallback in `Action._find_template`) rather than assuming a fixed signature.

## Commands

Run tests directly with the system Python (no build step, no package manager config):

```bash
pip install python-pptx pydantic
python test_parser.py   # tests openppt_action.py — parse_outline, build_pptx, layout matching
python test_filter.py   # tests openppt_filter.py — template primer injection
```

Both test files are runnable both under pytest and standalone (`if __name__ == "__main__"` runs every `test_*` function and prints `ok`), because contributors don't necessarily have pytest installed. When adding a test function, it's picked up automatically by both.

To run a single test: `pytest test_parser.py::test_parse_layout_attribute` or call the function directly, e.g. `python -c "from test_parser import test_parse_layout_attribute as t; t()"`.

## Architecture: the outline grammar

The core shared concept between both files is the **outline format** — a tolerant HTML grammar for slides, parsed with Python's standard `html.parser.HTMLParser` rather than regex, since local models don't emit clean markup any more reliably than they emitted clean Markdown. Understanding `_OutlineParser`/`parse_outline` in `openppt_action.py` is the key to touching either file:

1. **Slide-start is unambiguous, not guessed.** Each `<slide title="..." subtitle="..." layout="...">` element is one slide — no heading-level frequency heuristics, no fallback chain, because the tag itself is the marker. `_extract_fences` runs first and folds any ``` fence into an equivalent HTML form (unwrapped if it opens before the first `<slide>`, escaped into a `<pre>` block if it's a real code sample after one) so `_OutlineParser` only ever has to handle one grammar.
2. **Tolerance lives in the parser, not the grammar.** `_OutlineParser` auto-closes an unclosed `<slide>`/`<li>`/`<p>` (whatever ends it — the next `<slide>`, or end of input — flushes it), and bare text sitting directly inside `<slide>` (not wrapped in `<p>`/`<li>`) still becomes a bullet instead of being dropped. HTMLParser itself never raises on malformed markup (mismatched tags, bad nesting, missing quotes), which is most of why this beats hand-rolled regex.
3. **Everything before the first `<slide>` is discarded** (chat preamble like "Here's your deck:"), and everything from the `<!-- openppt -->` `APPENDIX_MARKER` comment onward is discarded too — `handle_comment` sets a hard stop flag checked at the top of every handler, so this holds even if the model's own last `<slide>` was left unclosed.
4. **Per-slide extras are child elements**, not line-prefix conventions: `<notes>`/`<blockquote>` → speaker notes, `<img src="..." alt="...">` (bare or inside an `<li>`) → an image bullet, `<table>`/`<tr>`/`<td>`/`<th>` → a table, `<pre>`/`<pre><code>` → a code block. Nesting a `<ul>`/`<ol>` inside an `<li>` nests that bullet — flushed at the point the nested list opens, not at `</li>`, so document order comes out right (see the `emitted` flag on each `li_stack` frame).
5. **Layout selection at build time** (`_pick_layout`, `_find_layout_by_name`) is placeholder-driven, not name-driven: it inspects each template layout's title/subtitle/body placeholder counts to guess "is this a title slide / divider / content slide" role, so arbitrary templates work without hardcoding PowerPoint's built-in layout names. A slide's `layout="..."` attribute is matched case-insensitively (exact, then substring) against the template's actual layout names, falling back to role-based selection when it doesn't match. This part is unchanged by the grammar rewrite — `build_pptx` still consumes the same slide-dict shape `parse_outline` always produced.

If you change the grammar in `_OutlineParser`/`parse_outline`, update the matching prose in README.md's "Outline format" section and in `SYSTEM_PROMPT.md` / the primer text in `openppt_filter.py._template_primer`/`_no_template_primer` — those three places describe the same format to humans and to models respectively, and they drift silently if only one is edited.

## Versioning convention

The `version:` line in each file's module docstring is bumped on every change to that file, including doc-only changes — Open WebUI caches Functions by pasted content, and a stale paste in the Functions list is otherwise invisible (see commit `50afa97`). Bump `openppt_action.py`'s version and the "Confirm the version reads X" line in README.md's "Updating an existing install" section together.
