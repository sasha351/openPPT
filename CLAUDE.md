# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

openPPT is a pair of **Open WebUI Functions** (not a standalone app/package) that let any local LLM — including ones with no function calling — produce a PowerPoint deck. The model just writes a Markdown outline in chat; a button converts it to a `.pptx`.

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

To run a single test: `pytest test_parser.py::test_parse_layout_directives` or call the function directly, e.g. `python -c "from test_parser import test_parse_layout_directives as t; t()"`.

## Architecture: the outline grammar

The core shared concept between both files is the **outline format** — a tolerant Markdown grammar for slides, since local models rarely emit a strict format. Understanding `parse_outline` in `openppt_action.py` is the key to touching either file:

1. **Slide-start detection is adaptive, not fixed at `# `.** `_slide_level` scans all heading lines up front and picks whichever `#` depth is *most frequent* as the slide marker — so a deck written entirely in `## ` per slide treats a single leading `# ` as the deck title, not a stray bullet under slide 1.
2. **Heading-less fallback chain** (only when no `#` headings exist at all): `Slide 3: Title` prefixes → lone `**Bold Title**` lines → `---` separators. This is why a deck using *only* `---` loses its first section (there's no marker before it) — documented as a known gap in README.md.
3. **Everything before the first slide marker is discarded** (chat preamble like "Here's your deck:").
4. **Per-slide extras** layer on top of bullets: `Notes:`/`> ` lines become speaker notes, `![alt](url)` bullets become images, and a layout can be named either inline on the heading (`@layout:` or `{layout: ...}`) or on its own `Layout:` line.
5. **Layout selection at build time** (`_pick_layout`, `_find_layout_by_name`) is placeholder-driven, not name-driven: it inspects each template layout's title/subtitle/body placeholder counts to guess "is this a title slide / divider / content slide" role, so arbitrary templates work without hardcoding PowerPoint's built-in layout names. A named `@layout:` is matched case-insensitively (exact, then substring) against the template's actual layout names, falling back to role-based selection when it doesn't match.

If you change the grammar in `parse_outline`, update the matching prose in README.md's "Outline format" section and in `SYSTEM_PROMPT.md` / the primer text in `openppt_filter.py._primer` — those three places describe the same format to humans and to models respectively, and they drift silently if only one is edited.

## Versioning convention

The `version:` line in each file's module docstring is bumped on every change to that file, including doc-only changes — Open WebUI caches Functions by pasted content, and a stale paste in the Functions list is otherwise invisible (see commit `50afa97`). Bump `openppt_action.py`'s version and the "Confirm the version reads X" line in README.md's "Updating an existing install" section together.
