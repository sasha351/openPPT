# openPPT v1.0 — Product Spec & Phased Plan

## Context

openPPT v0.1 is a single-file Open WebUI Action (`openppt_action.py`) that parses a strict `# heading` Markdown outline from the last assistant message and exports a `.pptx` via python-pptx. Goal: turn it into a full-fledged product — a polished, robust Open WebUI plugin published on the community hub, working with any local LLM (no function calling needed).

**Benchmark use case**: user uploads a set of README.md docs and email transcripts into an Open WebUI chat/Knowledge collection; the configured local LLM (via Open WebUI's RAG) synthesizes a project-report outline; the user iterates in chat, clicks **Export to PowerPoint**, gets a branded project-report deck.

## Decisions made with user

| Decision | Choice |
|---|---|
| Product form | Polished single-file Open WebUI plugin (community hub) |
| Deck features | Themes/templates, tables & code blocks, speaker notes, images |
| Parser | Tolerant — best-effort on messy local-model output |
| Doc ingestion | Open WebUI native uploads/Knowledge RAG (no custom ingestion code) |
| Benchmark LLM | Whatever model is configured in Open WebUI, driven via its API |

## Inputs & Outputs

**Inputs**
1. Assistant message Markdown (the message the export button is clicked under) — tolerant grammar (see Phase 1).
2. Source documents (READMEs, email transcripts) — via Open WebUI file upload / Knowledge collections; openPPT never touches these directly, the LLM+RAG does.
3. Config via Valves: default font size, filename pattern, template file (admin: filesystem path or Open WebUI file ID of a corporate `.pptx`), max image size.
4. Optional per-deck template: corporate `.pptx` uploaded by user.

**Outputs**
1. `.pptx` (title slide, content slides with nested bullets, tables, code boxes, images, speaker notes; themed from template) stored via Open WebUI Files, surfaced as a download link appended to the chat.
2. Status events (progress, actionable errors — e.g. "no structure found, ask the model for an outline").

## Constraints

- **Single file.** Open WebUI functions are installed by pasting one file. All code stays in `openppt_action.py`; tests import it directly (existing pattern in `test_parser.py`).
- Pure-stdlib + python-pptx only (auto-installed from frontmatter `requirements`).
- Must not crash the chat: every failure ends in a status message, never an exception (existing pattern, keep it).
- Open WebUI internal APIs (`Files`, `Storage`) drift across versions — keep the existing try/except compat shims, add a REST fallback only if a supported version breaks.

---

## Phase 1 — Tolerant parser v2

Rewrite `parse_outline` to accept real local-LLM output. Slide dict grows to
`{"title", "subtitle", "bullets": [(level, text)], "notes": str, "table": [[cells]], "code": str, "images": [url]}`.

Grammar (best-effort, in priority order):
- Slide starts: `# `, `## `, `### ` headings; `---` separators (Marp/pandoc style); `Slide 3:` / `**Slide 3: Title**` prefixes; lone `**Bold Title**` lines between blank lines.
- Bullets: `-`, `*`, `+`, `1.`/`1)`, indent = nesting (existing logic). Strip inline `**`/`*`/`` ` `` markdown.
- Speaker notes: lines starting `> ` (blockquote) or `Note:`/`Notes:` collect into `notes`.
- Tables: consecutive `|`-delimited lines → `table` (separator row `|---|` dropped).
- Fenced code: ` ``` ` blocks → `code` (currently skipped entirely — fix).
- Images: `![alt](url)` → `images`.
- First heading + immediate sub-heading with no bullets = title slide + subtitle (existing behavior, keep).
- Heading level detection: whichever level (`#`/`##`/`###`) occurs most often marks slides; higher levels above it are ignored/become the deck title.

**Tests (extend `test_parser.py`, plain asserts, no framework):**
- Existing strict-format tests keep passing (regression).
- One test per grammar feature: `---` decks, `##` decks, `Slide N:` decks, bold-title decks, notes, tables, code fences, images, inline-markdown stripping.
- Fixture-based regression: `fixtures/model_outputs/*.md` — real recorded outputs from local models (Gemma, Llama, Qwen) answering "make me a presentation"; assert each parses to ≥1 slide with a nonempty title. Collect these fixtures manually while developing.
- `parse_outline` never raises: feed it garbage (binary-ish text, emoji, 10k-line input) and assert it returns a list.

## Phase 2 — Renderer v2 (features + themes)

Extend `build_pptx(slides, template=None, font_size=20)`:
- Template: `Presentation(template_path_or_stream)` — corporate `.pptx` gives branding for free (python-pptx inherits masters/layouts). Fallback: bundled default. Layout picking by placeholder inspection, not index, so arbitrary templates work (find title/body placeholders; fall back to blank layout + textboxes).
- Tables → `shapes.add_table` below/instead of body text.
- Code → textbox with monospace font, light fill.
- Speaker notes → `slide.notes_slide.notes_text_frame.text`.
- Images → download (stdlib `urllib`, size/timeout capped via Valve) or fetch from Open WebUI Files when URL is `/api/v1/files/...`; place right half or below bullets; skip silently on failure (deck still exports).
- Overflow guard: >N bullets → shrink font stepwise (python-pptx has no autofit measurement; simple heuristic is fine).

**Tests:**
- Round-trip: build from fixture slides, reopen with `Presentation(io.BytesIO(data))`, assert slide count, title texts, notes text, table cell values, picture shape present.
- Template test: build with a minimal custom template fixture (`fixtures/template.pptx`), assert its layouts were used.
- Image failure test: unreachable URL → deck still builds, no image shape.

## Phase 3 — Open WebUI integration hardening

- **Valves** (`class Valves(BaseModel)` on the Action): `TEMPLATE_FILE` (path or file ID), `FONT_SIZE`, `MAX_IMAGE_MB`; `UserValves` for per-user template override.
- **Right message**: export the message the button was clicked on (`body["id"]` / clicked-message id), not blindly the last assistant message; fall back to last assistant message.
- **Errors**: every failure path emits a specific, actionable status (no structure found / template unreadable / storage failed).
- **Compat**: keep Storage/Files try-except shims; verify on current stable Open WebUI + one version back.

**Tests:**
- `test_action.py`: drive `Action.action()` with a fake `body`, stub `__event_emitter__`, and monkeypatched `open_webui.models.files`/`storage.provider` modules (inject fakes into `sys.modules`); assert: file inserted with `.pptx` content-type, download-link message emitted, "no slides" path emits guidance, exception path emits failure status (never raises).
- Manual checklist in README: install into a live Open WebUI, click through happy path + no-outline path.

## Phase 4 — Project Report benchmark (the flagship use case)

- Ship the **"Project Report Builder"** preset in the README (and as a snippet file): system prompt that instructs the model to synthesize uploaded READMEs + email transcripts into the outline grammar — sections like Overview, Timeline, Decisions (from emails), Status, Risks, Next Steps; use `> Note:` for detail that shouldn't clutter slides.
- Fixtures: `fixtures/project_report/` — 3–5 realistic README.md files + email-transcript `.txt` files for a fictional project.
- `benchmark/run_benchmark.py`: drives a running Open WebUI via its REST API (`OPENWEBUI_URL`, `OPENWEBUI_API_KEY` env vars): upload fixtures as files/knowledge → create chat with the preset system prompt → ask for a project report deck → take the response through `parse_outline` + `build_pptx` → write `report.pptx` + print a structural scorecard.

**Tests:**
- Deterministic: recorded real model responses for the project-report prompt go into `fixtures/model_outputs/`, covered by the Phase 1 fixture regression.
- Live benchmark (manual / opt-in, needs running Open WebUI): asserts the loose structural bar — ≥5 slides, title slide present, ≥1 slide with speaker notes, no slide with >8 bullets, deck opens in python-pptx. This is the product acceptance test.

## Phase 5 — Release

- README rewrite: install, outline grammar reference, valves, Project Report walkthrough (screenshots), template/branding how-to.
- Version bump to 1.0.0 in frontmatter; CHANGELOG.md.
- Publish to Open WebUI community hub (https://openwebui.com → Functions); add hub link + license (MIT) to repo.

**Tests:** full suite green (`python test_parser.py && python test_action.py`); live benchmark run against current stable Open WebUI recorded in the README.

---

## Files

- `openppt_action.py` — everything shippable (parser, renderer, Action, Valves). Reuse existing `parse_outline`/`build_pptx`/status patterns.
- `test_parser.py`, `test_action.py` — assert-based, `python file.py` runnable.
- `fixtures/model_outputs/*.md`, `fixtures/project_report/*`, `fixtures/template.pptx`
- `benchmark/run_benchmark.py`
- `README.md`, `CHANGELOG.md`, `LICENSE`

## Verification (end-to-end)

1. `pip install python-pptx && python test_parser.py && python test_action.py` — all green.
2. Paste `openppt_action.py` into a live Open WebUI with a local model (Ollama), upload `fixtures/project_report/*`, use the Project Report Builder preset, ask for a report, click Export, open the `.pptx` in PowerPoint/Keynote — title slide, sectioned content, notes, branding from template.
3. `OPENWEBUI_URL=... OPENWEBUI_API_KEY=... python benchmark/run_benchmark.py` — scorecard passes the structural bar.

Phases are sequential; each ends with its tests green before the next starts.
