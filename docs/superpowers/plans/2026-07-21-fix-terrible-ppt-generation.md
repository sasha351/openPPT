# Fix Terrible PPT Generation Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Stop openPPT from exporting its own error message as a deck, and stop long text from overflowing and overlapping on slides.

**Architecture:** Three independent defects in `openppt_action.py`, fixed in order of the failure chain. (1) The action picks the message to export by taking the clicked message's content verbatim — including text openPPT itself appended on an earlier click — so a second click re-parses openPPT's own diagnostic into slides. Fix: mark every appended block with a sentinel, strip it before parsing, and fall back to the newest assistant message that actually contains an outline when the clicked one is empty. (2) The diagnostic is written in slide-shaped Markdown (`---` separator, `- ` bullets). Fix: rewrite it so `parse_outline` finds zero slides in it. (3) Nothing bounds text size, and python-pptx cannot recompute PowerPoint's autofit, so a long title renders at the template's full point size and spills over the body. Fix: word-wrap on, a font-size ramp for titles and bullet lists, plus PowerPoint's shrink-on-overflow flag.

**Tech Stack:** Python 3, `python-pptx` 1.0.2, plain `assert`-based tests in `test_parser.py` (runnable under pytest *and* standalone).

## Global Constraints

- **Single-file functions.** `openppt_action.py` and `openppt_filter.py` are pasted into Open WebUI individually and cannot import one another. Never add an import between them; duplicate the helper instead.
- **No new dependencies.** Only `python-pptx`, already declared in the `requirements:` frontmatter line of the module docstring.
- **Never raise out of `Action.action`.** Every path reports via `notify`/`status`. The existing try/except stays.
- **Open WebUI internals stay lazily imported** inside methods (`open_webui.models.files`, `open_webui.storage.provider`), never at module level.
- **Tests must pass both ways:** `python test_parser.py` (prints `ok`) and `pytest test_parser.py`. Every new `test_*` function is picked up automatically by both.
- **Version bump:** the `version:` line in `openppt_action.py`'s docstring goes to **0.4.0**, and README.md's "Confirm the version reads **0.3.5**" line changes to **0.4.0**. Both happen once, in Task 4 — intermediate commits are never pasted into Open WebUI, so per-task bumps are noise.
- Run every command from the repo root, `/Users/sasha/Documents/openPPT`, using `.venv/bin/python` (python-pptx 1.0.2 and pytest are installed there).

## The bug, reproduced

The photo of the broken slide is openPPT's own v0.3.5 error report, re-parsed. Verified:

```
$ .venv/bin/python -c "from openppt_action import parse_outline; ..."
TITLE: "openPPT v0.3.5 diagnostic — openPPT: nothing slide-shaped in that message. Ask the model for an outline — '# Slide Title..."
   (0, 'messages in request: 2')
   (0, 'body id: 5a45952c-d7cd-4e87-8014-eb68eff10a42')
   (0, 'content type: str, length 0')
   (0, 'What it parsed:')
   (0, '(empty)')
```

Chain: the clicked assistant message arrived with `length 0` → no slides → openPPT appended a diagnostic to that message → the diagnostic's `---` + prose line + `- ` bullets parse as a slide → the next click exported the error text, with a 200-character title at the template's title point size overlapping the body.

## File Structure

- `openppt_action.py` — all three fixes. New module constants `APPENDIX_MARKER`, `VERSION`; new helpers `_strip_appendix`, `_pick_outline`, `_title_size`, `_body_size`, `_fit`; modified `Action.action`, `_set_title`, `_fill_bullets`.
- `test_parser.py` — new tests per task; `_run_action` gains a `content` parameter so a no-outline run can be driven.
- `README.md` — version line, a note on text fitting, and the "clicked it twice" troubleshooting entry.
- `openppt_filter.py` — **untouched.** The outline grammar does not change, so `SYSTEM_PROMPT.md`, `CLAUDE.md` and the filter's primer text stay as they are.

---

### Task 1: Never re-parse openPPT's own output

**Files:**
- Modify: `openppt_action.py` (add constants after line 19; add helpers before `class Action`; replace the message-selection block at lines 433-444)
- Test: `test_parser.py`

**Interfaces:**
- Produces: `APPENDIX_MARKER: str` (module constant), `_strip_appendix(text: str) -> str`, `_pick_outline(messages: list, target=None) -> str`. Task 2 emits its diagnostic behind `APPENDIX_MARKER`; Task 2's test imports `_pick_outline`.

- [ ] **Step 1: Write the failing tests**

Add to `test_parser.py`, after `test_never_raises` (around line 128). Also add `APPENDIX_MARKER` and `_pick_outline` to the import block at the top of the file (lines 5-10), which becomes:

```python
from openppt_action import (
    APPENDIX_MARKER,
    _find_layout_by_name,
    _pick_outline,
    build_pptx,
    list_layouts,
    parse_outline,
)
```

```python
def test_pick_outline_falls_back_when_the_clicked_message_is_empty():
    """Open WebUI sometimes hands the action a message with no content at all;
    reporting 'no outline' at a chat that plainly has one is the failure that
    started the v0.3.5 diagnostic-exported-as-a-deck chain."""
    messages = [
        {"id": "a", "role": "assistant", "content": SAMPLE},
        {"id": "b", "role": "user", "content": "export it"},
        {"id": "c", "role": "assistant", "content": ""},
    ]
    assert _pick_outline(messages, "c") == SAMPLE
    assert _pick_outline(messages, None) == SAMPLE
    assert _pick_outline([], None) == ""
    assert _pick_outline([{"id": "c", "role": "assistant", "content": ""}], "c") == ""


def test_pick_outline_strips_what_openppt_appended():
    """Exporting twice must not turn our own download link into a bullet."""
    content = (
        SAMPLE
        + f"\n\n{APPENDIX_MARKER}\n\n"
        + "📊 [Download deck.pptx](/api/v1/files/x/content)\n"
    )
    slides = parse_outline(_pick_outline([{"id": "a", "role": "assistant", "content": content}], "a"))
    assert len(slides) == 4
    assert slides[-1]["bullets"] == [
        (0, {"alt": "Team offsite", "image": "https://example.com/team.jpg"})
    ]
```

- [ ] **Step 2: Run the tests to verify they fail**

Run: `.venv/bin/python -m pytest test_parser.py -k pick_outline -v`
Expected: FAIL — `ImportError: cannot import name 'APPENDIX_MARKER' from 'openppt_action'`

- [ ] **Step 3: Add the constant and helpers**

In `openppt_action.py`, after the `PPTX_MIME` line (line 19), add:

```python
# Marker leading every block openPPT appends to a chat message. Anything after
# it is our own output — a download link, a diagnostic — and is stripped before
# parsing, so clicking export twice can't turn openPPT's own text into slides.
# ponytail: an HTML comment so Markdown hides it; if a future Open WebUI
# renders it literally, swap the string, nothing else depends on its shape.
APPENDIX_MARKER = "<!-- openPPT -->"
```

Then, immediately before `class Action:` (line 411), add:

```python
def _strip_appendix(text: str) -> str:
    """Drop anything openPPT appended to a message on an earlier click."""
    return (text or "").split(APPENDIX_MARKER)[0]


def _pick_outline(messages: list, target=None) -> str:
    """Content to export: the clicked assistant message, else the newest one
    that actually holds an outline.

    Open WebUI sometimes hands the action a message whose content is empty, and
    a strict clicked-message-only read then reports "no outline" at a chat that
    plainly has one. Returns the clicked/newest text when nothing parses, so
    the caller can echo what it actually read.
    """
    candidates = []
    for msg in reversed(messages or []):
        if msg.get("role") != "assistant":
            continue
        text = _strip_appendix(msg.get("content") or "")
        if target is not None and msg.get("id") == target:
            candidates.insert(0, text)
        else:
            candidates.append(text)
    for text in candidates:
        if parse_outline(text):
            return text
    return candidates[0] if candidates else ""
```

- [ ] **Step 4: Use it in `Action.action`**

In `openppt_action.py`, replace this block (lines 433-446, from `messages = body.get` through `slides = parse_outline(content)`):

```python
            messages = body.get("messages", [])
            target = body.get("id")
            content = ""
            for msg in reversed(messages):
                if msg.get("role") == "assistant" and (target is None or msg.get("id") == target):
                    content = msg.get("content", "")
                    break
            else:  # clicked message not in the list — fall back to the latest reply
                for msg in reversed(messages):
                    if msg.get("role") == "assistant":
                        content = msg.get("content", "")
                        break

            slides = parse_outline(content)
```

with:

```python
            messages = body.get("messages", [])
            target = body.get("id")
            content = _pick_outline(messages, target)
            slides = parse_outline(content)
```

- [ ] **Step 5: Run the tests to verify they pass**

Run: `.venv/bin/python -m pytest test_parser.py -v`
Expected: PASS — all tests, including the pre-existing `test_action_inserts_the_file_and_links_to_it`.

Run: `.venv/bin/python test_parser.py`
Expected: prints `ok`

- [ ] **Step 6: Commit**

```bash
git add openppt_action.py test_parser.py
git commit -m "Pick the newest message that has an outline, and ignore openPPT's own appended text"
```

---

### Task 2: A diagnostic that can't become a deck

**Files:**
- Modify: `openppt_action.py` (collapse two markers into one; add `VERSION` constant; rewrite the `if not slides:` block)
- Test: `test_parser.py` (parameterize `_run_action`, add one test, update marker references)

**Interfaces:**
- Consumes: `APPENDIX_MARKER` and `_pick_outline` from Task 1.
- Produces: `VERSION: str` module constant (`"0.4.0"`), kept equal to the docstring's `version:` line. `_run_action(async_api: bool, content: str = SAMPLE)` in the test file. `APPENDIX_MARKER` becomes the single marker for all appended text; `DOWNLOAD_MARKER` and `_strip_appendix` no longer exist.

> **Line numbers in this task are stale.** A parallel commit (`762428c`, v0.3.6) landed after this plan was written; locate code by its text, not by line number.

**First: collapse the duplicate markers.** v0.3.6 independently added `DOWNLOAD_MARKER = "<!-- openppt:download -->"` plus a `break` inside `parse_outline`, covering the same double-export bug as Task 1's `APPENDIX_MARKER`/`_strip_appendix`. Two mechanisms for one job is exactly the duplication this codebase can't afford. Keep the better-placed one — the `break` inside `parse_outline` protects every caller — under the more general name, and delete the rest. Backwards compatibility with the `<!-- openppt:download -->` literal is not required: v0.3.6 was committed minutes before this branch and has not been pasted into any Open WebUI install.

- [ ] **Step 1: Write the failing test**

In `test_parser.py`, change the `_run_action` signature from `def _run_action(async_api: bool):` to:

```python
def _run_action(async_api: bool, content: str = SAMPLE):
```

and in its `body` dict change the assistant message to use it:

```python
    body = {
        "id": "m1",
        "messages": [
            {"id": "m0", "role": "user", "content": "deck please", "files": [{"id": "f1"}]},
            {"id": "m1", "role": "assistant", "content": content},
        ],
    }
```

Then add this test after `test_action_inserts_the_file_and_links_to_it`:

```python
def test_no_outline_diagnostic_is_not_itself_slide_shaped():
    """v0.3.5 posted an error report written in '---' + '- ' Markdown, so the
    next click exported the error report as a deck."""
    events, inserted = _run_action(False, content="Sorry, I can't help with that.")
    assert inserted == {}  # nothing was built
    (message,) = [e for e in events if e["type"] == "message"]
    posted = message["data"]["content"]
    assert APPENDIX_MARKER in posted
    assert parse_outline(posted) == []
    # and once appended to the message, it is invisible to the next click
    appended = {"id": "m1", "role": "assistant", "content": "Sorry, I can't help with that." + posted}
    assert _pick_outline([appended], "m1") == "Sorry, I can't help with that."
```

- [ ] **Step 2: Run the test to verify it fails**

Run: `.venv/bin/python -m pytest test_parser.py -k diagnostic -v`
Expected: FAIL — `assert APPENDIX_MARKER in posted` (the current diagnostic has no marker), and `parse_outline(posted)` returns one slide rather than `[]`.

- [ ] **Step 3: Collapse the two markers into one**

Four edits in `openppt_action.py`:

1. Delete the `APPENDIX_MARKER = "<!-- openPPT -->"` block Task 1 added near the top (with its comment), and delete `_strip_appendix` entirely.
2. Replace the `DOWNLOAD_MARKER` definition (the block above `_strip_md`, currently `DOWNLOAD_MARKER = "<!-- openppt:download -->"` with its four-line comment) with:

```python
# Precedes anything openPPT appends to a chat message — the download link, the
# no-outline diagnostic. An HTML comment renders invisibly in chat but survives
# a round trip back in, and parse_outline stops at this line, so re-clicking
# Export can't parse openPPT's own output back into slides (v0.3.5 exported its
# own error report as a deck).
APPENDIX_MARKER = "<!-- openppt -->"
```

3. In `parse_outline`, change the break line and the docstring sentence that names the constant:

```python
        if line == APPENDIX_MARKER:
            break  # everything below is openPPT's own appended output
```

```python
    Text before the first slide marker (chat prose) is ignored. Returns []
    if no slides found. Everything from the `APPENDIX_MARKER` line onward is
    ignored too, so re-exporting a message openPPT already appended to (a
    download link, a diagnostic) doesn't parse that text back into slides.
```

4. In `_pick_outline`, drop the `_strip_appendix` call — `parse_outline` now handles the marker:

```python
        text = msg.get("content") or ""
```

In `test_parser.py`, change the `DOWNLOAD_MARKER` import to `APPENDIX_MARKER` (the import list is alphabetical) and update its two other uses to `APPENDIX_MARKER`. Task 1's `APPENDIX_MARKER` import stays; there must be exactly one marker name in the file afterwards.

Verify no orphans remain:

```bash
grep -rn "DOWNLOAD_MARKER\|_strip_appendix" openppt_action.py test_parser.py
```

Expected: no output.

- [ ] **Step 4: Rewrite the diagnostic**

In `openppt_action.py`, add directly below the `APPENDIX_MARKER` block:

```python
# Kept equal to the docstring's 'version:' line — a stale paste in Open WebUI's
# Functions list is otherwise invisible, and the diagnostic is where it shows.
VERSION = "0.4.0"
```

Then replace the whole `if not slides:` block (from `if not slides:` through its `return`) with:

```python
            if not slides:
                help_text = (
                    "openPPT: nothing slide-shaped in that message. Ask the model for "
                    "an outline — a '# Slide Title' heading per slide, with '-' "
                    "bullets under it."
                )
                await notify(help_text, "warning")
                await status(help_text, done=True)
                # Toasts and the status line get truncated, so the detail goes in
                # the chat. Written so parse_outline finds nothing in it: no '---'
                # rule, no bullet list, no lone '**bold**' line — otherwise the
                # next click exports this diagnostic as a deck (v0.3.5 did).
                seen = content[:400] if content else "(nothing — the message was empty)"
                await emit(
                    {
                        "type": "message",
                        "data": {
                            "content": (
                                f"\n\n{APPENDIX_MARKER}\n"
                                f"**openPPT {VERSION} diagnostic** — {help_text} "
                                f"(messages in request: {len(messages)}, clicked id: "
                                f"`{target}`, read {len(content)} characters)\n\n"
                                f"What openPPT read from that message:\n\n"
                                f"```\n{seen}\n```\n"
                            )
                        },
                    }
                )
                return
```

- [ ] **Step 5: Point the download emit at the renamed marker**

The success message already sits behind a marker (v0.3.6 added it); it just names the deleted constant. In `openppt_action.py`, in the `await emit` that posts the download link, change `f"\n\n{DOWNLOAD_MARKER}\n📊 [Download {name}]({full_url})"` to:

```python
                            f"\n\n{APPENDIX_MARKER}\n📊 [Download {name}]({full_url})"
```

Leave the rest of that emit — the trailing `f"\n\n{full_url}"` line and the comment above it — alone, except to update the comment's `parse_outline` reference if it names the old constant.

- [ ] **Step 6: Run the tests to verify they pass**

Run: `.venv/bin/python -m pytest test_parser.py -v`
Expected: PASS — all tests, including the v0.3.6 download-marker test and Task 1's `_pick_outline` tests.

Run: `.venv/bin/python test_parser.py && .venv/bin/python test_filter.py`
Expected: `ok` from each.

- [ ] **Step 7: Commit**

```bash
git add openppt_action.py test_parser.py
git commit -m "Write the no-outline diagnostic so the parser finds no slides in it"
```

---

### Task 3: Shrink text to fit the slide

**Files:**
- Modify: `openppt_action.py` (import `MSO_AUTO_SIZE`; add `_title_size`/`_body_size`/`_fit`; rewrite `_set_title` at lines 349-352 and `_fill_bullets` at lines 355-360)
- Test: `test_parser.py`

**Interfaces:**
- Consumes: nothing from Tasks 1-2.
- Produces: `_title_size(text: str) -> int | None` (None = keep the template's size), `_body_size(bullets: list) -> int`, `_fit(text_frame) -> None`. All sizes are point values, applied via `Pt()`.

- [ ] **Step 1: Write the failing tests**

Add to `test_parser.py` after `test_body_bullets_land_in_a_placeholder` (around line 179):

```python
def test_long_title_shrinks_instead_of_overlapping_the_body():
    """python-pptx can't recompute PowerPoint's autofit, so a long title keeps
    the template's ~44pt and spills across the slide over the bullets."""
    long_title = "openPPT: nothing slide-shaped in that message. " * 4
    out = Presentation(io.BytesIO(build_pptx(parse_outline(f"# {long_title}\n- a\n"))))
    frame = out.slides[0].shapes.title.text_frame
    assert frame.word_wrap is True
    assert frame.paragraphs[0].font.size.pt <= 20


def test_short_title_keeps_the_templates_own_size():
    out = Presentation(io.BytesIO(build_pptx(parse_outline("# Q3 Results\n- a\n"))))
    # None = inherited from the layout; we must not restyle a title that fits
    assert out.slides[0].shapes.title.text_frame.paragraphs[0].font.size is None


def test_dense_bullet_list_shrinks():
    outline = "# Ideas\n" + "".join(f"- point {i}\n" for i in range(14))
    out = Presentation(io.BytesIO(build_pptx(parse_outline(outline))))
    (body,) = [p for p in out.slides[0].placeholders if p.placeholder_format.idx == 1]
    assert len(body.text_frame.paragraphs) == 14
    assert body.text_frame.paragraphs[0].font.size.pt <= 14
```

- [ ] **Step 2: Run the tests to verify they fail**

Run: `.venv/bin/python -m pytest test_parser.py -k "title_shrinks or templates_own_size or dense_bullet" -v`
Expected: 2 FAIL, 1 PASS — `test_long_title_shrinks_instead_of_overlapping_the_body` fails on `assert frame.word_wrap is True` (it is `None`), `test_dense_bullet_list_shrinks` fails with `assert 20.0 <= 14`. `test_short_title_keeps_the_templates_own_size` already passes and guards against over-correcting.

- [ ] **Step 3: Add the sizing helpers**

In `openppt_action.py`, extend the pptx imports (lines 15-17) to:

```python
from pptx import Presentation
from pptx.enum.shapes import PP_PLACEHOLDER
from pptx.enum.text import MSO_AUTO_SIZE
from pptx.util import Pt
```

Then replace `_set_title` and `_fill_bullets` (lines 349-360) with:

```python
def _title_size(text: str):
    """Point size for a title long enough to overflow, or None to keep the
    template's. python-pptx never recomputes PowerPoint's autofit, so without
    this a 200-character title renders at full size across the whole slide."""
    for limit, size in ((50, None), (80, 32), (120, 26), (200, 20)):
        if len(text) <= limit:
            return size
    return 16


def _body_size(bullets) -> int:
    """Point size for a bullet list, shrinking as it gets denser. Long bullets
    wrap, so count the lines they'll take rather than the bullets."""
    lines = sum(1 + len(text) // 60 for _, text in bullets)
    for limit, size in ((5, 20), (8, 16), (12, 14)):
        if lines <= limit:
            return size
    return 12


def _fit(text_frame) -> None:
    """Wrap, and let PowerPoint shrink further than our estimate if it must."""
    text_frame.word_wrap = True
    text_frame.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE


def _set_title(slide, text: str) -> None:
    ph = _title_placeholder(slide)
    if ph is None or not ph.has_text_frame:
        return
    ph.text_frame.text = text
    _fit(ph.text_frame)
    size = _title_size(text)
    if size:
        for p in ph.text_frame.paragraphs:
            p.font.size = Pt(size)


def _fill_bullets(text_frame, bullets) -> None:
    _fit(text_frame)
    size = Pt(_body_size(bullets))
    for j, (level, text) in enumerate(bullets):
        p = text_frame.paragraphs[0] if j == 0 else text_frame.add_paragraph()
        p.text = text
        p.level = level
        p.font.size = size
```

- [ ] **Step 4: Run the tests to verify they pass**

Run: `.venv/bin/python -m pytest test_parser.py -v`
Expected: PASS — all tests, including `test_build_with_template_clears_sample_slides`, which reads `slides[0].shapes.title.text` and so proves `_set_title` still writes through the new `text_frame.text` path.

Run: `.venv/bin/python test_parser.py && .venv/bin/python test_filter.py`
Expected: `ok` from each.

- [ ] **Step 5: Eyeball a real deck**

```bash
.venv/bin/python -c "
from openppt_action import parse_outline, build_pptx
title = '# A title long enough that it used to run straight across the slide and over the bullets'
bullets = '\n'.join('- bullet number %d, with enough words on it to wrap' % i for i in range(12))
deck = build_pptx(parse_outline(title + '\n' + bullets))
open('/tmp/openppt-check.pptx', 'wb').write(deck)
print(len(deck), 'bytes')
"
open /tmp/openppt-check.pptx
```

Expected: the title wraps inside its placeholder and does not touch the bullet list; the 12 bullets fit on the slide.

- [ ] **Step 6: Commit**

```bash
git add openppt_action.py test_parser.py
git commit -m "Shrink long titles and dense bullet lists to fit their placeholders"
```

---

### Task 4: Version bump and docs

**Files:**
- Modify: `openppt_action.py:4` (docstring `version:` line), `README.md:40`, `README.md:36`, `README.md:145`

**Interfaces:**
- Consumes: `VERSION = "0.4.0"` added in Task 2 — the docstring line must match it exactly.

- [ ] **Step 1: Bump the version in the docstring**

In `openppt_action.py`, line 4, change `version: 0.3.5` to:

```
version: 0.4.0
```

- [ ] **Step 2: Verify the two version strings agree**

Run:

```bash
.venv/bin/python -c "
import re, openppt_action as a
doc = re.search(r'version:\s*(\S+)', a.__doc__).group(1)
assert doc == a.VERSION, (doc, a.VERSION)
print('version', doc)
"
```

Expected: `version 0.4.0`

- [ ] **Step 3: Update README.md**

In `README.md`, change `Confirm the version reads **0.3.6** afterwards.` (in "Updating an existing install") to:

```markdown
Confirm the version reads **0.4.0** afterwards.
```

In `README.md`, append this sentence to the end of the "Button shows but clicking it seems to do nothing?" paragraph:

```markdown
 If openPPT reports no outline, it posts a short diagnostic under the message saying what it actually read; that diagnostic is deliberately not slide-shaped, so clicking export again re-reads your outline rather than exporting the error report (which is what v0.3.5 did).
```

At the end of `README.md`, after the existing version-notes paragraphs, add:

```markdown
v0.4.0 fixes exports that came out garbled: openPPT now ignores text it appended to a message itself (its own download link or diagnostic) instead of parsing it back into slides, falls back to the newest message that actually holds an outline when the clicked one arrives empty, and shrinks long titles and dense bullet lists to fit their placeholders instead of letting them overflow across the slide.
```

- [ ] **Step 4: Full verification**

Run: `.venv/bin/python test_parser.py && .venv/bin/python test_filter.py && .venv/bin/python -m pytest -q`
Expected: `ok`, `ok`, and pytest reporting all tests passed with no failures.

- [ ] **Step 5: Commit and push**

```bash
git add openppt_action.py README.md docs/superpowers/plans/2026-07-21-fix-terrible-ppt-generation.md
git commit -m "v0.4.0: stop exporting openPPT's own diagnostic, and fit text to the slide"
git push origin main
```

- [ ] **Step 6: Re-paste into Open WebUI**

Admin Panel → Functions → **Edit** on *Export to PowerPoint* → select all → paste the current `openppt_action.py` → **Save**. Confirm the list shows **0.4.0**. Then, in the chat that produced the broken deck, click **Export to PowerPoint** again — it should now find the real outline rather than the diagnostic.

---

## Notes for the implementer

- **Line numbers are from the pre-change file.** Tasks 1-3 each shift them; locate the code by its text, not by line number, if you work out of order.
- **Don't "fix" the empty-message cause upstream.** Why Open WebUI sometimes sends `content: ""` is unknown and lives in its frontend, not here; Task 1 makes the action survive it, which is the whole fix available from this side.
- **The outline grammar does not change in this plan.** If you find yourself editing `parse_outline`, stop — `SYSTEM_PROMPT.md`, `README.md`'s "Outline format" section and `openppt_filter.py._primer` would all have to move with it, and none of these three bugs live there.
