---
name: readme-to-outline
description: Turn a README (or any project doc, spec, or repo) into a presentation outline that an engineering audience and their managers actually want to sit through — problem, approach, evidence, ask — not a feature tour. Use this skill whenever the user points at a README, repo, changelog, RFC, design doc, or project directory and wants slides, a deck, a presentation, a talk, a demo, a sprint review, a tech-talk, an all-hands update, or "something to show my manager." Trigger it even when they just say "make a deck out of this repo" or "present this project" without saying the word README, and trigger it before writing any outline by hand.
---

# README → presentation outline

## The core problem this skill solves

A README is **reference documentation**: organized by feature, exhaustive, written for
someone who already decided to use the thing and is now looking something up. A
presentation is an **argument**: organized by what the audience must decide, ruthlessly
partial, written for someone who has not decided anything and is giving you 20 minutes.

Transcribing a README section-by-section into slides produces the single most common bad
engineering deck: an install-guide-shaped feature tour with no problem statement, no
evidence, and no ask. **Never map README headings onto slides one-to-one.** Re-derive the
deck from the audience's decision, then go back to the README for material.

## Workflow

### 1. Read the source completely before outlining

Read the whole README plus, if a repo is available, anything that carries the parts a
README omits: `CHANGELOG`, `docs/`, ADRs/RFCs, recent `git log`, issue titles, test names.
Do not start outlining from the first section.

### 2. Fix the audience and the ask

Every structural choice below follows from these two. If the user did not say, ask in one
short batch — do not guess silently:

- **Who is in the room?** Engineering peers / engineering manager + leadership / mixed or
  exec / external users. See `references/audiences.md` for the spine each one needs.
- **How long?** Budget one idea per slide, ~1–2 minutes per slide: 10 min → 5–7 slides,
  20–30 min → 10–14 slides. More slides than that is a document, not a talk.
- **What do you want when it ends?** Approval, headcount, a decision on a tradeoff,
  adoption, or just awareness. That answer is the last slide, and it constrains all others.

Default when the user gives nothing: engineering manager + peers, 20 minutes, ~12 slides,
ask = "adopt this / keep funding this."

### 3. Extract, then name the gaps

Mine the README for material against this map. The right column is what the deck needs;
the left is where it usually hides.

| Deck needs | Where it hides in a README |
|---|---|
| Problem / status quo pain | The opening pitch line, "Why", "Motivation", the tagline |
| What it is, in one sentence | First paragraph, repo description |
| How it works | Architecture section, the component/file table, diagrams |
| Proof it works | Test instructions, benchmarks, badges, "Development" section |
| Scope and maturity | Version history, roadmap, "v0.x scope" notes |
| Known limits | "Known gaps", troubleshooting sections, caveats, FAQ |
| Effort / cost | Install steps, dependencies, migration/update instructions |
| Next steps | Roadmap, TODOs, open issues |

Then write down what the README **cannot** tell you. READMEs almost never contain:
adoption or usage numbers, before/after metrics, time spent, headcount, incidents avoided,
user quotes, cost, or the decision history behind a tradeoff.

**Ask the user for those in one batch. Never invent them.** A fabricated metric is the
fastest way to lose the room, and managers check numbers. If a number is unavailable, say
what you have instead — a qualitative signal, a demo, a before/after code sample — and mark
the slide's `Notes:` with what to fill in.

### 4. Lay the spine

Default arc for a results/status deck to an engineering audience and its manager:

1. **Title** — project name + a one-line subtitle that states the outcome, not the topic.
2. **Bottom line up front** — the conclusion, the ask, and the one number, on slide 2.
   Executives and managers want the destination, not the journey; everything after this
   slide is support for it.
3. **The problem** — what was broken, who felt it, what it cost. Concrete and dated.
4. **The approach** — one slide, the shape of the solution. Not the class diagram.
5. **How it works** — 1–3 slides max, only the parts that make the results believable.
6. **Evidence** — the heart of the deck. Numbers, before/after, demo, test results.
7. **Tradeoffs and limits** — what you gave up, what still doesn't work.
8. **Cost and status** — what it took, where it stands, what's left.
9. **Next steps + the ask** — dated, owned, specific.

Reorder for the audience, but keep the invariants: **the conclusion comes early**, and
**evidence outweighs mechanism**. If the deck spends more slides on how it works than on
whether it worked, it is an architecture review wearing a results deck's clothes.

### 5. Write assertion headlines

The strongest evidence-based rule in presentation research: a slide headline should be a
**full sentence stating that slide's message**, with the body as its evidence. Audiences
comprehend and retain assertion-evidence slides measurably better than topic-label slides.

- Bad: `# Performance` · `# Architecture` · `# Results`
- Good: `# Export now completes in 1.2s, down from 14s` · `# One parser serves both the
  button and the CLI`

If a headline could sit unchanged on any other project's deck, it is a topic label. Rewrite
it. Every headline read in sequence should tell the whole story on its own — that sequence
is the test of whether the deck has an argument.

### 6. Write bullets that carry facts

- 3–6 bullets per slide, under ~12 words each. One idea per slide.
- Every bullet carries a **specific number, name, date, version, or action verb**. Cut any
  bullet that would survive unchanged on another project ("Improved efficiency", "Key
  considerations", "Robust and scalable").
- Push explanation into `Notes:` — the speaker says it, the slide doesn't print it. A wall
  of text competes with the presenter and loses.
- Tabular data goes in a table, not in bullets. Commands go in a fenced code block.
- No bullet should be a sentence the presenter reads aloud verbatim.

### 7. Serve the manager specifically

Managers are scanning for a specific pattern — **context → action → measurable result** —
plus four things engineers routinely omit. Before finalizing, confirm the deck answers:

- **So what?** Impact stated in the audience's units: user-facing, revenue, risk, or
  engineer-hours — not lines of code or commits.
- **How do you know?** The validation is explicit. "Tests pass" is not a result; "handles
  the 6 malformed-outline shapes that broke v0.3" is.
- **What did it cost, and what's the risk?** Volunteered, not extracted in Q&A. Naming
  limits early buys credibility; hiding them until questions destroys it.
- **What do you want from me?** A dated, owned, unambiguous ask.

Also: credit the people who did the work by name, and cut the tour of every feature — a
manager needs the two features that carry the argument, not the other nine.

### 8. Self-review before emitting

Reject and rewrite if any of these are true:

- The slide order mirrors the README's heading order.
- Any headline is a topic label rather than a claim.
- The deck reaches slide 4 without a number.
- There is no slide the audience could disagree with (no argument = a status report).
- Installation, configuration, or API surface takes more than one slide.
- Any number in the deck did not come from the source or the user.
- The last slide is "Questions?" instead of the ask.

## Output format

Emit an openPPT Markdown outline and nothing else — no preamble, no wrapping code fence.
This is the format `openppt_action.py` / `outline_to_pptx.py` parse (see the repo README's
"Outline format" section):

```markdown
# Deck Title
## One-line subtitle stating the outcome

# Full-sentence assertion headline
- Concrete bullet with a number or name
- Another, under 12 words
  - Sub-point, indented two spaces
Notes: the explanation the speaker says out loud

# Slide with a table
| Metric | Before | After |
|--------|--------|-------|
| Export time | 14s | 1.2s |
```

Rules that matter for the export to work: every slide starts with `# ` at the start of a
line; the first slide is title + `## ` subtitle with no bullets; `Notes:` (or a `> `
blockquote) after a slide's bullets becomes speaker notes; a ` ``` ` fence inside a slide
renders as a monospace box; `![alt](url)` as a bullet embeds an image.

If the user supplied a template and its layout names, tag slides with
`@layout: <name>` on the heading (e.g. `# Where we landed @layout: Section Header`) using
only names they provided.

After emitting the outline, add one short line listing the placeholders you left for the
user to fill (the numbers you refused to invent) — outside the outline is fine, the parser
ignores prose before the first `#` and openPPT stops at appended links.

## Bundled references

- `references/audiences.md` — the distinct spine, depth, and vocabulary for each audience:
  engineering peers, manager/leadership, exec, external users. Read it once the audience is
  known.
- `references/research.md` — the evidence behind the rules above (assertion-evidence,
  pyramid principle / BLUF, cognitive load, what managers screen for), with sources. Read
  it when justifying a structural choice to the user or when the user pushes back on a rule.

## Worked example

A README opening with "Let any LLM create PowerPoint decks — including models with no
function calling", an install section, a format spec, and a troubleshooting section becomes:

```markdown
# openPPT
## Deck export for local models, shipped and running in v0.6.1

# Local models can't call functions, so they can't make decks
- Gemma, Mistral, most self-hosted models: no tool calling
- Users retyped model output into PowerPoint by hand
- Every "make me a deck" request dead-ended in chat
Notes: this is the whole reason the project exists — lead with the pain, not the tagline

# The model writes Markdown; a button does the rest
- No function calling required — plain chat output
- One Action + one Filter, pasted into Open WebUI
- Template in, formatted deck out
Notes: demo here — 30 seconds, attach template, click export
```

Note what happened: the README's install steps became a single sub-bullet, the format spec
disappeared entirely (it belongs in `Notes:` or a backup slide), and the troubleshooting
section — the README's longest — became one honest "limits" slide near the end rather than
three slides in the middle.
