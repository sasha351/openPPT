# Audience variants

Same README, four different decks. Pick the spine that matches who is in the room, then
apply the shared rules from `SKILL.md` (assertion headlines, concrete bullets, conclusion
early).

The single question that separates them: **what does this audience do with the information
after the talk?** Peers write code against it. Managers allocate against it. Execs decide
against it. Users adopt against it. Cut everything that doesn't serve that action.

---

## Engineering peers (tech talk, design review, guild session)

They can absorb mechanism, and they will judge you on whether the design is sound. They
still do not want the README read aloud.

**Spine (10–14 slides)**

1. Title + outcome subtitle
2. The problem, stated as a constraint they'd recognize
3. Why the obvious approach fails — this is what earns their attention
4. The design, one slide, at the box-and-arrow level
5. The interesting mechanism, 2–4 slides (this is the only audience that gets this budget)
6. Tradeoffs taken and alternatives rejected, with reasons
7. Evidence: benchmarks, tests, failure modes handled
8. What's still broken / where you want review
9. How to use or contribute — one slide, links

**Depth:** highest. Real code, real signatures, real edge cases.
**Vocabulary:** unrestricted, but define project-specific nouns once.
**The ask:** review, adoption, contribution, or a design decision.
**Watch for:** the failure mode here is the opposite of the others — under-explaining the
hard part and over-explaining the setup. Peers will forgive a missing agenda slide; they
will not forgive hand-waving the one interesting problem.

---

## Engineering manager + leadership (results review, project update, sprint/quarter close)

The default audience for this skill. They are tracking **context → action → measurable
result** and allocating people and time against what they hear.

**Spine (8–12 slides)**

1. Title + outcome subtitle
2. **Bottom line up front**: what happened, the one number, what you want
3. The problem in business or user terms, with what it cost
4. What we did — one slide, no architecture
5. Evidence, 2–3 slides — the center of gravity of the deck
6. Tradeoffs and known limits, volunteered
7. Cost and current status: effort spent, what's shipped vs. in flight
8. Risks and what could still go wrong
9. Next steps + the ask: dated, owned, specific

**Depth:** one level below the surface. Architecture only where it makes a result
believable or a risk legible.
**Vocabulary:** business/user units. Engineer-hours, incidents, latency users feel, support
tickets — never lines of code, commit counts, or story points as proof of value.
**The ask:** approval, headcount, priority, a go/no-go, or continued investment.
**Watch for:** the classic miss is spending six slides on how it works and one on whether
it worked. Invert that. Also: name the people who did the work — managers are running
performance calibration in the back of their mind, and unattributed work is invisible work.

---

## Executives / cross-functional leadership (steering committee, all-hands, budget review)

Assume 5–10 minutes, interruptions, and no shared technical vocabulary. Assume they read
only the first two slides.

**Spine (5–7 slides)**

1. Title + outcome subtitle
2. **The answer**: recommendation, impact, ask — complete on its own slide
3. Why it matters: the business problem and its cost
4. Proof, one slide, one chart or one table
5. Risk and what it costs to continue
6. Decision requested, with the date it's needed by

**Depth:** surface only. Everything else is an appendix slide after the ask, reached only
if someone asks.
**Vocabulary:** money, risk, customers, time. Zero implementation detail. If a term needs
defining, replace the term.
**The ask:** a decision or a resource, stated as a sentence with a deadline.
**Watch for:** building to a conclusion. Engineering training says set up the problem, then
reveal the result; executive attention runs the other way. Lead with the destination and
let them pull for the journey. If you are cut off after slide 2, the deck must still have
worked.

---

## External users / adopters (launch talk, conference, community demo)

Closest to the README's own purpose, and therefore the easiest to let collapse back into a
feature tour. Resist it.

**Spine (8–12 slides)**

1. Title + what it does in one line
2. The pain they already have — they must recognize themselves here
3. The demo. Live if possible, and give it real time — more than half the slot
4. How it fits their existing setup: install, integration, dependencies
5. What it doesn't do yet, honestly
6. Getting started: the first command, the link, the docs

**Depth:** operational — how a user interacts with it, not how it's built.
**Vocabulary:** their workflow, not your internals.
**The ask:** try it, star it, file issues, adopt it.
**Watch for:** the tour of all nine features. Show the two that solve their pain, and let
the README carry the rest — this is the one audience that will actually go read it.
