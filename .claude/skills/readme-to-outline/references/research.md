# Why the rules in SKILL.md are the rules

Read this when the user pushes back on a structural choice, or when you need to justify
cutting something they wrote. Each rule below is stated with the evidence behind it.

---

## 1. Lead with the conclusion (BLUF / Minto pyramid)

**Rule:** slide 2 states the answer, the impact, and the ask. Everything after supports it.

Barbara Minto's pyramid principle, the house style of McKinsey and the standard for
executive communication, puts a single governing assertion at the top, supported by ~3
points, each backed by evidence — "summarize above, provide details below." BLUF (bottom
line up front, originally US Army doctrine) is the same instinct compressed: state the
cost, the benefit, and the timeframe first.

**Why engineers get this wrong:** the top-down structure is the inverse of the scientific
paper and the design doc — background, method, results, conclusion — which is the form
engineers are trained in. That form is correct for a document the reader controls the pace
of, and wrong for a talk where the audience's attention is spent early and they may cut you
off. Executives want the destination, then decide whether they need the journey.

- [The Pyramid Principle: How to Structure Presentations to Executives](https://beyondthebacklog.com/2023/12/11/pyramid-principle/)
- [The Minto Pyramid Principle Explained](https://www.betterup.com/blog/minto-pyramid)
- [The Pyramid Principle Applied | Management Consulted](https://managementconsulted.com/pyramid-principle/)
- [BLUF, AIM, AIDA: 17 Presentation Frameworks](https://benjaminball.com/blog/guide-to-powerful-presentation-frameworks/)

---

## 2. Sentence headlines, not topic labels (assertion-evidence)

**Rule:** every slide headline is a full sentence stating that slide's message; the body is
its evidence.

This is the best-evidenced rule in the skill. Michael Alley (Penn State engineering
communication) has studied the assertion-evidence structure — a sentence headline supported
by visual evidence — against the common topic-phrase-plus-bullets slide for three decades.
Controlled studies found the assertion-evidence audience understood and retained the
content better at statistical significance (p < .01), held fewer misconceptions, and
reported lower perceived cognitive load.

A second finding matters more for outline generation than for delivery: **presenters who
built assertion-evidence slides understood their own content better** than those using
topic-subtopic slides. Being forced to state each slide's claim as a sentence exposes
slides that have no claim. That is exactly the failure mode of a README-derived deck, where
every heading is a noun phrase inherited from documentation.

- [Rethinking the Design of Presentation Slides: the Assertion-Evidence Structure (Penn State)](https://writing.engr.psu.edu/slides_references.html)
- [How the design of presentation slides affects audience comprehension (Garner & Alley)](https://pure.psu.edu/en/publications/how-the-design-of-presentation-slides-affects-audience-comprehens/)
- [Michael Alley: Assertion-Evidence Slides for a Research Talk (iBiology)](https://www.ibiology.org/ibiology_podcasts/michael-alley-part-2-assertion-evidence-slides-for-a-research-talk/)

---

## 3. Few bullets, short bullets, one idea per slide

**Rule:** 3–6 bullets, under ~12 words, one idea per slide, explanation pushed to speaker
notes.

The mechanism is the redundancy principle: when an audience reads text while listening to a
speaker deliver the same content, the two inputs compete for working memory and
comprehension drops — roughly, 1 + 1 ≈ 0. A dense slide doesn't reinforce the speaker, it
replaces them, and the audience reads ahead and stops listening. Common guidance caps a
slide at ~7 lines of ~7 words; tighter is better, and splitting across slides beats
cramming.

This is why `Notes:` exists in the openPPT format and why the skill pushes detail there
rather than dropping it: the content survives, it just moves to the presenter's mouth.

- [How to Avoid Death by PowerPoint (Lucid)](https://lucid.co/blog/how-to-avoid-death-by-powerpoint)
- [Death by PowerPoint — how to make bad presentations (SlideLizard)](https://slidelizard.com/en/blog/bad-presentations)
- [10 ways to avoid death by bullet points](https://presentitude.com/10-ways-avoid-death-bullet-points/)

---

## 4. What managers actually screen for

**Rule:** context → action → measurable result, plus cost, risk, and an explicit ask.

Technical sponsors and engineering managers are results-first: the result determines
whether they invest attention in your process and justification at all. The pattern they
are matching is context → action → measurable result, and they want all three — an action
with no measured result reads as activity, and a result with no context can't be judged.

Two more manager-specific findings shape the skill's step 7:

- **Calibrate detail to the audience's function.** For an engineering audience, expand the
  technical section; for a finance or business audience, compress it to a minimum and
  expand impact in their units. Same project, different deck.
- **End with the expectation, explicitly.** The closing slide restates the key message and
  names the action requested, so the discussion afterward is directed at what matters
  rather than wandering. "Questions?" forfeits that.

On honesty about limits: guidance for technical demos is consistent that limitations and
risks should be addressed openly in the deck rather than held back until Q&A — volunteering
them reads as command of the problem, while being surprised by them in questions reads as
not having looked. Validation should be shaped around "how do you know it worked?" — a
repeatable demo, test data, or metrics, rather than claims.

- [Technical Presentation — MIT Comm Lab](https://mitcommlab.mit.edu/meche/commkit/technical-presentation/)
- [What Engineering Managers Look For (DEV)](https://dev.to/nataliaherself/what-engineering-managers-look-for-in-performance-reviews-493l)
- [How to present to management: a guide for developers and engineers (InfoWorld)](https://www.infoworld.com/article/2255468/how-to-present-to-management-a-guide-for-developers-and-engineers.html)
- [Successful Engineering Presentations](https://myengineeringtools.com/Soft_Skills/Engineering_Presentation.html)
- [How to give a demo presentation (UW CSE 403)](https://courses.cs.washington.edu/courses/cse403/24wi/project/demo.html)

---

## 5. Why a README specifically resists becoming a deck

**Rule:** never map README headings onto slides one-to-one.

The Diátaxis framework splits documentation into four modes, and a README is predominantly
**reference**: factual, structured, lookup-friendly, consulted while working, and
deliberately *not* narrative. Explanation — the mode that discusses why, explores
alternatives, and covers tradeoffs — is a different document type entirely, and good
documentation practice keeps it *out* of reference material.

So the narrative content a presentation is made of is not merely buried in a README; a
well-written README has been actively purged of it. That is why step 3 of the skill is an
extraction map plus an explicit gap list: the deck's spine (problem, evidence, tradeoffs,
cost, ask) mostly has to come from the user or from the repo's history, not from the README
prose. And it is why a section-by-section transcription produces a feature tour — you are
faithfully reproducing a document whose organizing principle is lookup, in a medium whose
organizing principle is argument.

- [Diátaxis](https://diataxis.fr/)
- [Reference and explanation (Diátaxis source)](https://github.com/evildmp/diataxis-documentation-framework/blob/main/reference-explanation.rst)
- [What is Diátaxis? (I'd Rather Be Writing)](https://idratherbewriting.com/blog/what-is-diataxis-documentation-framework)

---

## 6. Demos beat slides about demos

**Rule:** for user- and peer-facing talks, give the demo more than half the slot.

Course guidance for project demos puts the demo at half the slot or a bit more, with slides
introducing and motivating it — what problem it solves, why it beats current practice.
Sprint-review guidance goes further and argues against formal slide decks entirely, in
favor of working software plus interactive Q&A, structured around user stories and business
outcomes rather than technical features.

Practical consequence for outlining: when the format is a demo, the deck's job is setup,
framing, and the ask — not narration of what the audience is about to watch. Keep it to a
handful of slides around the demo.

- [How to give a demo presentation (UW CSE 403)](https://courses.cs.washington.edu/courses/cse403/24wi/project/demo.html)
- [How to conduct an effective sprint demo (Atlassian)](https://www.atlassian.com/agile/project-management/sprint-demo)
- [Sprint Review: demonstrating value to stakeholders](https://medium.com/@noorfatimaafzalbutt/sprint-review-demonstrating-value-to-stakeholders-1a77a5f0391a)
