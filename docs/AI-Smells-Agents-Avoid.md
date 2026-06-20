# AI Smells: Agents Please Avoid These

A field guide of writing patterns that read as machine-generated. Feed this file to an
LLM or agent as a style constraint. The goal is **prose that carries information cleanly**,
not prose that fools a detector. Those are different targets, and only the first one is
worth optimizing for.

---

## How to use this file

If you are an agent or model reading this: treat every entry below as a default-off
constraint on your own output. Before returning text, scan it for these patterns and
remove them **unless** a specific, defensible reason makes the pattern the right call for
this audience. The reasons matter. A blanket ban produces its own tell (see *Overcorrection*
below), so apply judgment, not a filter.

If you are a human configuring an agent: paste this into a system prompt, a `CLAUDE.md` /
`AGENTS.md`, a style guide, or a context file. The **Quick Reference** table at the end is
the compressed version if you are tight on tokens.

---

## Inclusion criteria (how each entry earned its place)

Every pattern below cleared three gates:

1. **Credible origin.** It traces to a peer-reviewed study, a major publication, or an
   established institution — not a single blog.
2. **Independent corroboration.** It is reported by **three or more independent sources**.
3. **Internal review.** It survived a sanity check for whether the pattern is *actually*
   AI-correlated versus just common bad writing, and whether avoiding it improves the text.

Each entry carries an **evidence tier** so you can weight it honestly:

- **Tier A — Peer-reviewed.** Backed by a published study or large empirical corpus analysis, plus independent corroboration.
- **Tier B — Documented.** Reported by major journalism and/or established institutions, corroborated across three or more independent outlets.
- **Tier C — Convergent.** Consistently and independently identified by three or more practitioner/editor sources, but not yet the subject of a dedicated peer-reviewed study.

Tier C is not weak — convergent expert observation is real evidence — but it has not been
measured the way Tier A has. Calibrate accordingly.

### Three honest caveats before the list

- **None of these prove AI authorship.** Humans use em dashes, write in threes, and say
  "moreover." The strongest framing is statistical: AI produces these patterns at a
  *frequency and density* humans rarely match. Avoid them to write better, not to pass as human.
- **Detection-gaming is a treadmill.** Tools that flag AI text are unreliable and change
  constantly; vendors already sell "humanizers" to defeat them. Optimizing for the detector
  is a losing game. Optimize for the reader.
- **Overcorrection is its own smell.** Text that has clearly been scrubbed of every em dash,
  every "delve," and every third item reads as processed in a different way. The fix for
  robotic writing is specificity and real content, not a find-and-replace pass.

---

## 1. Characters and punctuation

### 1.1 Non-ASCII punctuation with no reason for it
**Rule:** Default to ASCII. Do not emit em dashes (—), en dashes (–), curly/smart quotes
(" " ' '), the ellipsis character (…), or non-breaking spaces unless the output target
specifically calls for typeset punctuation (e.g. a print-bound document or a house style
that requires it).

**Why it smells:** The em dash became the single most talked-about visual tell of AI text,
nicknamed the "ChatGPT hyphen," because models deploy it far more often than typical writers
and drop it where a comma, colon, or parentheses would be more natural. The pattern got
prominent enough that OpenAI shipped a setting to suppress it. The deeper issue is generic:
plain ASCII keyboards do not produce these glyphs, so their unprompted appearance signals an
intermediary processed the text.

**Instead:** Use a hyphen or restructure. Use straight quotes. Type three periods only when
you actually need an ellipsis, and rarely.

**Example:**
- Smells: `The result was clear — and it changed everything.`
- Better: `The result was clear, and it changed everything.`

**Evidence — Tier B:** *The Washington Post*, *Rolling Stone*, and *TechCrunch* all covered
the em-dash-as-AI-tell phenomenon; OpenAI publicly acknowledged it by adding a control.

### 1.2 Decorative emoji, especially in headers and bullets
**Rule:** No emoji as decoration. No 🚀 / ✨ / 📊 leading a heading or a list item. Use emoji
only if the user uses them first or explicitly asks.

**Why it smells:** Emoji-prefixed section headers and bullet points are a recognized signature
of AI-generated READMEs, documentation, and posts. Almost nobody hand-formats running prose
this way.

**Instead:** Let the words carry the structure.

**Evidence — Tier C:** Convergent across AI-writing-pattern catalogs (tropes.fyi, the
community "AI Tropes" lists) and the Wikipedia *Signs of AI writing* guide's markdown section.

---

## 2. Vocabulary and diction

### 2.1 The "AI style" word cluster
**Rule:** Do not reach by default for: *delve, underscore, intricate, intricacies, realm,
tapestry, testament, nuanced, multifaceted, meticulous, commendable, garner, boast, leverage
(as a verb), robust, pivotal, crucial, vital, foster, navigate (figurative), landscape
(figurative), showcase, underpin, notably, comprehensive.* Use the plain word the sentence
actually needs.

**Why it smells:** This is the **most empirically grounded** entry here. A large-scale analysis
of over 15 million PubMed abstracts (2010–2024) found an abrupt, unprecedented spike in exactly
this class of "style words" after ChatGPT's release, and used their frequency to estimate that
a sizeable fraction of recent abstracts were LLM-processed. "Delve" and "underscore" are
specifically named. A separate peer-reviewed review aggregated earlier word lists
(*commendable, meticulous, intricate, realm*) and confirmed the trend.

**Instead:** Prefer the concrete verb or noun. "Examine," "show," "complex," "area," "thorough,"
"important," "use," "strong," "key."

**Example:**
- Smells: `This paper delves into the intricate realm of supply-chain robustness.`
- Better: `This paper examines how supply chains stay reliable under stress.`

**Note on origin:** A widely cited theory links "delve" to RLHF annotation outsourced to
regions (notably Nigeria) where the word is more common in formal English. Treat that as a
plausible explanation, not settled fact — but the *overuse itself* is well measured regardless
of cause.

**Evidence — Tier A:** Kobak, González-Márquez, Horvát & Lause, *Science Advances* (2025;
arXiv:2406.07016), corroborated by a systematic review in *Perspectives on Medical Education*
(2025) and the analysis in arXiv:2412.11385 ("Why Does ChatGPT 'Delve' So Much?").

### 2.2 Promotional inflation / peacock language
**Rule:** State facts at their actual weight. Do not inflate significance with: *stands as a
testament to, rich (cultural) heritage, enduring legacy, plays a significant/vital role,
a pivotal moment in, breathtaking, must-see, renowned, celebrated.*

**Why it smells:** The most consistent observation from Wikipedia's editors is that AI text
reads like a brochure — it adds grand claims about importance and legacy where a plain
statement of fact belongs. A small regional office gets described as "a pivotal moment in the
evolution of regional statistics."

**Instead:** Drop the editorializing. Give the fact and let the reader weigh it.

**Example:**
- Smells: `The festival stands as a testament to the region's rich cultural heritage.`
- Better: `The festival has run annually since 1981 and draws about 4,000 attendees.`

**Evidence — Tier B:** Wikipedia *Signs of AI writing* (WikiProject AI Cleanup), covered by
*NPR* and Digital Watch Observatory; corroborated by the ETBI library guide.

### 2.3 Marketing clichés and "journey" framing
**Rule:** Avoid: *unlock the potential, unleash the power, harness the power, in today's
fast-paced world, in the digital age, game-changer, revolutionize the way, embark on a journey,
take a deep dive, push the boundaries, at the forefront of, a gateway to, bridge the gap.*

**Why it smells:** These canned phrases recur across AI output because they were over-represented
in the SEO-heavy training data. They signal nothing and pad length.

**Instead:** Say the specific thing. If you can't replace the phrase with a concrete claim, cut it.

**Evidence — Tier C:** Convergent across multiple practitioner word-lists and editor guides
(Twixify, fomo.ai, heybex, and similar), which independently catalog the same expressions.

---

## 3. Sentence structures

### 3.1 Negation/antithesis: "It's not X, it's Y"
**Rule:** Do not manufacture contrast with: *"It's not X — it's Y," "isn't X, it's Y," "not
just X, but Y," "not only X but also Y,"* or the cross-sentence variant *"The question isn't
X. The question is Y."* One such construction in a long piece can land. Several is the tell.

**Why it smells:** This negative-parallelism reframe is the **single most frequently named
structural AI tell**. It fabricates false profundity by presenting an ordinary point as a
surprising correction to a misconception nobody held. At scale, humans did not write this way.

**Instead:** Make the positive claim directly.

**Example:**
- Smells: `This isn't just a tool — it's a whole new way of thinking.`
- Better: `This tool changes how the team triages incoming tickets.`

**Evidence — Tier B/C:** Wikipedia *Signs of AI writing* documents negative parallelism;
corroborated by the ETBI library guide and multiple independent pattern catalogs that rank it
the most reliable single tell.

### 3.2 Rule-of-three / tricolon overuse
**Rule:** Do not default to three parallel items. Watch for back-to-back tricolons and for a
third item that only completes a rhythm without adding information.

**Why it smells:** Three is the statistical average for parallel structure in the training data,
so models reach for it on autopilot. A single tricolon is a rhetorical tool; stacked tricolons,
or a filler third beat ("speed, efficiency, and innovation"), are a pattern-completion artifact.

**Instead:** Use two items when two is the truth. Use four. Break the parallelism deliberately
if you do use three. Earn the third item with real content.

**Example:**
- Smells: `It's faster, cleaner, and more scalable.`
- Better: `It's faster, and it stops corrupting the cache on retry.`

**Evidence — Tier C:** Independently identified by GPTZero (an AI-detection vendor), the
Wikipedia guide, and multiple pattern catalogs; grounded in established rhetoric (the tricolon
is classical, which is exactly why the model overuses it).

### 3.3 Trailing "-ing" significance clauses
**Rule:** Do not bolt a present-participle clause onto a sentence to gesture at meaning:
*"…, highlighting its importance," "…, reflecting broader trends," "…, underscoring the need
for innovation," "…, cementing its legacy," "…, paving the way for…"*

**Why it smells:** AI habitually ends a factual statement with a vague interpretive tail that
adds no information and often invents an opinion. It is filler dressed as analysis.

**Instead:** End the sentence at the fact. If the implication matters, state it as its own
claim with actual substance.

**Example:**
- Smells: `The store added contactless payment, improving convenience for shoppers.`
- Better: `The store added contactless payment in March.`

**Evidence — Tier B:** Wikipedia *Signs of AI writing* and the ETBI library guide both name the
trailing-participle "superficial analysis" pattern; corroborated by independent catalogs.

### 3.4 Self-answered rhetorical questions
**Rule:** Do not pose a question and answer it in the next breath as a structural device:
*"So what does this mean? It means…," "Why does this matter? Because…"*

**Why it smells:** It is a canned engagement scaffold. Used once it can work; used as a
recurring rhythm it reads as generated filler.

**Instead:** Just make the statement.

**Evidence — Tier C:** Convergent across independent AI-writing-pattern catalogs.

### 3.5 False-suspense transitions
**Rule:** Avoid manufactured buildup: *"Here's the thing," "Here's where it gets interesting,"
"Here's what most people miss," "But here's the kicker."*

**Why it smells:** These promise a revelation and then deliver an unremarkable point. The drama
is artificial.

**Instead:** Lead with the point. If it's interesting, it won't need the drumroll.

**Evidence — Tier C:** Convergent across independent pattern catalogs (tropes.fyi and the
community "AI Tropes" lists).

---

## 4. Connectives and filler

### 4.1 Transition-word pileup
**Rule:** Do not open consecutive sentences or paragraphs with *Moreover, Furthermore,
Additionally, In addition, Consequently, Notably, Importantly, Interestingly, Ultimately.*
These are valid words; the smell is the frequency and the mechanical cycling through them.

**Why it smells:** AI leans on a small repertoire of connectives like crutches, which makes
paragraphs feel assembled rather than written. The words appear far more often in model output
than in natural prose.

**Instead:** Let logical flow carry the reader. Most "Furthermore" sentences read fine with the
connective deleted.

**Evidence — Tier B/C:** Wikipedia *Signs of AI writing* (repetitive transitions), reported in
*NPR*'s coverage; corroborated by multiple independent editor guides.

### 4.2 Hedging and throat-clearing
**Rule:** Cut: *"It's important to note that," "It's worth mentioning that," "It should be
noted that," "Generally speaking," "To some extent," "In many ways," "It's worth considering."*

**Why it smells:** Safety-tuned models hedge by default and pre-pad statements with importance
framing. The phrase "It's important to note that" recurring across a piece is a well-known
indicator. The hedge usually adds nothing the following clause didn't already carry.

**Instead:** Delete the preamble and state the thing. "It's important to note that X" → "X."

**Example:**
- Smells: `It's important to note that the migration must run during off-peak hours.`
- Better: `Run the migration during off-peak hours.`

**Evidence — Tier C:** Convergent and heavily cross-referenced across independent editor and
practitioner sources (heybex, The Augmented Educator's "ten telltale signs," REM Web Solutions,
and others), each independently flagging "it's important to note."

---

## 5. Stance and attribution

### 5.1 Vague attribution / weasel words
**Rule:** Do not attribute claims to phantom authorities: *"studies show," "experts say,"
"research suggests," "some critics argue," "observers have noted," "it is widely believed,"
"industry reports indicate."* Either cite a real, specific source or make the claim in your
own voice and own it.

**Why it smells:** AI reaches for unsourced authorities to lend the appearance of evidence when
none is behind it — the "weasel wording" pattern. It often pairs one vague attribution with a
single cherry-picked source, or none.

**Instead:** Name the source and year, or state the claim plainly as your assessment. If you
cannot source it, say that you are uncertain.

**Example:**
- Smells: `Experts agree that this approach scales better.`
- Better: `In their 2023 benchmark, Patel et al. found this approach scaled to ~8x throughput.` (only if true and verifiable)

**Evidence — Tier B:** Wikipedia *Signs of AI writing* (weasel wording), corroborated by the
ETBI library guide and independent catalogs.

### 5.2 Sycophancy and conversational filler aimed at the user
**Rule:** Drop: *"Great question!," "That's a great point," "Certainly!," "Absolutely!,"
"I'd be happy to help," "I hope this helps!," "Feel free to reach out."* Do not open with
flattery or close with a customer-service sign-off unless the context genuinely calls for it.

**Why it smells:** The reflexive eagerness-to-please tone is a direct artifact of feedback
tuning, and leftover assistant pleasantries (especially when they survive into a published
document) are an obvious giveaway. The flattery is also empty: it rates the question instead
of answering it.

**Instead:** Answer directly. Respect the reader's time by skipping the warm-up and the wind-down.

**Evidence — Tier B/C:** Wikipedia's WikiProject AI Cleanup flags leftover "communication
intended for the user" as a high-signal tell; the obsequious-tone observation is widely
corroborated in independent commentary on AI register.

---

## 6. Structure and formatting

### 6.1 Bolded lead-in list items and excessive boldface
**Rule:** Do not format lists so that every item is **a bolded phrase followed by a colon and a
sentence that restates the bold phrase.** Do not bold key terms, product names, and phrases
throughout running prose. Bold sparingly and only for genuine emphasis.

**Why it smells:** The bold-lead-in bullet is a specific ChatGPT/Claude markdown signature that
barely exists in handwritten text. Pervasive boldface reads as mechanical highlighting rather
than editorial judgment. (Note: a *reference document like this one* legitimately uses some
structure; the smell is when this formatting saturates ordinary prose or a blog post.)

**Instead:** Write list items as plain phrases or sentences. Reserve bold for the rare word that
truly needs the weight.

**Evidence — Tier B/C:** Wikipedia *Signs of AI writing* (excessive boldface, markdown
artifacts), reported in coverage of the guide; corroborated by independent catalogs that
specifically name the bolded-lead-in bullet.

### 6.2 Formulaic openings and conclusions
**Rule:** Do not open with *"In today's fast-paced world…"* or *"In this article, we will…"*.
Do not close with *"In conclusion," "In summary," "To sum up,"* or a wrap-up paragraph that
zooms out to "broader implications" the body never earned.

**Why it smells:** Explicit conclusion-announcing and boilerplate openers are template scaffolding.
The reflexive "broader implications" ending inflates a small point into cosmic significance,
which is the promotional drift from §2.2 in structural form.

**Instead:** Start with the actual first point. End when you've said the last useful thing. A
genuine conclusion can exist; it just shouldn't announce itself or manufacture grandeur.

**Evidence — Tier C:** Convergent across independent pattern catalogs and editor guides.

---

## 7. Content shape

### 7.1 Genericism / restating the obvious
**Rule:** Do not pad with sentences that state what the reader already knows
("Punctuality is an important habit. To be on time, plan for delays."). Maximize information
per sentence. If a sentence would survive being deleted with no loss, delete it.

**Why it smells:** Predictive generation gravitates to the most probable, least specific
continuation, producing fluent text with near-zero information density — the "Land of the
Obvious." Editors describe it as a hollowed-out, vanilla quality.

**Instead:** Be specific. Use concrete numbers, names, mechanisms, and examples. Specificity is
the most reliable single antidote to AI-flavored prose.

**Evidence — Tier C:** Convergent across professional-editor accounts of spotting AI text,
independently describing the same low-density genericism.

### 7.2 Uniform rhythm (no burstiness)
**Rule:** Vary sentence length and structure deliberately. Do not produce paragraph after
paragraph of similarly sized, similarly shaped sentences.

**Why it smells:** Human writing moves in bursts — a long sentence, then a short one — while AI
output tends toward an even, metronomic rhythm. The flatness is part of what makes generated
text feel bland even when it is grammatically clean.

**Instead:** Let some sentences run. Cut others to three words. Read it aloud; if the cadence is
flat, break it.

**Evidence — Tier C:** Convergent across independent editor accounts contrasting human
"burstiness" with the uniform cadence of model output. (Related to "perplexity/burstiness"
heuristics used, imperfectly, by detection tools — flagged here as a writing target, not as a
reliable detector.)

---

## Overcorrection: the anti-pattern

Do not solve this list with a find-and-replace. Output that has been mechanically stripped of
every em dash, every "delve," and every group of three reads as *processed* — which is its own
tell, and worse, it usually means the writer optimized for surface features instead of fixing
the substance. The real cure for AI-flavored writing is **specific content, committed claims,
and a real point of view.** Once those are present, an occasional em dash or tricolon is
invisible, because it is doing actual work. Use this guide to raise quality. Do not use it to
sand text into a different kind of uniformity.

---

## Quick Reference

| Category | Avoid | Prefer |
|---|---|---|
| Punctuation | Em/en dash, curly quotes, …, emoji decoration | Plain ASCII; hyphen, comma, straight quotes |
| Diction | delve, underscore, intricate, realm, tapestry, leverage, robust, pivotal, foster | The plain word the sentence needs |
| Promotion | "stands as a testament," "rich heritage," "pivotal moment" | State the fact at its real weight |
| Clichés | "unlock the potential," "in today's fast-paced world," "game-changer," "deep dive" | The specific claim, or nothing |
| Antithesis | "It's not X, it's Y," "not only X but Y," "isn't X, is Y" | The direct positive statement |
| Rule of three | Reflexive triads, filler third item | Two items, or four, or a broken pattern |
| -ing tails | "…, highlighting its importance," "…, paving the way" | End at the fact |
| Transitions | Moreover / Furthermore / Additionally / Notably pileup | Logical flow; delete the connective |
| Hedging | "It's important to note that," "It's worth mentioning" | Delete the preamble; state it |
| Attribution | "studies show," "experts say," "some critics argue" | A real, specific citation or your own owned claim |
| Sycophancy | "Great question!," "Certainly!," "I hope this helps!" | Answer directly |
| Formatting | Bolded-lead-in bullets, pervasive **bold**, emoji headers | Plain items; bold only for true emphasis |
| Openings/endings | "In today's world…," "In conclusion," forced "broader implications" | Start at the point; stop when done |
| Content | Restating the obvious, low density | Numbers, names, mechanisms, specifics |
| Rhythm | Uniform sentence length | Deliberate variation; read aloud |

---

## Sources

Primary and corroborating sources behind the entries above. Tiers reflect the strength of
evidence, not the credibility of any single source.

**Peer-reviewed / empirical (Tier A backbone):**
- Kobak, D., González-Márquez, R., Horvát, E.-Á., & Lause, J. — *Delving into LLM-assisted writing in biomedical publications through excess vocabulary.* Science Advances (2025); arXiv:2406.07016; PubMed: 40601754.
- *Delving into PubMed Records: How AI-Influenced Vocabulary has Transformed Medical Writing since ChatGPT.* Perspectives on Medical Education (2025), pmejournal.org/articles/10.5334/pme.1929 (systematic review aggregating LLM-associated word lists).
- *Why Does ChatGPT "Delve" So Much? Exploring the Sources of Lexical Overrepresentation in LLMs.* arXiv:2412.11385.

**Institutional / community field guides (Tier B backbone):**
- Wikipedia — *Signs of AI writing* (WikiProject AI Cleanup), en.wikipedia.org/wiki/Wikipedia:Signs_of_AI_writing, and the associated WikiProject AI Cleanup Guide.
- Education and Training Boards Ireland (ETBI) Digital Library — *Signs of AI Writing*, library.etbi.ie (independent institutional reproduction/curation).

**Major journalism (Tier B corroboration):**
- *NPR* — Wikipedia editors publish a guide to detect AI-written entries (Sept 2025).
- *The Washington Post* — coverage of the em dash as a debated AI tell (Apr 2025).
- *Rolling Stone* — "ChatGPT Hyphen": em dashes and AI writing (Apr 2025).
- *TechCrunch* — OpenAI adds control to suppress em dashes (Nov 2025).
- Digital Watch Observatory — coverage of the Wikipedia field guide (Sept 2025).

**Detection vendor and practitioner catalogs (Tier C corroboration):**
- GPTZero — on the rule of three as a stylistic tell.
- tropes.fyi and the community "AI Tropes" lists — directory of AI writing patterns (intended for inclusion in system prompts).
- Independent editor/practitioner guides (heybex; The Augmented Educator, "ten telltale signs"; REM Web Solutions; Twixify; fomo.ai; Sean Kernan's "13 signs") — converging, independently compiled lists of overused phrases and structures.

**Origin theory (context, not a heuristic):**
- Alex Hern (*The Guardian*) and Simon Willison — the RLHF/Nigerian-English hypothesis for "delve." Treat as plausible, unconfirmed explanation; the overuse itself is independently measured (Tier A).

---

*Last compiled June 2026. AI writing patterns shift as models update (for instance, em-dash
behavior changed once vendors responded to the discourse). Re-verify periodically. This guide
targets writing quality and reader respect — not evasion of detection tools, which are
unreliable and not worth optimizing against.*
