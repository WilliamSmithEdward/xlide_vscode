# Host-global semantic position gates

`collectHostGlobalTokens` filters member accesses (`x.Name`, `x!Name`) and type
positions (`As Name`, `New Name`) from its output. It previously resolved the
name through the host globals, Global members and constants before applying
those position checks. Names that cannot produce a bare-value token now leave
before those lookups. Declaration shadowing, filtered-stream adjacency, host
resolution precedence and emitted token fields remain unchanged.

## Reproduction

```powershell
npx vitest run tests/semanticGlobalPositionWork.test.ts tests/vbaSemanticTokens.test.ts tests/editorSurfaceFollowupPerformance.test.ts
node scripts/benchmark-semantic-global-gates.mjs --baseline=453511f94a6078c227cde6db766daa45335f6176
node scripts/benchmark-semantic-global-gates.mjs
```

The work-count cases use 1,000 repeated dot/bang member lines and type positions,
with a final genuine bare `Application`. All three EOL cases fail against the
baseline because discarded positions still query host metadata, while the bare
output control passes. The fixed cases query only the genuine bare occurrence
among those names, and compare its complete expected token. A separate control
covers host globals, Global methods, constants and the existing adjacency after
comment/newline filtering.

## Measurement

Node 24.18.0, Ryzen 7 9800X3D, nine samples after three warmups, separate process
order baseline/current/current/baseline. Each sample uses a fresh source snapshot
with its parser and lexer primed. The actual global collector runs once; its
name gathering, token filtering and host lookups remain timed. Source creation,
model construction and assertions are excluded. Every result is checked against
independently constructed complete expected token arrays.

At 10,000 repeated lines:

| Shape | Baseline median ms | Fixed median ms |
| --- | ---: | ---: |
| Dot/bang members | 2.7018–2.9588 | 2.2417–2.4516 |
| `New Application` type positions | 0.9988–1.4101 | 0.8558–0.9152 |
| Genuine bare `Application` control | 0.8760–0.8928 | 0.8789–0.9017 |

This is a modest component improvement. Small-input timings are mixed and
outliers remain: fixed maxima among nine samples reached 4.6578 ms for member
lines and 5.7779 ms for type lines. This does not establish faster parsing,
consistently lower tail latency, or an end-to-end editor improvement.

## Output parity

A separate baseline/current comparison froze ASTs, lexer tokens and trivia and
checked 8,256 oracle sources plus 36 generated dot/bang/type/comment/EOL cases.
Excel/Word contexts and designer-name shadowing were alternated, with copied
source queries repeated. All 49,752 complete semantic output arrays matched;
6,615 emitted tokens had exact source spans. No shared input mutation occurred.

Types and 70 focused tests passed (one existing skipped test). Full suite passed: 739 files, 14,575 tests, 33 skipped tests, in 75.16 seconds.

Repeated declaration-name collection and filtered token allocations remain
separate, unmeasured candidates. No cache or public API was added in this fix.
