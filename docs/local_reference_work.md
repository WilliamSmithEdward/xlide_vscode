# Local reference work

Document highlights use the same collector as references and rename. A fresh
local-variable query previously built class member surfaces with diagnostic
value facts, stripped every line of the module, and requested its complete
lexer stream to classify reads and writes. The next highlight could therefore
consume tens of milliseconds after a keystroke in a large class.

The collector now resolves the scope before translating declarations through
member surfaces. Parameters, ordinary local variables and procedure constants
scan and classify only their owning procedure. Each occurrence still resolves
against the project index at its original absolute offset, so shadows and
ambiguous bindings retain the existing rules. Function and Property result
variables keep the complete path: they share a callable declaration whose
qualified references can occur in other modules.

Reference-only member lookups omit diagnostic class value facts. The exported
declaration lookup preserves its default metadata for callers that need it.
Whole-source occurrence scans retain their existing cache; procedure slices
opt out, avoiding a long-lived cache key that could retain a much larger parent
string. The lexer cache remains capped by entry count.

This bounds stripping and lexing for eligible locals. It does not remove all
module work: absolute line starts and project scope metadata are still needed.
Member targets, unresolved names and invalid/stale procedure spans retain the
existing fallback path. Integration UI timings include VS Code scheduling,
IPC and painting; the ordinary menu and hover delays are separate from the
pure collector benchmark.

## Regression coverage

- `referenceLocalScopeWork.test.ts` checks complete read/write results on six
  fresh large-class revisions, bounds lexer/split inputs below 300 characters,
  and proves zero member-surface/value-fact requests on the local path.
- The same test checks an unmatched qualified member through the whole
  collector, preserves the default declaration metadata, keeps cross-module
  Function result references, and verifies an indented Unicode parameter with
  mixed physical line endings.
- `identifierOccurrenceRange.test.ts` verifies default cache reuse and both
  single-name/multiple-name ephemeral scans, preserving continued comment
  context and absolute positions.
- `editorSurfaces.test.ts` checks actual highlight declaration/read/write ranges
  after six edits in a 700-method class and rejects semantic read/write highlights
  for an unmatched member (VS Code can still supply its generic Text fallback). The native large-class harness checks physical keyboard
  recovery and mouse hovers with fresh source revisions.

The private workload and all exported source, bundles, captures and observations
stay local and ignored. Reproducible public regressions contain synthetic code.

## Controlled collector measurements (2026-10-04)

The two bundles share the same project/index dependencies. The before bundle
uses the frozen original reference and source-scan files; the after bundle uses
the production files in this change. A 969,470-character class is revised with
a fixed-width nonce comment. Each trial takes 15 warmups, 30 warm queries and
30 fresh queries; the second trial reverses variant order. Index updates occur
before and outside the query timer. The local fixture must return exactly its
three declaration/write/read references; the unmatched member must stay empty.

| Complete query | Before median, two trials | After median, two trials |
| --- | ---: | ---: |
| Fresh local | 30.85 / 31.53 ms | 1.66 / 1.60 ms |
| Warm local | 4.66 / 4.88 ms | 1.42 / 1.43 ms |
| Fresh unmatched member | 11.63 / 11.23 ms | 5.40 / 4.98 ms |
| Warm unmatched member | 0.17 / 0.16 ms | 0.17 / 0.17 ms |

Fresh local p95 was 32.40 / 34.78 ms before and 1.79 / 1.88 ms after;
maxima were 36.56 / 34.97 ms before and 4.19 / 5.14 ms after. Unmatched-member
p95 improved from 19.46 / 15.32 ms to 8.88 / 11.07 ms, but its second after
trial still had a 21.31 ms outlier, exceeding that trial's 15.58 ms before
maximum. Index-update medians were approximately 4–5 ms and are excluded.

Frozen differential checks agreed on 5,760 complete public reference results
and 160 sampled private results, covering scopes, callable result variables,
member receivers, continuations, strings/comments, conditional declarations,
Unicode names and LF/CRLF/CR line endings. Compilation and 15,169 unit tests
passed, with 33 existing opt-in tests skipped. These measurements establish a
collector improvement, not a promise about every UI outlier.
