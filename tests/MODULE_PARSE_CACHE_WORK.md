# Module parser snapshots under short lookup pressure

Completion, hover, semantic tokens and analysis reuse immutable parsed-module
snapshots. The previous eight-entry cache also admitted small helper snippets.
A large-class diagnostic observed an analysis request finding only eight short
sources in that cache and parsing all 143,377 class tokens again.

The parser now keeps separate eight-entry caches for sources below 4,096
characters and larger sources. Short lookups preserve exact hits, unchanged
prefixes and compatible body snapshots for large modules. The large-module
budget remains eight; the additional eight records hold only short sources.
Parser recovery, conditional compilation and reuse algorithms are unchanged.
Callers must keep returned ASTs immutable.

## Correctness and request work

Twelve public regressions cover LF/CRLF/CR, exact snapshot identity after twenty
short lookups, five edited revisions with complete AST equality and immutable
old snapshots, conditional branch projections, both eight-entry bounds,
independence in both directions, and the 4,095/4,096-character boundary. The
original implementation fails nine cases and passes three.

A differential on base `60d5faa9` checked 2,400 generated public edits and 160
sampled private-class edits. Original and revised parses matched an independent
raw-lexer/fresh-parser oracle for complete ASTs, directive facts and complete
symbol graphs, preserving old snapshots. It included malformed edits,
conditional branches, declarations, properties, nested blocks and all three
physical line endings. Private source and failure payloads stay local.
The original parser on current base `62d739af` was verified byte-identical.

Two interleaved complete-request trials on `62d739af` used matched original and
revised bundles with identical dependencies. Each timed indexing, project-type
context, the actual macro-name lookup and all four semantic collectors. The
pressure mode also parsed twenty small public helper procedures inside the
same timer. Each variant used forty fresh warm-up requests and fifty measured
fresh requests, plus warm-query controls. Complete symbols, visible types and
query answers matched at every measured revision.

| Pressure request | Original median / p95 / max, ms | Revised median / p95 / max, ms |
| --- | --- | --- |
| Trial 1 | 79.01 / 128.22 / 141.15 | 32.58 / 59.32 / 60.42 |
| Trial 2 | 155.68 / 262.90 / 320.22 | 56.04 / 89.92 / 143.78 |

Timings varied substantially between trials. Nearby-edit control medians were
21.19/23.20 versus 22.12/22.69 ms. Line-break-change control medians were
126.77/114.88 versus 120.56/112.18 ms, with worse revised p95/maxima in both
trials. Warm-query and control tails were mixed. The result describes this
pressure request path, not a general UI speed comparison.

A prior corrected 64-cycle diagnostic on `60d5faa9` passed both variants.
Original recorded two cold full parses, six prefix parses of 25/26 tokens and
one full analysis parse after eviction by short snippets. Revised recorded
those two cold parses and seven prefix parses of 25/26 tokens. An earlier
instrumentation attempt hooked the reference helper instead of the normal
fallback; its zero counters were invalid and were discarded.

## Integrated editor validation

On `62d739af`, compilation and type checking pass, along with all 15,843 unit
cases across 803 files (33 existing skips across seven files). All 39 editor
integration cases pass. All thirteen native cases pass in a 64-cycle run and
an unchanged 1,000-cycle repeat. The ordinary loop uses renderer keys to delete
`.cez` to `.ce`, requires a visible `Cells` suggestion, types `z` and requires
the missed-prefix menu to clear. A preceding synthetic statement changes
between cycles to defeat identical source reuse. Real mouse hovers resolve
its variable's `Long` type every sixteen cycles. The workbook is accessed
through a disposable copy.

| Operation | 64 cycles: median / p95 / max, ms | 1,000 cycles: median / p95 / max, ms |
| --- | --- | --- |
| Backspace paint | 17 / 31 / 33 | 16 / 33 / 55 |
| Suggestion menu after deletion paint | 120 / 138 / 141 | 119 / 136 / 152 |
| Typing paint | 21 / 47 / 49 | 18 / 49 / 79 |
| Missed-prefix dismissal | 1 / 10 / 12 | 1 / 11 / 41 |
| Resolved mouse hover | 410 / 411 / 411 (4 hovers) | 379 / 414 / 417 (62 hovers) |

The long repeat uses pinned VS Code 1.139.1, host/renderer CPU profiling,
fresh-source edits, navigation observations and disabled word suggestions.
The owned page is brought forward and one focused editor is required before
input sequences, hovers and fresh-source edits. Focus is not restored inside
timed polling. Foreground process IDs and executable names were sampled every
500 ms by a read-only external observer; there was no window-content capture
or foreground manipulation by the observer. Inspection commands were not
launched during the repeat. Focus setup is outside action timing but part of
the overall profiled workload. These observations are not a controlled
before/after UI comparison.

The repeat passed all 1,000 recoveries and 62 hovers. Its largest recorded host
heartbeat delay was approximately 79 ms. Five profile/timing/observation files
and the exact extension bundle/map were retained, with source and capture-time
bundle hashes, freshness, counts and all seven artifact hashes verified.
Profiles, private source and foreground metadata stay in ignored local files.

## Retained failures and limits

A prior ordinary candidate run had a focused recovery failure at cycle 722,
after 722 successes and 45 resolved hovers. That capture did not record
recovery state, so it cannot establish a cause. A separate controlled native
reproduction identified the one-second delayed-caret recovery expiry; #1208
fixes that lifetime defect. The cache was held and then revalidated with that
fix already on main. Successful diagnostic runs do not erase the old failure.

The first integrated long attempt failed menu recovery at cycle 309 after
309 successes and 19 resolved hovers. Backspace had correctly painted `.ce`,
but the renderer reported no document focus and no focused editor. Its exact
seven-artifact capture is retained. The user reported no manual screen
interaction. The unchanged repeat passed and recorded foreground process
changes during profiling; it does not identify the earlier focus loss's cause.
Earlier focus-loss captures and the invalid-hook diagnostic are also retained.

Remaining profile work includes casing, symbol construction and project
context. Inclusive categories overlap and do not establish a slow frame's
cause. This change removes confirmed parser eviction work. The broader
latency investigation remains open.
