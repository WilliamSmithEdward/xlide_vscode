# Function and type facts: adopting equal caller source

`knownFunctionResults` and `moduleTypes` each retain one fact record per weakly
owned parsed module. A hit requires matching source text and conditional activity
identity. On an equal-text hit, retaining the old string could require a complete
source comparison on every following lookup. Both caches now adopt the caller
string after all existing keys match. Fact contents, result identity, invalidation
and weak ownership are unchanged. The function call-evaluation context still uses
the source content associated with those facts; equal source content is unchanged.

## Reproduction and measured scope

Run `node scripts/benchmark-module-fact-source.mjs`, or add
`--baseline=73023d76e7e0f240abd0f77ea18f97fe320deebc` for the previous behavior.
Both consumers are tested at approximately 1 KB, 100 KB and 1 MB, using original
strings and equal copies. Three warmups precede nine measured rounds. Each round
uses fresh source and a fresh parsed module. The script independently asserts
complete expected facts before timing and checks result identity on every hit.
Construction, parsing and initial fact computation are excluded from timing.

Paired before/after/after/before runs on Node 24.18.0 / Ryzen 7 9800X3D:

| 1 MB source, 1,000 warm hits | Before median ms | After median ms |
| --- | --- | --- |
| Equal copy, function results | 16.6609–17.6903 | 0.0546–0.0560 |
| Equal copy, type fields | 16.6852–16.7918 | 0.0538–0.0560 |
| Same string, function results | 0.0167–0.0168 | 0.0172 |
| Same string, type fields | 0.0168–0.0169 | 0.0169–0.0170 |

The same-string function control adds about 0.4–0.5 ns per lookup in these runs.
The first equal-copy comparison is still necessary. Supplying a new copy on every
lookup may still require full comparisons. These are component measurements;
they do not establish whole-editor or full-diagnostics latency improvements.

## Validation

The linked PR records type checking, focused and full-suite results, and baseline
differential evidence. The differential checks complete fact maps and diagnostic
arrays on frozen syntax trees and tokens, including LF/CRLF/CR, active/inactive
conditional declarations, function result types, fixed array bounds, result
identity, and source replacement under a retained module node.

Type checking and all 16 focused tests pass. The baseline differential includes
8,256 corpus sources and 18 generated cases: 13,434 diagnostics, no internal
errors, and identical complete fact maps and diagnostic arrays. Source A/B/A
checks also independently verify function literals and array upper bounds.

The full suite passes: 733 files and 14,503 tests (7 files and 33 tests skipped).
