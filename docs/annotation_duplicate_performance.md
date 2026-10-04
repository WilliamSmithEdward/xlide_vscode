# Attribute annotation duplicate checks

The annotation reader previously searched every accepted annotation to decide whether each new annotation was a duplicate. Distinct valid annotated declarations consequently performed quadratic work. It now tracks accepted module kinds for one read, accepted kinds for the current procedure header, and exact variable target spellings across the module. Empty procedure pending blocks return before creating a set. All sets are local to the read; no persistent cache is introduced.

Validation remains before key insertion. Invalid arguments do not reserve a kind/target. Global DefaultMember rejection keeps its existing precedence and message. Procedure keys are scoped per header, retaining distinct property legs and their targetOccurrence values. Variable targets intentionally retain the previous exact-case comparison across declarations; the optimization does not change that policy.

## Deterministic work

Actual readAttributeAnnotations on distinct annotated procedures/variables yielded 45 duplicate candidate visits at ten declarations, 4,950 at 100 and 499,500 at 1,000. A separate stable getter observer freezes accepted annotation objects as they are emitted and counts their kind reads during the actual reader call. It captures the count before output assertions.

| Fixture | Baseline accepted-kind reads | Fixed accepted-kind reads |
| --- | ---: | ---: |
| 1,000 distinct described procedures | 499,500 | 0 |
| 1,000 distinct described variables | 499,500 | 0 |
| 1,000 described variables plus 100 module descriptions | 599,599 | 0 |

Zero means the reader no longer rereads emitted annotation kinds for duplication; it still parses the source and checks the local keys. Three new work tests fail on baseline. Three controls check exact variable spelling across declarations, invalid-argument precedence and property-leg/global DefaultMember behavior.

## Timings

Windows, Ryzen 7 9800X3D, Node v24.18.0; baseline b678a385. Four isolated processes ran baseline/fixed/fixed/baseline, with three warmups and fifteen measured rounds. Each range spans the two processes for that version, in milliseconds; p95 is the largest of fifteen samples. Very small observations are not reliable speedup estimates. All 42 complete result hashes match across all four processes.

Reader timings call the actual annotation reader. Reader-writer timings additionally call the actual writer only when accepted annotations exist, matching the save helper's guard. They exclude save IO and final change-message formatting. LF sources and expected invariants are prepared outside timing; complete output comparisons and annotation/problem/change counts are checked outside timing. No Office execution or VS Code UI latency was measured.

Procedures/variables have distinct names and one description each. Procedure-duplicates repeats a description above one procedure. Module-duplicates combines count described variables with count module descriptions. Module-only has one module description followed by count unannotated procedures; unknown-only uses an unmanaged Folder tag; no-annotations contains ordinary unannotated procedures. All have a module-name attribute header.

| Declarations/tags | Shape | Scope | Baseline median ms | Fixed median ms | Baseline p95 ms | Fixed p95 ms |
| ---: | --- | --- | ---: | ---: | ---: | ---: |
| 1 | procedures | reader | 0.003–0.005 | 0.004–0.005 | 0.010–0.012 | 0.013–0.014 |
| 1 | procedures | reader-writer | 0.007–0.010 | 0.009–0.009 | 0.023–0.024 | 0.017–0.018 |
| 1 | variables | reader | 0.003–0.003 | 0.002–0.002 | 0.008–0.008 | 0.005–0.005 |
| 1 | variables | reader-writer | 0.003–0.006 | 0.003–0.003 | 0.008–0.009 | 0.004–0.005 |
| 1 | procedure-duplicates | reader | 0.002–0.004 | 0.002–0.002 | 0.005–0.007 | 0.004–0.011 |
| 1 | procedure-duplicates | reader-writer | 0.006–0.007 | 0.004–0.004 | 0.012–0.030 | 0.006–0.008 |
| 1 | module-duplicates | reader | 0.003–0.004 | 0.003–0.003 | 0.005–0.007 | 0.006–0.009 |
| 1 | module-duplicates | reader-writer | 0.006–0.009 | 0.005–0.006 | 0.011–0.016 | 0.011–0.011 |
| 1 | module-only | reader | 0.002–0.002 | 0.002–0.002 | 0.004–0.007 | 0.002–0.003 |
| 1 | module-only | reader-writer | 0.005–0.005 | 0.003–0.003 | 0.012–0.023 | 0.008–0.009 |
| 1 | unknown-only | reader | 0.002–0.002 | 0.001–0.001 | 0.004–0.010 | 0.003–0.005 |
| 1 | unknown-only | reader-writer | 0.002–0.002 | 0.002–0.002 | 0.004–0.012 | 0.004–0.004 |
| 1 | no-annotations | reader | 0.000–0.000 | 0.000–0.000 | 0.000–0.000 | 0.000–0.003 |
| 1 | no-annotations | reader-writer | 0.000–0.000 | 0.000–0.000 | 0.000–0.000 | 0.000–0.000 |
| 100 | procedures | reader | 0.153–0.193 | 0.086–0.087 | 0.289–0.566 | 0.245–0.258 |
| 100 | procedures | reader-writer | 0.263–0.270 | 0.141–0.148 | 0.439–0.549 | 0.387–0.416 |
| 100 | variables | reader | 0.104–0.177 | 0.084–0.091 | 0.263–0.274 | 0.183–0.212 |
| 100 | variables | reader-writer | 0.160–0.205 | 0.101–0.103 | 0.294–0.324 | 0.169–0.216 |
| 100 | procedure-duplicates | reader | 0.023–0.048 | 0.029–0.029 | 0.075–0.177 | 0.092–0.127 |
| 100 | procedure-duplicates | reader-writer | 0.024–0.047 | 0.029–0.029 | 0.105–0.180 | 0.090–0.145 |
| 100 | module-duplicates | reader | 0.175–0.237 | 0.104–0.109 | 0.308–0.368 | 0.251–0.288 |
| 100 | module-duplicates | reader-writer | 0.203–0.326 | 0.128–0.133 | 0.384–0.455 | 0.308–0.470 |
| 100 | module-only | reader | 0.022–0.033 | 0.021–0.022 | 0.033–0.046 | 0.076–0.088 |
| 100 | module-only | reader-writer | 0.028–0.042 | 0.039–0.039 | 0.058–0.109 | 0.047–0.059 |
| 100 | unknown-only | reader | 0.032–0.038 | 0.019–0.020 | 0.058–0.228 | 0.026–0.128 |
| 100 | unknown-only | reader-writer | 0.040–0.047 | 0.019–0.020 | 0.041–0.182 | 0.057–0.085 |
| 100 | no-annotations | reader | 0.001–0.001 | 0.000–0.001 | 0.001–0.001 | 0.001–0.002 |
| 100 | no-annotations | reader-writer | 0.001–0.001 | 0.001–0.001 | 0.001–0.001 | 0.001–0.002 |
| 1000 | procedures | reader | 1.857–1.925 | 0.541–0.592 | 2.291–3.621 | 0.980–1.284 |
| 1000 | procedures | reader-writer | 2.349–2.390 | 1.017–1.061 | 2.739–2.816 | 1.239–1.657 |
| 1000 | variables | reader | 1.575–2.349 | 0.442–0.452 | 1.865–3.275 | 0.658–0.865 |
| 1000 | variables | reader-writer | 2.055–2.147 | 1.126–1.299 | 2.421–3.813 | 1.501–1.709 |
| 1000 | procedure-duplicates | reader | 0.169–0.174 | 0.163–0.196 | 0.295–0.400 | 0.323–0.390 |
| 1000 | procedure-duplicates | reader-writer | 0.208–0.215 | 0.200–0.210 | 0.216–0.462 | 0.341–0.370 |
| 1000 | module-duplicates | reader | 2.810–2.913 | 0.608–0.628 | 3.252–3.483 | 0.847–0.883 |
| 1000 | module-duplicates | reader-writer | 3.231–3.362 | 1.046–1.201 | 4.082–4.208 | 1.404–1.599 |
| 1000 | module-only | reader | 0.181–0.189 | 0.174–0.176 | 0.187–0.313 | 0.180–0.198 |
| 1000 | module-only | reader-writer | 0.235–0.245 | 0.228–0.229 | 0.240–0.355 | 0.261–0.311 |
| 1000 | unknown-only | reader | 0.180–0.188 | 0.173–0.290 | 0.182–0.268 | 0.177–0.472 |
| 1000 | unknown-only | reader-writer | 0.181–0.247 | 0.175–0.178 | 0.365–0.383 | 0.203–0.279 |
| 1000 | no-annotations | reader | 0.004–0.004 | 0.004–0.007 | 0.004–0.004 | 0.004–0.014 |
| 1000 | no-annotations | reader-writer | 0.004–0.004 | 0.004–0.007 | 0.004–0.004 | 0.007–0.009 |

For 1,000 described procedures, reader medians improve from 1.857–1.925 ms to 0.541–0.592 ms, and reader-writer from 2.349–2.390 ms to 1.017–1.061 ms. The 1,000-variable reader improves from 1.575–2.349 ms to 0.442–0.452 ms. Already-cheap duplicate-only and some small/unknown/no-annotation controls are mixed or slower: the 1,000 unknown-only reader spans 0.180–0.188 ms baseline and 0.173–0.290 ms fixed; the unchanged no-annotation guard spans 0.004 ms baseline and 0.004–0.007 ms fixed. This is a targeted reader/save-stage improvement, not a universal analyzer or save-time speedup.

Reproduce from the repository root:

```powershell
node scripts/benchmark-annotation-duplicates.mjs --baseline=b678a385
node scripts/benchmark-annotation-duplicates.mjs
node scripts/benchmark-annotation-duplicates.mjs
node scripts/benchmark-annotation-duplicates.mjs --baseline=b678a385
npm exec vitest run tests/annotationDuplicateWork.test.ts tests/attributeAnnotations.test.ts tests/attributeRewriteTargetIndex.test.ts tests/diagnostics/classAttributes256.test.ts
npm run check-types
npm test -- --run
```

## Validation

Six new tests (three baseline work failures), 48 focused tests and type checking pass. After rebasing onto mainb678a385, the full suite passes 657 files, 13,296 tests, with 13 skipped. Only observation comments, benchmark formatting and documentation changed afterward.

An ad hoc baseline/fixed check compares complete reader and writer results for all 8,256 corpus sources and 7,056 generated cases, 15,312 comparisons total. Generated cases combine fourteen valid/invalid/unknown annotation spellings, four pairs of variable/procedure/property/array declarations, three header contexts and LF/CRLF/CR. They repeat declarations and vary name case, checking annotations, problems, occurrences, rewritten text, changes and skips. Reader results, annotation/problem objects and their arrays are frozen before passing them to the writer; all comparisons match with zero uncaught exceptions. This preserves existing newline behavior and does not claim frozen tokens/ASTs or complete repository coverage.
