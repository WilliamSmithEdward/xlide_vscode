# ActiveSheet project class membership

The Set rule checks whether an unshadowed bare ActiveSheet is assigned to a Collection or any project class. Previously it scanned project metadata on every eligible assignment, including valid Object and Worksheet targets where the repeated search finds no class. Reuse the existing private incremental projectTypeNameLookup with kind class and case-insensitive names. This avoids duplicating its cursor/set algorithm.

The predicate still accepts any matching class, including duplicate names, and ignores other kinds. It compares the full metadata name after lowercasing, without trimming or stripping qualifications. Collection, shadowed ActiveSheet, qualified/non-ActiveSheet values and an existing incompatibility reason retain their short circuits. The index is shared across procedures in one Set-rule pass and discarded afterward; metadata is stable during a pass and changes are observed on later invocations.

## Deterministic cost

Actual metadata getters, outside timed runs, count all reads with 100 unrelated project classes plus Box and 100 assignments:

| Mode | Before names / kinds | After names / kinds | Diagnostics |
| --- | ---: | ---: | ---: |
| Object | 10,100 / 10,100 | 101 / 101 | 0 |
| Worksheet | 10,100 / 10,100 | 101 / 101 | 0 |
| Box class | 30,500 / 70,700 | 20,501 / 60,701 | 100 |
| Collection | 0 / 0 | 0 / 0 | 100 |
| Shadowed ActiveSheet | 0 / 0 | 0 / 0 | 0 |
| Nothing value | 0 / 0 | 0 / 0 | 0 |

Box's remaining cost includes repeated resolveKnownObjectAssignmentType filters for expected types. That is a separate audit finding; this change only removes the repeated ANY-class membership search. There are no production counters.

## Reproduction and timings

From the repository root with dependencies installed:

```powershell
node scripts/benchmark-active-sheet-project-types.mjs --baseline=ab5df8a6
node scripts/benchmark-active-sheet-project-types.mjs
```

The baseline replaces only assignments.ts from ab5df8a6 in the current harness. Fixtures use the listed number of unrelated classes plus Box and 1,000 Set writes. Object/Worksheet/Box/Collection modes vary the target type; shadow declares a local ActiveSheet, and other-value assigns Nothing. Every sample checks 1,000 diagnostics for class/Collection targets and zero for other modes.

Parsing, tokenization and binding are outside timing. Parsed bodies and tokens are reused; each sample has a fresh bound root and fresh completion caches, with a root identity assertion. These are Set-rule timings with reused parsed bodies, not cold dataflow or full analyzer/editor latency. Three warmups precede 15 samples. Before/after runs were sequential after tests and differential probes completed, with no concurrent audit workloads.

Node v24.18.0; AMD Ryzen 7 9800X3D 8-Core Processor           ; measured locally on 2026-10-03. Times in milliseconds.

| Surfaces / mode | Before median | After median | Before p95 | After p95 |
| --- | ---: | ---: | ---: | ---: |
| 10-object | 2.797 | 2.703 | 3.235 | 3.112 |
| 10-worksheet | 2.570 | 2.610 | 3.378 | 3.494 |
| 10-class | 2.678 | 2.605 | 3.772 | 3.646 |
| 10-collection | 2.285 | 2.959 | 3.918 | 3.448 |
| 10-shadow | 1.554 | 1.577 | 2.262 | 2.238 |
| 10-other-value | 1.417 | 1.428 | 2.129 | 1.769 |
| 100-object | 2.352 | 1.906 | 2.940 | 2.425 |
| 100-worksheet | 2.826 | 2.379 | 3.525 | 3.411 |
| 100-class | 4.402 | 3.860 | 8.364 | 4.465 |
| 100-collection | 2.224 | 2.216 | 2.901 | 2.894 |
| 100-shadow | 1.408 | 1.378 | 2.057 | 1.880 |
| 100-other-value | 1.197 | 1.145 | 1.417 | 1.613 |
| 1000-object | 6.453 | 1.876 | 6.940 | 2.630 |
| 1000-worksheet | 7.338 | 2.435 | 8.366 | 4.067 |
| 1000-class | 19.814 | 15.592 | 20.543 | 17.190 |
| 1000-collection | 2.172 | 2.215 | 3.087 | 2.713 |
| 1000-shadow | 1.417 | 1.362 | 1.816 | 1.781 |
| 1000-other-value | 1.183 | 1.139 | 1.443 | 1.392 |
| 3000-object | 16.140 | 1.935 | 16.934 | 2.618 |
| 3000-worksheet | 16.585 | 2.484 | 17.386 | 3.226 |
| 3000-class | 55.850 | 42.291 | 57.472 | 42.998 |
| 3000-collection | 2.218 | 2.250 | 3.063 | 2.816 |
| 3000-shadow | 1.415 | 1.385 | 2.168 | 1.802 |
| 3000-other-value | 1.187 | 1.153 | 1.545 | 1.643 |

The 3,000-class Object and Worksheet medians improve about 8.3 and 6.7 times; Box improves from 55.850 to 42.291 ms while retaining the separate resolution cost. Small fixtures and bypass controls are mixed: 10-class Collection median increases from 2.285 to 2.959 ms, 10-class Worksheet median/p95 increase slightly, and some Nothing control p95 values increase. The indexed predicate adds a lazy closure and retains examined class names for the pass. The first lookup still scans until a match or exhaustion; these measurements do not establish a universal latency improvement.

## Validation

- Type check passed.
- 24 focused tests passed, including 11 new regressions for bounded actual reads across procedures, real class/Collection mismatches, any matching duplicate class, case/full-name/kind semantics, resumed queries, metadata refresh and bypasses.
- Full suite: 622 files, 12,813 tests passed, 13 skipped.
- 1,500 differential cases match complete analyzer and direct Set outputs, with no internal errors, covering metadata kinds/casing/duplicates/misses, host/generic/class/scalar/Variant/qualified/array targets, bare/qualified/shadowed ActiveSheet, other values, conditional declarations, branches/loops and multiple procedures.

Generated differential and counting probes remain audit scratch material. The timing benchmark and regression tests are committed.
