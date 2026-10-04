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
| 10-object | 2.769 | 2.704 | 3.303 | 3.102 |
| 10-worksheet | 2.569 | 2.726 | 3.539 | 3.622 |
| 10-class | 2.863 | 2.712 | 3.875 | 3.717 |
| 10-collection | 2.271 | 2.945 | 3.857 | 3.654 |
| 10-shadow | 1.539 | 1.601 | 2.333 | 2.488 |
| 10-other-value | 1.420 | 1.468 | 2.100 | 2.133 |
| 100-object | 2.360 | 1.907 | 3.008 | 3.431 |
| 100-worksheet | 2.838 | 2.368 | 3.614 | 3.236 |
| 100-class | 4.465 | 3.848 | 5.603 | 4.505 |
| 100-collection | 2.269 | 2.192 | 3.125 | 2.931 |
| 100-shadow | 1.416 | 1.388 | 2.120 | 1.979 |
| 100-other-value | 1.217 | 1.156 | 1.478 | 1.462 |
| 1000-object | 6.483 | 1.839 | 7.522 | 2.731 |
| 1000-worksheet | 7.075 | 2.376 | 8.788 | 4.218 |
| 1000-class | 19.699 | 15.434 | 20.015 | 15.683 |
| 1000-collection | 2.195 | 2.178 | 2.885 | 2.820 |
| 1000-shadow | 1.406 | 1.385 | 1.844 | 1.803 |
| 1000-other-value | 1.177 | 1.157 | 1.604 | 1.371 |
| 3000-object | 16.353 | 1.907 | 18.613 | 2.611 |
| 3000-worksheet | 16.804 | 2.433 | 17.422 | 3.151 |
| 3000-class | 55.939 | 42.582 | 59.992 | 43.018 |
| 3000-collection | 2.206 | 2.163 | 2.966 | 2.771 |
| 3000-shadow | 1.428 | 1.374 | 2.309 | 1.802 |
| 3000-other-value | 1.163 | 1.153 | 1.627 | 1.414 |

The 3,000-class Object and Worksheet medians improve about 8.6 and 6.9 times; Box improves from 55.939 to 42.582 ms while retaining the separate resolution cost. Small fixtures and bypass controls are mixed: 10-class Collection median increases from 2.271 to 2.945 ms, and 10-class Worksheet/shadow/Nothing medians and p95 increase slightly. The 100-class Object p95 also increases. The indexed predicate adds a lazy closure and retains examined class names for the pass. The first lookup still scans until a match or exhaustion; these measurements do not establish a universal latency improvement.

## Validation

- Type check passed.
- 24 focused tests passed, including 11 new regressions for bounded actual reads across procedures, real class/Collection mismatches, any matching duplicate class, case/full-name/kind semantics, resumed queries, metadata refresh and bypasses.
- Full suite after integration with current main: 623 files, 12,825 tests passed, 13 skipped.
- Integrated focused assignment tests: 285 passed; integrated type check and differential comparison passed again.
- 1,500 differential cases match complete analyzer and direct Set outputs, with no internal errors, covering metadata kinds/casing/duplicates/misses, host/generic/class/scalar/Variant/qualified/array targets, bare/qualified/shadowed ActiveSheet, other values, conditional declarations, branches/loops and multiple procedures.

Generated differential and counting probes remain audit scratch material. The timing benchmark and regression tests are committed.
