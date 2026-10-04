# Set object-type resolution

The Set rule resolved the expected object type for its eligibility gate, then again for compatibility; a known actual type could cause a third resolution. Every project resolution filtered all metadata. A lazy resolver now lives for one Set-rule invocation, shared across procedures and the expected/actual compatibility checks.

Raw query keys retain generic/host display spelling and cache unknown results as well as hits. Project lookup is built only after generic/scalar/host/library resolution fails to settle a query. It indexes eligible project types by lowercased name, excluding userType, enum and standardModule; every eligible duplicate makes its name ambiguous, including repeated references to the same object. Canonical project display names and implements arrays remain unchanged. A new rule pass constructs a fresh resolver and observes metadata changes. Metadata is stable within a pass; no global mutable-array cache is added.

The existing resolveKnownObjectAssignmentType function retains its uncached behavior for other callers. The compatibility helper accepts an optional resolver and defaults to the original function. Set passes its own resolver for both the initially inferred value and a statement-specific held-object value. Actual value inference, held facts and interface-sharing compatibility remain statement-specific; caching type resolution does not cache a compatibility verdict. The ActiveSheet ANY-class predicate remains separate from this unique-type index.

## Deterministic cost

Actual metadata getters outside timing, with 100 unrelated project classes plus Box and 100 assignments:

| Mode | Before names / kinds | After names / kinds | Diagnostics |
| --- | ---: | ---: | ---: |
| Box / Nothing | 20,400 / 60,600 | 102 / 303 | 0 |
| Box / ActiveSheet | 20,501 / 60,701 | 203 / 404 | 100 |
| Box / New Box | 30,600 / 90,900 | 102 / 303 | 0 |
| Missing type / Nothing | 10,100 / 30,300 | 101 / 303 | 0 |
| 100 distinct types / Nothing | 20,400 / 60,600 | 201 / 303 | 0 |
| Object / Nothing | 0 / 0 | 0 / 0 | 0 |
| Worksheet / Nothing | 0 / 0 | 0 / 0 | 0 |
| Collection / Nothing | 0 / 0 | 0 / 0 | 0 |
| Long / Nothing | 0 / 0 | 0 / 0 | 100 |

Class indexing reads metadata once; canonical display reads account for another read for each resolved raw project type. ActiveSheet also uses its separate membership index. Other inference and interface-sharing checks can still inspect metadata. There are no production counters.

## Reproduction and timings

From the repository root with dependencies installed:

```powershell
node scripts/benchmark-set-object-types.mjs --baseline=df52ccea
node scripts/benchmark-set-object-types.mjs
```

The baseline replaces assignments.ts and typeInference.ts from df52ccea, after the ActiveSheet membership fix, in the current harness. Fixtures have the listed number of unrelated classes plus Box and 1,000 Set writes. Class modes assign Nothing, ActiveSheet or New Box; missing uses a missing declared type. Distinct mode has 1,000 variables whose K type names cycle over the available unrelated classes, reaching 1,000 distinct types at 1,000/3,000 classes. Object, Worksheet, Collection and scalar modes assign Nothing. Every sample checks 1,000 diagnostics for class-active/scalar and zero for all other modes.

Parsing, tokenization and binding are outside timing. Parsed bodies and tokens are reused, and each sample receives a fresh bound root and fresh completion caches with a root identity assertion. Measurements cover the Set rule with reused parsed bodies, not cold dataflow or full analyzer/editor latency. Three warmups precede 15 samples. Before/after timing runs were sequential after all audit tests and differential probes completed, without concurrent audit workloads.

Node v24.18.0; AMD Ryzen 7 9800X3D 8-Core Processor           ; measured locally on 2026-10-03. Times in milliseconds.

| Surfaces / mode | Before median | After median | Before p95 | After p95 |
| --- | ---: | ---: | ---: | ---: |
| 10-class-nothing | 2.125 | 1.611 | 2.964 | 1.903 |
| 10-class-active | 3.097 | 2.172 | 4.205 | 3.585 |
| 10-class-new | 1.585 | 0.889 | 2.555 | 1.408 |
| 10-missing | 0.643 | 0.363 | 1.221 | 0.634 |
| 10-distinct | 2.103 | 1.565 | 2.563 | 2.746 |
| 10-object | 1.328 | 1.363 | 2.119 | 2.217 |
| 10-worksheet | 1.472 | 1.110 | 2.397 | 1.906 |
| 10-collection | 1.250 | 1.120 | 1.568 | 1.455 |
| 10-scalar | 0.387 | 0.413 | 0.847 | 1.049 |
| 100-class-nothing | 2.902 | 1.124 | 3.490 | 1.707 |
| 100-class-active | 3.835 | 1.657 | 4.519 | 2.513 |
| 100-class-new | 3.416 | 0.795 | 4.319 | 1.631 |
| 100-missing | 1.222 | 0.381 | 1.975 | 1.134 |
| 100-distinct | 3.249 | 1.463 | 5.090 | 2.424 |
| 100-object | 1.152 | 1.123 | 3.352 | 2.027 |
| 100-worksheet | 1.369 | 1.101 | 1.780 | 2.005 |
| 100-collection | 1.190 | 1.123 | 1.702 | 1.597 |
| 100-scalar | 0.325 | 0.353 | 0.338 | 0.893 |
| 1000-class-nothing | 13.960 | 1.144 | 14.565 | 1.745 |
| 1000-class-active | 14.932 | 1.716 | 15.563 | 2.703 |
| 1000-class-new | 20.031 | 0.830 | 21.178 | 1.042 |
| 1000-missing | 6.734 | 0.388 | 6.992 | 1.089 |
| 1000-distinct | 16.047 | 2.357 | 16.348 | 3.461 |
| 1000-object | 1.163 | 1.201 | 1.226 | 1.945 |
| 1000-worksheet | 1.349 | 1.092 | 1.707 | 1.595 |
| 1000-collection | 1.189 | 1.107 | 1.591 | 1.761 |
| 1000-scalar | 0.317 | 0.355 | 0.786 | 0.770 |
| 3000-class-nothing | 40.903 | 1.218 | 44.503 | 1.635 |
| 3000-class-active | 42.232 | 1.915 | 42.920 | 2.588 |
| 3000-class-new | 60.306 | 0.920 | 65.076 | 1.455 |
| 3000-missing | 19.971 | 0.486 | 21.633 | 0.566 |
| 3000-distinct | 43.153 | 2.043 | 47.914 | 2.473 |
| 3000-object | 1.139 | 1.105 | 1.629 | 2.204 |
| 3000-worksheet | 1.400 | 1.210 | 1.782 | 1.609 |
| 3000-collection | 1.190 | 1.087 | 1.911 | 1.623 |
| 3000-scalar | 0.328 | 0.350 | 0.339 | 0.929 |

At 3,000 classes, repeated Nothing assignments improve about 33.6 times, ActiveSheet about 22.1 times, New Box about 65.6 times, missing types about 41.1 times and distinct typed targets about 21.1 times. Controls are mixed: Object medians/p95 at 10/1,000 classes regress, several scalar medians/p95 regress, 10-class distinct p95 increases, and some other control p95 values increase. The factory adds closures per pass; query maps are lazy and the project index is avoided for scalar/generic/host/library-only passes. Retained memory is proportional to raw queried types and eligible project names and is discarded with the pass. These synthetic results do not establish a universal latency improvement.

## Validation

- Type check passed.
- 257 focused tests passed, including nine new cost/semantic regressions and the existing ActiveSheet and object-assignment suites.
- Full suite: 624 files, 12,834 tests passed, 13 skipped.
- New regressions cover actual reads across procedures for Nothing/New/ActiveSheet, distinct/missing targets, eligible/value kinds, duplicate ambiguity including the same object twice, raw display/canonical implements values, lazy host/library bypasses, metadata refresh and uncached statement-specific held values. Helper results and compatibility reasons are compared with their default uncached behavior.
- 2,500 differential cases match complete analyzer and direct Set outputs with no internal errors, across four host models, metadata casing/kinds/duplicates/misses/interfaces, generic/host/library/project/scalar/Variant/qualified/array targets, bare/qualified/shadowed ActiveSheet, held-object changes, scripting values, conditional declarations, branches/loops and multiple procedures.

The earlier ActiveSheet predicate test now mocks the new resolver factory when isolating that separate guard; its real metadata-read tests remain unchanged. Generated counting and differential probes stay in audit scratch material. The timing benchmark and regression tests are committed.
