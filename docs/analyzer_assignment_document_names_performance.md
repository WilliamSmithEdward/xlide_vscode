# Incremental document-name checks for assignments

Let and Set assignment guards previously scanned every project surface for each unresolved target to ask whether any document module had the same name. The rules now share a lookup factory. Each rule invocation owns a lazy set of examined document names and a cursor into the project surfaces. A hit stops at the first matching document; later queries first check remembered names, then resume scanning. A miss exhausts the remaining surfaces once, so distinct missing names also avoid repeated scans.

Each examined surface's kind is read once per rule invocation, and each examined document name is normalized once. Only examined document names are retained; an early successful query leaves the rest of the project untouched. Declared targets and irrelevant statements never request the index. Let and Set own separate lookup instances; there is no persistent cache. Metadata is stable during an invocation and refreshed when the next invocation starts, including reuse of the same mutable context/array.

The predicate preserves any-document membership, including duplicate names on non-document surfaces. It preserves exact lowercasing without trimming or stripping qualifications. Existing declaration guards, host-specific messages, diagnostic order and spans are unchanged. Completion indexes with ambiguity or broader type semantics are not substituted.

## Reproduction

From the repository root with installed dependencies:

```powershell
node scripts/benchmark-assignment-document-names.mjs --baseline=9c5adbb9
node scripts/benchmark-assignment-document-names.mjs
```

The baseline replaces only assignments.ts from that commit in the current harness. Each fixture has the listed number of alternating class/document surfaces plus Sheet1. There are 1,000 assignments: document fixtures target Sheet1 at the end; missing fixtures repeat one missing name; distinct fixtures use 1,000 missing names; declared fixtures declare the target as Long (Let) or Object (Set); first fixtures put Sheet1 first. Let assigns 1 and Set assigns Nothing. The benchmark checks 1,000 diagnostics for document/first fixtures and zero for missing/distinct/declared fixtures.

Parsing and binding are outside timing. The parsed module is reused, and each sample gets a fresh bound root with an identity assertion. Measurements cover the direct Let rule or Set visitor plus statement traversal, not the complete analyzer or editor. Three warmups precede 15 samples. Before and after ran sequentially after tests and differential probes finished, with no concurrent audit benchmarks.

## Results

Node v24.18.0; AMD Ryzen 7 9800X3D 8-Core Processor; measured locally on 2026-10-03. Times in milliseconds.

| Surfaces / rule / mode | Before median | After median | Before p95 | After p95 |
| --- | ---: | ---: | ---: | ---: |
| 10-let-document | 0.538 | 0.373 | 3.677 | 3.822 |
| 10-set-document | 0.460 | 0.362 | 0.973 | 0.682 |
| 10-let-missing | 0.380 | 0.304 | 0.767 | 0.665 |
| 10-set-missing | 0.436 | 0.359 | 0.916 | 0.737 |
| 10-let-distinct | 0.385 | 0.315 | 0.694 | 0.753 |
| 10-set-distinct | 0.410 | 0.327 | 1.309 | 1.166 |
| 10-let-declared | 1.744 | 1.749 | 3.385 | 3.372 |
| 10-set-declared | 1.415 | 1.429 | 1.922 | 1.861 |
| 10-let-first | 0.337 | 0.348 | 0.657 | 0.714 |
| 10-set-first | 0.220 | 0.206 | 0.237 | 0.263 |
| 100-let-document | 0.530 | 0.336 | 0.835 | 0.613 |
| 100-set-document | 0.421 | 0.206 | 0.606 | 0.214 |
| 100-let-missing | 0.507 | 0.307 | 0.807 | 0.567 |
| 100-set-missing | 0.486 | 0.281 | 1.563 | 1.724 |
| 100-let-distinct | 0.915 | 0.366 | 1.026 | 0.591 |
| 100-set-distinct | 0.852 | 0.309 | 0.912 | 0.518 |
| 100-let-declared | 1.770 | 1.753 | 2.132 | 2.675 |
| 100-set-declared | 1.161 | 1.130 | 1.422 | 1.422 |
| 100-let-first | 0.323 | 0.334 | 0.812 | 0.632 |
| 100-set-first | 0.211 | 0.185 | 0.215 | 0.192 |
| 1000-let-document | 2.545 | 0.362 | 3.312 | 0.789 |
| 1000-set-document | 2.423 | 0.195 | 2.777 | 0.198 |
| 1000-let-missing | 2.550 | 0.330 | 2.987 | 0.697 |
| 1000-set-missing | 2.480 | 0.288 | 2.808 | 0.316 |
| 1000-let-distinct | 2.564 | 0.380 | 3.013 | 0.384 |
| 1000-set-distinct | 2.506 | 0.313 | 2.877 | 0.983 |
| 1000-let-declared | 1.683 | 1.684 | 2.111 | 2.558 |
| 1000-set-declared | 1.143 | 1.120 | 1.363 | 1.384 |
| 1000-let-first | 0.328 | 0.340 | 0.382 | 0.719 |
| 1000-set-first | 0.210 | 0.183 | 0.496 | 0.189 |
| 3000-let-document | 7.465 | 0.406 | 7.555 | 1.178 |
| 3000-set-document | 7.250 | 0.232 | 8.090 | 0.237 |
| 3000-let-missing | 7.366 | 0.369 | 7.585 | 0.438 |
| 3000-set-missing | 7.316 | 0.323 | 7.830 | 0.330 |
| 3000-let-distinct | 7.399 | 0.412 | 8.035 | 0.468 |
| 3000-set-distinct | 7.347 | 0.347 | 7.576 | 0.804 |
| 3000-let-declared | 1.707 | 1.709 | 2.084 | 2.148 |
| 3000-set-declared | 1.112 | 1.099 | 1.691 | 1.503 |
| 3000-let-first | 0.342 | 0.341 | 0.426 | 1.014 |
| 3000-set-first | 0.211 | 0.184 | 0.486 | 0.236 |

For 3,000 surfaces, document-hit medians improve about 18 times for Let and 31 times for Set. Distinct misses improve about 18 and 21 times respectively. Declared and first-hit controls have nearly unchanged medians, while some small/control p95 results regress. These results establish the repeated-scan reduction, not a universal analyzer latency improvement. A query may still inspect all remaining surfaces; retaining scanned document names uses memory proportional to the examined document subset.

## Deterministic cost and validation

Separate getter instrumentation on actual metadata, outside timing, used 100 alternating surfaces plus Sheet1 and 100 assignments:

| Query mode (both rules separately) | Before kind / name reads | After kind / name reads | Diagnostics |
| --- | ---: | ---: | ---: |
| Document at end | 10,100 / 5,100 | 101 / 51 | 100 |
| Repeated missing | 10,100 / 5,100 | 101 / 51 | 0 |
| Distinct missing | 10,100 / 5,100 | 101 / 51 | 0 |
| Declared | 0 / 0 | 0 / 0 | 0 |
| Document first | 100 / 100 | 1 / 1 | 100 |

- Type check passed.
- Focused assignment/form tests: 190 passed.
- Full suite: 616 files, 12,757 tests passed, 13 skipped.
- Eighteen regressions cover exact cost bounds across procedures, early unused-tail behavior, resumed queries, retained names after misses, duplicate kinds/case, exact qualification/whitespace, declarations, mutable metadata refresh and host-specific messages.
- 1,500 differential cases preserve complete analyzer outputs and direct Let/Set outputs across declarations, conditional activity, project duplicates/kinds/case, host variants, multiple procedures, branches/loops and member assignments.

Instrumentation and generated differential cases remain audit scratch material. The timing benchmark and regression tests are committed.
