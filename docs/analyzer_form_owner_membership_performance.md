# Qualified form-control owner membership

Set f.Answer = Nothing checks whether a member returning MSForms.* belongs to a project UserForm. Previously every eligible assignment scanned all project surfaces for an exact owner name. The rule now owns a lazy incremental UserForm-name lookup shared across procedure checks. Examined owner names are remembered, a hit stops at the first match, and new queries resume from the cursor. A miss exhausts the remaining surfaces once.

The existing document-name lookup uses the same cursor/set algorithm. Generalize that private helper with an explicit surface kind and case mode instead of duplicating it: documents remain case-insensitive and form owners remain case-sensitive. Both predicates retain any matching selected kind, even with duplicate names on other kinds. No trimming or qualification stripping is introduced. Each rule/predicate owns its index, refreshed for subsequent invocations; project metadata is stable within an invocation. Target, argument, return-type and later completion behavior remains unchanged.

## Deterministic cost

Actual metadata getters count all project name/kind reads, including completion setup, with 100 unrelated UserForms and 100 assignments:

| Fixture | Before names / kinds | After names / kinds | Diagnostics |
| --- | ---: | ---: | ---: |
| Form1 last | 10,202 / 10,302 | 203 / 303 | 100 |
| Form1 first | 202 / 302 | 103 / 203 | 100 |
| Let control | 102 / 202 | 102 / 202 | 0 |
| Other return type | 102 / 202 | 102 / 202 | 0 |

Remaining reads belong to other completion work. Separate document probes confirm identical metadata read counts and diagnostics in all ten prior document controls, including repeated/distinct misses, first hits and declared targets. Counters run outside timing; production has no counters.

## Reproduction and timings

From the repository root with installed dependencies:

```powershell
node scripts/benchmark-form-owner-membership.mjs --baseline=7296f7cb
node scripts/benchmark-form-owner-membership.mjs
node scripts/benchmark-assignment-document-names.mjs --baseline=7296f7cb
node scripts/benchmark-assignment-document-names.mjs
```

The baseline replaces only assignments.ts from that commit in the current harness. Form fixtures have the listed number of unrelated UserForms plus Form1, which exposes Answer As MSForms.Label. Each procedure has 1,000 qualified Set control assignments. First/last fixtures vary Form1's position; Let controls assign 1, and other-return controls expose Answer As Object. The benchmark checks 1,000 diagnostics for qualified Set controls and none for unused predicate controls.

Parsing, tokenization and binding are outside timing. Parsed modules/tokens are reused, while each sample gets a fresh bound root and fresh completion caches with an identity assertion. Measurements cover checkAssignmentTypes, not full analyzer/editor latency. Three warmups precede 15 samples. All four timing runs were sequential, after audit tests/differential probes completed.

Node v24.18.0; AMD Ryzen 7 9800X3D 8-Core Processor; measured locally on 2026-10-03. Times in milliseconds.

| Surfaces / mode | Before median | After median | Before p95 | After p95 |
| --- | ---: | ---: | ---: | ---: |
| 10-last | 2.430 | 2.206 | 3.398 | 2.945 |
| 10-first | 1.280 | 1.222 | 2.104 | 1.698 |
| 10-let | 1.914 | 1.848 | 2.874 | 2.728 |
| 10-other-return | 1.136 | 1.139 | 1.652 | 1.866 |
| 100-last | 1.384 | 1.172 | 1.949 | 1.672 |
| 100-first | 1.163 | 1.170 | 1.601 | 1.688 |
| 100-let | 1.478 | 1.453 | 2.825 | 2.832 |
| 100-other-return | 1.379 | 1.363 | 2.068 | 2.222 |
| 1000-last | 3.155 | 1.357 | 3.235 | 1.749 |
| 1000-first | 1.170 | 1.158 | 1.219 | 2.324 |
| 1000-let | 1.439 | 1.434 | 2.316 | 2.380 |
| 1000-other-return | 1.144 | 1.126 | 2.121 | 1.730 |
| 3000-last | 7.657 | 1.215 | 8.268 | 1.261 |
| 3000-first | 1.192 | 1.162 | 1.238 | 2.139 |
| 3000-let | 1.447 | 1.433 | 2.645 | 2.995 |
| 3000-other-return | 1.149 | 1.118 | 2.272 | 2.362 |

The 3,000-surface last-owner median improves about 6.3 times. First-match and unused-control medians remain close to baseline, while some control p95 results regress. The first query still scans until a match or the end, and other completion costs remain. The lookup retains only examined selected-kind names and is discarded with the rule pass. These results do not establish a universal latency improvement.

## Shared document lookup controls

The existing document benchmark uses alternating class/document surfaces plus Sheet1 and 1,000 Let/Set assignments. Document fixtures put Sheet1 last; missing/distinct fixtures use one or many missing names; declared fixtures declare the target; first fixtures put Sheet1 first. Diagnostic counts are checked for every sample.

| Surfaces / mode | Before median | After median | Before p95 | After p95 |
| --- | ---: | ---: | ---: | ---: |
| 10-let-document | 0.505 | 0.478 | 1.262 | 3.923 |
| 10-set-document | 0.356 | 0.328 | 0.487 | 0.592 |
| 10-let-missing | 0.367 | 0.302 | 0.665 | 0.375 |
| 10-set-missing | 0.370 | 0.354 | 0.752 | 0.743 |
| 10-let-distinct | 0.336 | 0.329 | 0.716 | 0.661 |
| 10-set-distinct | 0.347 | 0.368 | 1.222 | 1.234 |
| 10-let-declared | 1.798 | 1.738 | 3.824 | 3.361 |
| 10-set-declared | 1.452 | 1.440 | 2.155 | 1.899 |
| 10-let-first | 0.359 | 0.333 | 0.587 | 0.582 |
| 10-set-first | 0.227 | 0.204 | 0.250 | 0.226 |
| 100-let-document | 0.348 | 0.325 | 0.690 | 0.633 |
| 100-set-document | 0.221 | 0.211 | 0.248 | 0.245 |
| 100-let-missing | 0.348 | 0.306 | 0.650 | 0.390 |
| 100-set-missing | 0.280 | 0.305 | 1.425 | 1.652 |
| 100-let-distinct | 0.395 | 0.366 | 0.633 | 0.631 |
| 100-set-distinct | 0.300 | 0.360 | 0.749 | 0.765 |
| 100-let-declared | 2.597 | 1.770 | 5.374 | 2.700 |
| 100-set-declared | 1.159 | 1.162 | 1.429 | 1.429 |
| 100-let-first | 0.342 | 0.315 | 0.657 | 0.783 |
| 100-set-first | 0.215 | 0.206 | 0.237 | 0.296 |
| 1000-let-document | 0.362 | 0.338 | 0.769 | 0.750 |
| 1000-set-document | 0.220 | 0.210 | 0.225 | 0.221 |
| 1000-let-missing | 0.352 | 0.322 | 0.403 | 0.364 |
| 1000-set-missing | 0.295 | 0.307 | 0.328 | 0.324 |
| 1000-let-distinct | 0.394 | 0.372 | 0.669 | 0.519 |
| 1000-set-distinct | 0.315 | 0.333 | 0.345 | 0.369 |
| 1000-let-declared | 1.739 | 1.684 | 2.074 | 2.365 |
| 1000-set-declared | 1.144 | 1.150 | 1.510 | 1.467 |
| 1000-let-first | 0.346 | 0.324 | 0.672 | 0.772 |
| 1000-set-first | 0.209 | 0.200 | 0.227 | 0.237 |
| 3000-let-document | 0.396 | 0.373 | 0.850 | 0.852 |
| 3000-set-document | 0.267 | 0.250 | 0.330 | 0.290 |
| 3000-let-missing | 0.392 | 0.373 | 0.929 | 0.722 |
| 3000-set-missing | 0.326 | 0.345 | 0.349 | 0.406 |
| 3000-let-distinct | 0.422 | 0.405 | 0.434 | 0.421 |
| 3000-set-distinct | 0.353 | 0.373 | 0.948 | 0.785 |
| 3000-let-declared | 1.720 | 1.687 | 2.127 | 2.092 |
| 3000-set-declared | 1.122 | 1.123 | 1.417 | 1.394 |
| 3000-let-first | 0.338 | 0.318 | 0.741 | 0.745 |
| 3000-set-first | 0.206 | 0.202 | 0.218 | 0.210 |

Document timings are small and mixed despite unchanged deterministic metadata counts. For example, 3,000-surface Set missing/distinct medians change from 0.326/0.353 to 0.345/0.373 ms. The generalized helper adds a kind/case choice; these controls and p95 regressions are included rather than claiming unrelated document speedups.

## Validation

- Type check passed.
- Focused form/document/assignment tests: 197 passed.
- Full suite: 621 files, 12,802 tests passed, 13 skipped.
- Seven regressions cover actual read bounds across procedures, exact case/kind and duplicate semantics, resumed queries, retained names after misses, mutable metadata refresh and unused controls. Real completion tests establish the cost reduction; isolated resolved-member fixtures exercise exact owner casing independently of completion normalization.
- Existing document lookup regressions pass, and all ten document counting controls match exactly.
- 1,500 differential cases preserve complete analyzer and direct assignment outputs across owner casing, duplicate kinds, accessors/return types, Let/arguments/missing controls, With/branches/loops, conditional receivers, multiple procedures and shared document lowercasing.

Generated differential/counting probes remain audit scratch material. The form timing benchmark and regression tests are committed; the existing document benchmark is reused.
