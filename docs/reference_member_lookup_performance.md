# Reference member definition lookup

Reference collection previously reconstructed member surfaces and linearly searched their names for each qualified occurrence. Each module's reference pass now gets a fresh memberSurfaceCache. When a caller supplies that cache, resolveMemberDefinitionsAt reuses the existing first-match surfaceMemberNamed index. Direct calls without a cache retain the linear lookup. There is no new global source cache or second index implementation.

The module pass owns the context lifetime: fixed source and project metadata during the synchronous collection, discarded before later collections or another module. Cached direct callers retain the existing requirement to use a fresh context when metadata changes.

## Deterministic work

Counting the actual definition find in collectSymbolReferences showed 100 comparisons for ten last-member references against ten members; 10,000 for 100 against 100; and 1,000,000 for 1,000 against 1,000. Source-backed metadata construction repeats work even for first-member references.

A separate stable getter measurement on frozen actual project-index member objects/arrays checks complete expected caller locations and yields:

| Members and references | Target | Baseline metadata name reads | Fixed metadata name reads |
| ---: | --- | ---: | ---: |
| 1,000 | First | 1,002,002 | 3,002 |
| 1,000 | Last | 2,002,001 | 4,001 |

Six new tests include these two collection work bounds, a cached last-hit work bound, an uncached first-hit control, first case-insensitive duplicate definitions and changed metadata between fresh contexts. Three work tests fail on baseline; all six pass with the fix. The uncached host API still reads only the first member's name for a first hit after host-model warmup. A cached context reads 1,000 names once across 1,000 last-member queries.

## Timings

Windows, Ryzen 7 9800X3D, Node v24.18.0; baseline 0f685d45. Four isolated processes ran baseline/fixed/fixed/baseline. Each configuration has three warmups and fifteen measured rounds. Ranges span the two processes for each version, in milliseconds. p95 is the largest of fifteen samples; very small observations are not reliable speedup estimates. All 42 result hashes match across all four processes.

Collector timings include the complete actual reference collection with warm source/project caches. Each module has ordinary public procedures, and callers use Library.P0 or Library.P[last]. Zero-call controls still request the declaration. Project and source construction plus output/invariant assertions are outside timing.

Direct API timings use source-backed project member definitions with a fresh context per measured round. Cached-one includes initial surface/index construction. Cached-many runs 1,000 identical queries through one fresh context and returns the final result; all repeated-query behavior is separately tested. Uncached-one retains the original scan. First, last and missing lookups are included. No cold project construction, Office execution or VS Code UI timing is claimed.

| Scope | Members | Calls/queries | Target | Baseline median ms | Fixed median ms | Baseline p95 ms | Fixed p95 ms |
| --- | ---: | ---: | --- | ---: | ---: | ---: | ---: |
| collector | 1 | 0 | last | 0.027–0.037 | 0.029–0.031 | 0.075–0.085 | 0.077–0.088 |
| collector | 1 | 1 | first | 0.028–0.030 | 0.030–0.031 | 0.041–0.044 | 0.042–0.044 |
| collector | 1 | 1 | last | 0.016–0.019 | 0.020–0.022 | 0.025–0.030 | 0.029–0.033 |
| collector | 1 | 1000 | first | 2.300–2.316 | 2.246–2.421 | 2.514–2.916 | 3.328–4.077 |
| collector | 1 | 1000 | last | 2.060–2.094 | 1.935–2.002 | 2.276–2.476 | 2.726–2.767 |
| direct-uncached-one | 1 | 1 | first | 0.002–0.002 | 0.002–0.002 | 0.005–0.009 | 0.005–0.005 |
| direct-cached-one | 1 | 1 | first | 0.002–0.002 | 0.002–0.002 | 0.002–0.002 | 0.003–0.014 |
| direct-cached-many | 1 | 1000 | first | 0.632–0.652 | 0.633–0.646 | 0.886–1.040 | 0.807–1.043 |
| direct-uncached-one | 1 | 1 | last | 0.001–0.001 | 0.001–0.001 | 0.001–0.002 | 0.001–0.003 |
| direct-cached-one | 1 | 1 | last | 0.001–0.001 | 0.001–0.002 | 0.001–0.001 | 0.003–0.004 |
| direct-cached-many | 1 | 1000 | last | 0.410–0.418 | 0.433–0.438 | 0.459–0.543 | 0.560–0.606 |
| direct-uncached-one | 1 | 1 | missing | 0.001–0.001 | 0.001–0.001 | 0.001–0.001 | 0.003–0.003 |
| direct-cached-one | 1 | 1 | missing | 0.001–0.001 | 0.001–0.001 | 0.001–0.001 | 0.001–0.001 |
| direct-cached-many | 1 | 1000 | missing | 0.414–0.432 | 0.410–0.427 | 0.469–0.572 | 0.440–0.485 |
| collector | 100 | 0 | last | 0.061–0.062 | 0.059–0.061 | 0.093–0.121 | 0.090–0.094 |
| collector | 100 | 1 | first | 0.061–0.061 | 0.064–0.066 | 0.068–0.230 | 0.070–0.072 |
| collector | 100 | 1 | last | 0.064–0.064 | 0.068–0.072 | 0.069–0.075 | 0.320–0.347 |
| collector | 100 | 1000 | first | 5.313–5.501 | 1.995–2.035 | 5.817–5.959 | 2.636–2.705 |
| collector | 100 | 1000 | last | 5.840–6.231 | 2.020–2.049 | 6.979–7.440 | 2.317–2.525 |
| direct-uncached-one | 100 | 1 | first | 0.004–0.004 | 0.005–0.005 | 0.004–0.007 | 0.005–0.005 |
| direct-cached-one | 100 | 1 | first | 0.004–0.004 | 0.008–0.008 | 0.004–0.005 | 0.009–0.009 |
| direct-cached-many | 100 | 1000 | first | 0.413–0.434 | 0.405–0.431 | 0.438–0.496 | 0.431–0.452 |
| direct-uncached-one | 100 | 1 | last | 0.004–0.005 | 0.006–0.006 | 0.005–0.005 | 0.013–0.015 |
| direct-cached-one | 100 | 1 | last | 0.005–0.005 | 0.008–0.008 | 0.005–0.007 | 0.009–0.009 |
| direct-cached-many | 100 | 1000 | last | 0.970–0.974 | 0.479–0.489 | 1.145–1.496 | 0.626–0.630 |
| direct-uncached-one | 100 | 1 | missing | 0.004–0.004 | 0.005–0.005 | 0.004–0.005 | 0.009–0.013 |
| direct-cached-one | 100 | 1 | missing | 0.004–0.004 | 0.008–0.008 | 0.004–0.006 | 0.009–0.009 |
| direct-cached-many | 100 | 1000 | missing | 0.798–0.822 | 0.403–0.424 | 0.880–0.920 | 0.550–0.559 |
| collector | 1000 | 0 | last | 0.522–0.538 | 0.539–0.600 | 1.076–1.345 | 1.226–1.487 |
| collector | 1000 | 1 | first | 0.489–0.501 | 0.553–0.691 | 0.693–1.137 | 1.197–1.235 |
| collector | 1000 | 1 | last | 0.510–0.517 | 0.549–0.564 | 0.798–0.807 | 0.891–0.902 |
| collector | 1000 | 1000 | first | 40.153–40.986 | 2.495–2.543 | 41.544–41.856 | 2.878–2.909 |
| collector | 1000 | 1000 | last | 47.394–47.408 | 2.583–2.614 | 52.354–54.011 | 2.942–3.894 |
| direct-uncached-one | 1000 | 1 | first | 0.038–0.039 | 0.038–0.039 | 0.042–0.043 | 0.040–0.047 |
| direct-cached-one | 1000 | 1 | first | 0.038–0.039 | 0.075–0.077 | 0.041–0.065 | 0.095–0.101 |
| direct-cached-many | 1000 | 1000 | first | 0.439–0.439 | 0.473–0.480 | 0.453–0.682 | 0.537–0.915 |
| direct-uncached-one | 1000 | 1 | last | 0.043–0.045 | 0.045–0.045 | 0.047–0.048 | 0.060–0.061 |
| direct-cached-one | 1000 | 1 | last | 0.044–0.045 | 0.074–0.076 | 0.050–0.330 | 0.079–0.081 |
| direct-cached-many | 1000 | 1000 | last | 5.890–6.122 | 0.477–0.490 | 6.327–9.452 | 0.645–1.309 |
| direct-uncached-one | 1000 | 1 | missing | 0.043–0.044 | 0.043–0.044 | 0.043–0.045 | 0.049–0.050 |
| direct-cached-one | 1000 | 1 | missing | 0.043–0.044 | 0.074–0.076 | 0.048–0.058 | 0.080–0.907 |
| direct-cached-many | 1000 | 1000 | missing | 4.741–5.230 | 0.476–0.484 | 4.884–5.544 | 0.487–0.499 |

For 1,000 members and 1,000 last-member references, complete collection improves from 47.394–47.408 ms to 2.583–2.614 ms; first-member references improve from 40.153–40.986 ms to 2.495–2.543 ms. A single cached first-member query with 1,000 members increases from 0.038–0.039 ms to 0.075–0.077 ms, paying index construction up front. The uncached first-member API remains 0.038–0.039 ms. Zero-call and single-call collector controls are slower or mixed: the 1,000-member zero-call control increases from 0.522–0.538 ms to 0.539–0.600 ms. The change addresses repeated queries and does not establish universal speedups.

Reproduce from the repository root:

```powershell
node scripts/benchmark-reference-member-definitions.mjs --baseline=0f685d45
node scripts/benchmark-reference-member-definitions.mjs
node scripts/benchmark-reference-member-definitions.mjs
node scripts/benchmark-reference-member-definitions.mjs --baseline=0f685d45
npm exec vitest run tests/referenceDefinitionMemberWork.test.ts tests/referenceCollectorPrefixWork.test.ts tests/vbaReferenceResolution.test.ts tests/vbaBangAndBracketReferences.test.ts tests/vbaReferenceKinds.test.ts tests/vbaProjectReferences.test.ts tests/vbaFormReferenceWiring.test.ts
npm run check-types
npm test -- --run
```

## Validation

Seven focused files pass 60 tests. Type checking passes. After resolving the context insertion conflict with #911, both per-pass caches are retained. Type checking and the final full suite pass on main0f685d45 plus this change: 654 files and 13,248 tests, with 13 skipped. All 3,600 differential comparisons, metadata work counts and four timing runs were repeated on this combined tree. Later edits only update documentation.

An ad hoc differential audit loads both original runtime files for its baseline and compares complete reference results, ambiguity and kinds across 450 fixtures and 3,600 invocations. It covers qualified/bare/With and nested/malformed/module-level With, colons, labels, continuations, comments/strings, inline If, bracketed/suffixed names, local/module/procedure shadows, unrelated class receivers, attributes, LF/CRLF/CR and declaration/caller invocation sites with both declaration settings. All match, with cached full-source token objects/arrays frozen and zero uncaught exceptions. Metadata work tests freeze member objects and arrays; no frozen AST or complete repository coverage is claimed.
