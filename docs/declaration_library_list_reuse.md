# Declaration reference-library list reuse

`checkInvalidAsTypeNames` resolved the same referenced-library list for every unresolved type declaration. With 1,000 declarations and four references, that made 4,000 calls to `libraryTypeNames` and allocated 1,000 arrays.

Resolve the list lazily once per rule invocation. The corresponding regression now requires four library lookups. Resolved builtin/host types and inactive declarations still avoid library lookups. No list survives the invocation: caller reference-list changes between analyses remain visible. Absent, empty and unmodelled references retain their existing diagnostic behavior.

## Validation

Twelve new tests cover deterministic work counts, complete diagnostic messages/spans, LF/CRLF/CR, resolved types, inactive conditional branches, unknown references, and caller-list changes between analyses. Six work-count cases fail on the baseline. Type checking and 29 focused tests passed.

An independent differential check against `7b0f745bcce41c386cab2e0d31dc362d64fcaadf` compared 102,672 complete rule results and 16,512 complete analyzer diagnostic/failure results across 8,556 sources, six reference-list variants, and two conditional environments. AST nodes and reference lists were frozen. Three additional caller-list mutation queries matched.

## Measurement

Run `node scripts/benchmark-declaration-library-list.mjs`, or add `--baseline=7b0f745bcce41c386cab2e0d31dc362d64fcaadf` for the old rule implementation. The benchmark asserts independent complete diagnostic results outside the timer. It measures complete rule invocation with parsing excluded and metadata/source facts warmed: three warmups, nine rounds, separate processes in before/after/after/before order. Environment: Node 24.18.0, AMD Ryzen 7 9800X3D.

For 10,000 declarations, medians across the two runs were:

| Reference mode | Before (ms) | After (ms) |
| --- | --- | --- |
| Four modelled libraries | 3.5004–3.6843 | 2.3721–2.6557 |
| Includes unmodelled ADODB | 3.6736–4.4125 | 2.3045–2.5117 |
| Empty list | 2.0023–2.2919 | 1.9440–2.1445 |
| Absent list | 1.9644–1.9780 | 2.1274–2.1616 |
| Resolved builtin type | 0.3455–0.3513 | 0.4440–0.4736 |

The absent-list and resolved-builtin controls were slightly slower; this is not a universal speedup. The maximum measured fixed modelled-library sample was 5.6416 ms. This does not measure complete analyzer or editor latency, cold initialization, or heap bytes.

Full repository suite: 751 files passed, seven skipped; 14,680 tests passed, 33 skipped (two workers, 286.19 seconds).
