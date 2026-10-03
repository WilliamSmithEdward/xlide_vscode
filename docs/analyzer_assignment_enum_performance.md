# Analyzer enum assignment lookup — 2026-10-03

checkAssignmentTypes concatenated and scanned all module/project symbols for every assignment to determine whether the target type is an enum. A module without enums still paid for the full scan. Collect case-insensitive enum names once per rule pass and use Set membership for each assignment. This preserves qualified short-name matching and the existing module/project enum semantics. The index is local to the pass; it has no cross-pass invalidation requirements.

Run `node scripts/benchmark-assignment-enums.mjs --rounds=15` in each checkout. The benchmark holds parsing and symbol construction outside timing, uses three warmups and 15 samples, and validates that the scalar assignments emit no errors. Node 24.18.0, Ryzen 7 9800X3D; median milliseconds:

| Module members / scalar assignments | Before | After |
| --- | ---: | ---: |
| 100 / 1,000 | 6.963 | 5.975 |
| 1,000 / 1,000 | 9.830 | 7.628 |
| 1,000 / 5,000 | 51.623 | 40.819 |

These synthetic timings isolate one diagnostic rule and do not predict total editor latency. The regression counts unrelated module-variable kind reads, so it detects repeated whole-symbol scans without depending on timing. It fails on the original rule and passes after indexing. Behavior tests cover case-insensitive module enum names, qualified project enum targets and replacement project visibility. Existing enum and assignment tests pass; type checking and compilation pass. Full-suite results are recorded in the PR.
