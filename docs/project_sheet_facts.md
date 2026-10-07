# Project worksheet fact reuse

After any module edit, ProjectIndex.sheetChanges rescanned every source, including unchanged modules. Retain each current module's worksheet facts lazily, remove its entry on case-insensitive replacement/removal, and merge those facts under the existing project-revision cache. Names and their insertion order, addsSheets and assignsComputedName remain unchanged. The cache keeps no source version chain.

Validation against aac22134779b3984cf834e575aa302e1d6c34452:

- Full suite: 726 files passed, seven skipped; 14,437 tests passed, 32 skipped. No failures.
- Types and 50 focused tests across five files pass.
- Four work-count regressions fail on the baseline; the lazy-query control passes. Three edits in 2/20/100-module projects require 6/60/300 source scans before and three after. Removal scans no survivors. Replacement/removal clear stale names and flags; addition, module roles, empty projects, case-insensitive keys and stable query identity are covered.
- Frozen-input differential: 8,256 corpus sources plus 42 generated resource/activity cases; 33,192 complete project-fact queries per version through initial, edited, removed and empty states match, with independent aggregate/removal expectations. All 8,298 complete diagnostic arrays match (13,491 findings), with zero internal errors. ASTs, tokens and trivia are frozen.

Benchmark: node scripts/benchmark-project-sheet-facts.mjs --baseline=aac22134779b3984cf834e575aa302e1d6c34452 --rounds=9 and the same command without baseline. Baseline/candidate/candidate/baseline order, three warmups and nine measured rounds; Node 24.18.0 on Ryzen 7 9800X3D. Each module has 1,000 statements. Both versions warm the initial query; replacement/parsing are excluded from timing. Every complete result and name order is checked independently.

| Modules | Before median ms | After median ms |
| --- | --- | --- |
| 1 | 0.093–0.120 | 0.058–0.060 |
| 20 | 6.490–6.593 | 0.053–0.054 |
| 100 | 30.873–31.826 | 0.057–0.059 |

These measurements cover the project worksheet query after an edit, not total editor latency. The initial query still scans every module. Memory grows with current module facts and their distinct worksheet names; entries are deleted on replacement/removal.
