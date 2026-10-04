# Project opened-file fact reuse

ProjectIndex previously rescanned every module for Open file numbers after any edit, despite unchanged source in the other modules. The whole-project query memo is still revision-scoped. A new per-module map retains only current module facts, populated lazily and removed on setModule/removeModule using the existing case-insensitive key. Aggregation order, any and the complete number set are unchanged. Removed/replaced sources leave no version chain in this cache.

Validation against f006e788:

- Full unit suite: 719 files passed, seven skipped; 14,325 tests passed, 31 skipped. No failures.
- Four work-count regressions fail before the fix; the initial-lazy/stable-query control passes. With 2/20/100 modules and three edits to one module, scans decrease from 6/60/300 to three. Addition scans one new module; removal scans none; replacing an unknown number with a literal clears any. Empty-project and case-insensitive replacement/removal behavior are covered.
- Frozen-input differential: 8,256 corpus sources plus 42 generated resource/activity cases; 33,192 complete project-fact queries per version through initial, edited, removed and empty states all match, with independently checked aggregation/removal controls. All 8,298 complete diagnostic arrays (13,504 findings) match; no internal errors. ASTs and lexer tokens/trivia are frozen.
- Types and 29 focused checks across five files pass, including file-number diagnostics, project visibility/service and module-state reuse.

Benchmark: node scripts/benchmark-project-file-facts.mjs --baseline=f006e788 --rounds=9 and the same command without baseline. Three warmups/nine measured rounds; baseline/candidate/candidate/baseline order; Node 24.18.0 / Ryzen 7 9800X3D. Each module has 1,000 statements. The initial query is warmed in both versions; setModule/parsing are excluded from timing. Every complete post-edit set and any flag is independently checked.

| Modules | Before median ms | After median ms |
| --- | --- | --- |
| 1 | 0.064–0.093 | 0.068–0.079 |
| 20 | 8.550–8.672 | 0.038–0.041 |
| 100 | 33.808–35.328 | 0.040–0.041 |

This improves a repeated project query. The initial project scan still reads every module, and these component timings do not establish total typing/diagnostic latency. Retained fact sets add memory proportional to the current modules' distinct file numbers; entries disappear on replacement/removal.
