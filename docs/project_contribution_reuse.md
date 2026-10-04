# Project visibility contribution reuse

ProjectIndex cached each module's visibility contribution in the whole-project revision memo. An edit to any module discarded every contribution, so unchanged symbol lists and signatures were rebuilt. Contributions now live in a map owned by the current module key. setModule/removeModule delete that module's map; whole-project queries still clear on every project change. Each query retains its own/other visibility sides separately. Aggregation order and freshly returned outer collections are unchanged.

The ten contribution callbacks use the contributing module's symbols and kind, its own source/host predeclared-instance metadata, its locally resolved integer constants, or other contributions of that same module. None incorporates another module's declarations. Type shadowing, duplicate-constant handling and merged member surfaces remain outside these local caches and are rebuilt against the current project.

Each module stores at most the existing query/visibility-side combinations, populated lazily. Replacement/removal drops all its entries. There is no per-asking-module answer cache or source-version chain. Retained nested contribution objects follow the existing shared-part behavior; this change does not introduce deep cloning or a new mutation-isolation contract.

Validation against d084e4caad70fed707316bfed9178b13fd2c7df5:

- Types and 28 focused tests across five files pass.
- Full suite: 727 files passed, seven skipped; 14,448 tests passed, 32 skipped. No failures.
- Three work-count tests fail on the baseline; the replacement/removal control passes. Across three edits, unchanged root.children reads decrease from 48/534/2,694 to zero for 2/20/100-module projects. All nine complete query outputs, for external and own-module callers, match freshly rebuilt indexes after each edit. Case-insensitive replacement, changed module roles, predeclared-instance metadata, explicitly active/inactive conditional branches, removal and re-addition are covered. Existing private visibility and caller-owned outer collection tests pass.

Frozen-input differential: 8,256 corpus sources plus 42 generated resource/activity cases. All 66,384 visibility snapshots per version match across initial/edited own and external callers, removal and empty-project states; edited/removal/empty outputs also match fresh indexes. Each snapshot contains all nine complete query outputs. The 8,298 complete project-aware diagnostic arrays match (13,628 findings), with zero internal errors. ASTs, lexer tokens and trivia are frozen.

Benchmark: node scripts/benchmark-project-contributions.mjs --baseline=d084e4caad70fed707316bfed9178b13fd2c7df5 --rounds=9 and the same command without baseline. Baseline/candidate/candidate/baseline order, three warmups/nine measured rounds; Node 24.18.0 on Ryzen 7 9800X3D. Each module has 100 public procedures. Timing includes all nine visibility queries after one module edit, with parsing/replacement excluded and initial queries warmed. Every complete output matches an independent fresh index; generated count/new-name/old-name expectations are checked too.

| Modules | Before median ms | After median ms |
| --- | --- | --- |
| 1 | 0.184–0.198 | 0.193–0.236 |
| 20 | 1.472–1.710 | 0.352–0.362 |
| 100 | 8.024–9.395 | 1.549–1.641 |

The single-module control has no unchanged contributions to reuse and shows no speedup. These component measurements do not establish total typing or completion latency. Initial cold queries retain their existing symbol work; merging still visits all project modules.
