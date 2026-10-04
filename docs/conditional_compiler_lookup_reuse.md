# Query-local conditional compiler lookup reuse

Each evaluated directive previously rebuilt and case-folded the compiler-constant map. These compiler values stay fixed during a replay even as module `#Const` definitions evolve. A 400-definition replay initialized that lookup 400 times for indexing, 401 for tracker/offset activity, and 402 for Null scanning.

The repair creates one lazy expression evaluator per replay. Its compiler lookup initializes at most once; its overlay reads the live module-constant map so later definitions still shadow compiler and project values. Indexing preserves its distinct missing-name-as-zero policy, and each new query constructs a fresh evaluator. No evaluator or map is persisted between queries. Constant indexing also reuses the existing project-constant initialization helper instead of duplicating it.

Validation against merged base `92db624c63439ef44a3e2a8f5e54ddc6a6745237`:

- Types pass; 130 focused tests pass, 2 skipped.
- Eight new independent work/semantic controls. Baseline fails four compiler-map initialization assertions and passes four controls; repaired code passes all eight. The 400 / 401 / 402 initializations become one, and a replay with no expression still initializes none.
- Full suite: 773 files / 15,109 tests pass, 7 files / 33 tests skipped, 72.90 seconds. The already-published replay-gate dependency was included in this run. After moving onto its actual merged base, 1,317 tested source, test and configuration file contents remain unchanged.
- 8,256 oracle sources across LF, CRLF and CR and two environments yield 49,536 complete comparisons, with no differences. These compare constant indexes, Null conditions, tracker activity, offset activity replay, full module diagnostics and internal failures against the actual merged base.

Reproduce:

```powershell
node scripts/benchmark-conditional-compiler-lookup.mjs --baseline=92db624c63439ef44a3e2a8f5e54ddc6a6745237
node scripts/benchmark-conditional-compiler-lookup.mjs
```

Node v24.18.0, AMD Ryzen 7 9800X3D. ABBA order, two baseline and two repaired runs; three warmups and nine measured rounds. Source/environment creation and AST warming stay outside the clock. Complete tracker construction, constant indexing or `analyzeModule` diagnostics stay inside. Every call independently verifies the complete expected constant sequence, branch activity or complete empty diagnostic output and the absence of internal failures.

| Definitions | Extra compiler flags | Operation | Baseline median ms | Repaired median ms |
| ---: | ---: | --- | ---: | ---: |
| 1 | 0 | Tracker | 0.0235–0.0266 | 0.0235–0.0243 |
| 1 | 0 | Index | 0.0060–0.0062 | 0.0055 |
| 1 | 0 | Full diagnostics | 0.5230–0.5433 | 0.5129–0.5189 |
| 1,000 | 0 | Tracker | 1.8042–2.0662 | 1.2998–1.3168 |
| 1,000 | 0 | Index | 1.4127–1.4225 | 0.9875–1.0607 |
| 1,000 | 0 | Full diagnostics | 11.1695–11.2820 | 9.9091–10.2161 |
| 1 | 1,000 | Tracker | 0.3834–0.4171 | 0.3318–0.3676 |
| 1 | 1,000 | Index | 0.1322–0.1384 | 0.1467–0.1526 |
| 1 | 1,000 | Full diagnostics | 1.3019–1.3169 | 1.0468–1.1139 |
| 1,000 | 1,000 | Tracker | 64.0500–65.7410 | 1.1396–1.1496 |
| 1,000 | 1,000 | Index | 63.9160–64.4941 | 1.1194–1.1300 |
| 1,000 | 1,000 | Full diagnostics | 204.1130–207.5285 | 11.0624–11.0815 |

Zero extra flags uses the normal default compiler constants. One thousand extra compiler flags is a custom compiler-environment stress workload, not the normal count of VBA project constants; project constants already use their own live overlay. Small-input costs are mixed, including a slower single-definition index in the stress environment. These results do not establish cold-parser, retained-heap, worker-roundtrip or renderer-latency improvements.
