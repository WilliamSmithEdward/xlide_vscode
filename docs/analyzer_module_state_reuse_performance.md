# Shared module initial-state facts

Arrays, object-state checks and type inference consult the same module initial-state helpers. Previously every lookup rebuilt the map of module variables that nothing writes, then copied that map for its procedure even when no name was hidden.

Module facts are now shared within a bound-symbol/source/project-write context. Procedure views are prepared once; scopes that hide no module name reuse the module map directly. Facts use weak keys for bound symbols and procedures. A source change refreshes write facts, and recording project writes clears prior module/procedure views, including when project information becomes unavailable. Returned maps are read-only, as their existing API specifies.

Shadowed scopes retain one filtered map per procedure while the bound symbols remain live. This trades temporary repeated map allocation for reuse within that binding lifetime. Normal diagnostic passes build fresh bound symbols, so those caches can be collected after the pass; unshadowed scopes share one module map.

## Measurement

Run `node scripts/benchmark-module-state-reuse.mjs --rounds=15`; compare with `--baseline=312a4959`. The baseline substitutes only the previous module-state implementation. Fixtures contain N private variables and 100 procedures; three lookup rounds model the three diagnostic consumers. The shadowed fixture gives every procedure a parameter that hides one module variable. Every result checks its variable count.

Parsing/binding happen outside timing. Every sample wraps the same bound metadata in a fresh symbols object, so write scanning, module setup and initial scope preparation remain inside timing; only later lookups within that sample reuse facts.

Node 24.18.0, AMD Ryzen 7 9800X3D, three warmups / 15 samples. Values are median milliseconds:

| Module variables | Unshadowed before | Unshadowed after | Shadowed before | Shadowed after |
| --- | ---: | ---: | ---: | ---: |
| 100 | 1.768 | 0.124 | 1.816 | 0.575 |
| 1,000 | 18.139 | 0.498 | 18.945 | 4.456 |
| 3,000 | 67.876 | 0.572 | 73.127 | 13.393 |

The operation-count fixture with 200 variables, 20 procedures and three consumers reads module symbol names 12,000 times before vs 200 after. This is a helper-level stress workload, not end-to-end editor latency.

## Validation

1,000 generated bound modules preserve complete repeated state-map entries and iteration order, covering private/public variables, writes, auto-instantiation, fixed-length strings, arrays, project information and local/parameter shadowing. New regressions cover the read budget, repeated scope results, changed source, project-write updates, same-set updates and unavailable project information.

Validation completed: 40 targeted tests and the type check passed; full suite 586 files / 12,346 tests passed (13 skipped).
