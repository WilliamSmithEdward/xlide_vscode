# Private inference reader source adoption

`callEffectsFor` caches the callback that describes ByRef call effects within
bound symbols, source and conditional activity. `declaredFactsFor` caches type,
fixed-bound and constant readers within bound symbols, procedure and source.
Their equal-source hits retained the older source string and could compare the
whole module on every later reader lookup. Matching hits now adopt the caller
string while returning the same callback/fact object. Every existing key,
closure context, lazily computed answer and invalidation rule is unchanged.

## Reproduction and measured scope

Run `node scripts/benchmark-inference-reader-source.mjs`, or add
`--baseline=9c9cc784d827e094e3b40b0fb1d40c0df152d960` for the prior implementation.
The private helpers are exposed only in the temporary benchmark bundle.
Both readers are measured at approximately 1 KB, 100 KB and 1 MB using original
strings and equal copies. Three warmups precede nine measured rounds, each with
fresh source and bound symbols. Independent assertions check ByRef effect
values, declared types, fixed bounds, constants and shadowed names before timing;
identity is checked on every cache hit. Construction, parsing, symbol binding,
priming and callback evaluation are excluded.

Paired before/after/after/before runs on Node 24.18.0 / Ryzen 7 9800X3D:

| 1 MB source, 1,000 warm reader lookups | Before median ms | After median ms |
| --- | --- | --- |
| Equal copy, effects | 16.7690–16.7945 | 0.0495–0.0512 |
| Equal copy, declarations | 17.2410–17.8872 | 0.0534–0.0536 |
| Same string, effects | 0.0201–0.0349 | 0.0206–0.0213 |
| Same string, declarations | 0.0219–0.0293 | 0.0216–0.0217 |

Same-string controls are mixed. The first equal-copy comparison remains necessary,
and fresh strings on every call may still incur it. These are reader-retrieval
measurements, not callback evaluation, complete diagnostics or editor latency.

## Validation

The linked PR records type checking, focused/full suites and frozen baseline
comparison counts. Complete callback answers and diagnostic arrays are compared
across the syntax corpus and generated LF/CRLF/CR, Option Base and ByRef/ByVal/
Optional call cases. Independent source A/B/A controls verify changed ByRef
results and module constants while preserving callback identity on equal copies.

Type checking and 21 focused tests pass. The frozen baseline differential
includes 8,256 corpus sources plus 18 generated cases: 13,452 diagnostics,
no internal errors and identical complete callback answers and diagnostic arrays.

The full suite passes: 737 files and 14,535 tests (7 files and 33 tests skipped).
