# Unstructured flow existence checks

Dataflow preparation needs one boolean: whether a procedure contains a label, label reference, On Error or Resume that requires conservative control-flow handling. The original first computation built complete label-reference and declaration arrays in separate full-body walks, then performed another walk for error/resume statements without label targets.

The existence check now walks active bodies once and stops at the first witness. A statement-level predicate shares one cached token lookup and the normalized first token across the existing label and error detectors. Ordinary multi-token statements that cannot declare a named label avoid constructing a label object. The procedure/source/activity cache keys and recognized flow cases are preserved.

Label collection and diagnostics also share one dependency-light statement walker in `parser/statementWalk.ts`. The old diagnostics export is retained, and the private duplicate label walker is removed. This avoids a flow-to-diagnostics dependency cycle. This extraction has no separate speedup claim.

## Measurement

Run `node scripts/benchmark-flow-existence.mjs`, or add `--baseline=9091f74d` for the original flow and label modules. Three warmups and 15 samples, Node 24.18.0, AMD Ryzen 7 9800X3D. Source construction and parsing excluded; statement-token caches are warm. Each first-computation sample creates a fresh shallow procedure wrapper to bypass the flow result cache, with the original body retained. Milliseconds, median / p95; zero rounded medians are shown as less than 0.001:

| Assignments / witness | Original | Optimized |
| --- | --- | --- |
| 100-none | 0.037 / 0.421 | 0.023 / 0.081 |
| 100-early-reference | 0.021 / 0.049 | 0.004 / 0.008 |
| 100-early-label | 0.053 / 0.077 | 0.001 / 0.004 |
| 100-early-error | 0.025 / 0.047 | <0.001 / 0.001 |
| 100-late-reference | 0.014 / 0.046 | 0.024 / 0.039 |
| 1000-none | 0.275 / 0.568 | 0.123 / 0.354 |
| 1000-early-reference | 0.067 / 0.615 | 0.001 / 0.001 |
| 1000-early-label | 0.134 / 0.182 | 0.001 / 0.001 |
| 1000-early-error | 0.169 / 0.199 | 0.001 / 0.002 |
| 1000-late-reference | 0.065 / 0.093 | 0.147 / 1.627 |
| 10000-none | 2.451 / 4.356 | 0.729 / 2.795 |
| 10000-early-reference | 0.768 / 3.694 | 0.001 / 0.002 |
| 10000-early-label | 1.248 / 2.305 | 0.001 / 0.001 |
| 10000-early-error | 1.283 / 4.238 | <0.001 / 0.001 |
| 10000-late-reference | 0.717 / 0.794 | 0.685 / 1.293 |

The no-flow and late-reference fixtures must examine the whole procedure and are included as controls. Early-reference, early-label and early-error fixtures have their witness before the assignments. These figures apply to a first flow-fact computation, not complete module analysis.

Cached-hit controls make 1,000 calls using the same already-primed procedure. They retain the same cache algorithm and keys; these rows exercise the cached return path:

| 10,000-assignment fixture / 1,000 cached calls | Original | Optimized |
| --- | --- | --- |
| 10000-none-1000-cached-calls | 0.005 / 0.005 | 0.005 / 0.005 |
| 10000-early-reference-1000-cached-calls | 0.005 / 0.015 | 0.01 / 0.011 |
| 10000-early-label-1000-cached-calls | 0.022 / 0.05 | 0.005 / 0.005 |
| 10000-early-error-1000-cached-calls | 0.014 / 0.041 | 0.022 / 0.029 |
| 10000-late-reference-1000-cached-calls | 0.005 / 0.005 | 0.014 / 0.037 |

## Validation

Structural regressions bound span reads for early jumps, labels, error handlers and Resume Next; all four fail on the original full scans. Further fixtures cover nested witnesses, conditional inactive branches, source/activity cache invalidation, cached hits, and label order. Exact flow facts, label declarations/references and shared visitor span sequences match the original for 500 generated procedures under both conditional configurations (1,000 comparisons).
