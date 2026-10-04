# Bare form-control lookup performance

The member-not-found rule searched project metadata for the current UserForm before rejecting a receiver shadowed by a local/parameter, and even when the current form name was absent. These existing rejection checks now run before the bare-control metadata search. Qualified control handling stays in its existing branch. There is no new cache, index, dependency or UI behavior.

## Reproduce

```powershell
node scripts/benchmark-analyzer-form-control-lookup.mjs --baseline=328f46b0 > before.json
node scripts/benchmark-analyzer-form-control-lookup.mjs > after.json
```

Node v24.18.0, Windows, Ryzen 7 9800X3D. Four isolated processes ran sequentially in baseline/fixed/fixed/baseline order. Each fixture has 15 timed rounds after three warm-ups. The full suite and corpus job did not overlap these timings.

## Method and limits

- `rule` runs the actual diagnostic through the real statement walker, with parsing and full-source significant tokens prepared outside the timer and fresh per-pass receiver/member caches. `full` runs complete analyzeModule; `full-fresh` changes a trailing comment to bypass its parser/source cache.
- Each round gets a fresh project metadata list of 5 or 1,000 unrelated document types plus F1, an exhaustive UserForm with T1 and nested-control-shaped T2 entries. Metadata construction and result assertions/hashes are outside the timer. Timings use ordinary data properties, rather than the counted getters used in work regressions.
- The local-form fixture reads declared r As Range in a current form; the host-no-form fixture reads ActiveSheet.Name without a current form. Positive bare controls read T1.Nope with F1 first/last in metadata; qualified controls read f.T1.Nope outside the form. The deliberate missing-control-member diagnostics are compared as complete outputs.
- Every measured round matches its complete fixture result. All 42 configurations have matching complete result hashes across all four processes. These are in-memory analyzer measurements, without Office execution or editor/UI timing.
- Twelve work-counter regressions cover local/parameter shadows, case-insensitive names, nested declarations and host receivers outside a form. For 1,000 references and 1,000 unrelated types, project-kind reads fall from 1,000,000 to zero. Thirteen behavioral controls preserve actual bare/qualified/nested-control diagnostics, parameter/local control-name shadows, valid members, case-insensitive form matching, an explicitly empty form name and mutable control metadata across invocations.

## Results

Times are milliseconds. Each range spans the two process medians or p95s for that implementation. Unrelated types excludes the additional F1 form.

| Fixture | Unrelated types | References | Scope | Baseline median | Fixed median | Baseline p95 | Fixed p95 |
| --- | ---: | ---: | --- | ---: | ---: | ---: | ---: |
| local-form | 5 | 1 | rule | 0.012–0.013 | 0.014–0.015 | 0.036–0.041 | 0.045–0.050 |
| local-form | 5 | 1 | full | 1.180–1.183 | 1.230–1.402 | 1.632–1.705 | 1.900–2.461 |
| local-form | 5 | 1 | full-fresh | 1.092–1.157 | 1.124–1.127 | 1.627–1.676 | 1.562–1.729 |
| local-form | 5 | 1000 | rule | 1.284–1.357 | 1.197–1.228 | 1.926–2.142 | 1.660–1.730 |
| local-form | 5 | 1000 | full | 36.132–36.855 | 35.903–37.410 | 42.622–44.520 | 41.601–43.438 |
| local-form | 5 | 1000 | full-fresh | 38.653–38.829 | 40.222–40.917 | 60.500–63.324 | 53.950–78.983 |
| local-form | 1000 | 1 | rule | 0.048–0.048 | 0.047–0.047 | 0.052–0.062 | 0.051–0.052 |
| local-form | 1000 | 1 | full | 0.806–0.824 | 0.819–0.837 | 0.880–0.897 | 0.901–1.928 |
| local-form | 1000 | 1 | full-fresh | 0.849–0.867 | 0.846–0.865 | 0.922–0.953 | 0.923–2.139 |
| local-form | 1000 | 1000 | rule | 2.054–2.093 | 1.036–1.084 | 2.166–3.371 | 1.115–2.402 |
| local-form | 1000 | 1000 | full | 33.029–33.702 | 31.715–32.022 | 34.822–36.263 | 33.016–34.016 |
| local-form | 1000 | 1000 | full-fresh | 38.327–38.443 | 36.725–37.083 | 40.423–42.313 | 38.614–42.864 |
| host-no-form | 5 | 1 | rule | 0.030–0.031 | 0.029–0.033 | 0.046–0.050 | 0.047–0.050 |
| host-no-form | 5 | 1 | full | 0.662–0.672 | 0.658–0.658 | 0.715–0.900 | 0.711–0.729 |
| host-no-form | 5 | 1 | full-fresh | 0.690–0.696 | 0.670–0.693 | 1.008–1.025 | 1.112–1.142 |
| host-no-form | 5 | 1000 | rule | 2.410–2.454 | 2.342–2.388 | 2.696–3.449 | 3.039–3.169 |
| host-no-form | 5 | 1000 | full | 45.503–45.577 | 44.532–45.586 | 49.658–54.129 | 46.961–53.823 |
| host-no-form | 5 | 1000 | full-fresh | 50.946–51.323 | 50.745–50.782 | 52.081–55.514 | 52.171–78.139 |
| host-no-form | 1000 | 1 | rule | 0.062–0.063 | 0.060–0.062 | 0.070–0.071 | 0.096–0.113 |
| host-no-form | 1000 | 1 | full | 0.788–0.790 | 0.773–0.815 | 0.890–1.953 | 0.876–0.912 |
| host-no-form | 1000 | 1 | full-fresh | 0.787–0.834 | 0.806–0.843 | 0.827–2.474 | 0.833–0.912 |
| host-no-form | 1000 | 1000 | rule | 3.010–3.100 | 2.247–2.261 | 3.678–4.491 | 3.551–3.923 |
| host-no-form | 1000 | 1000 | full | 47.230–48.558 | 46.107–47.562 | 47.658–71.893 | 50.281–55.533 |
| host-no-form | 1000 | 1000 | full-fresh | 54.194–54.417 | 52.086–53.814 | 59.169–64.340 | 61.182–65.648 |
| bare-control-first | 1000 | 1 | rule | 0.054–0.060 | 0.056–0.057 | 0.068–0.115 | 0.069–0.078 |
| bare-control-first | 1000 | 1 | full | 0.736–0.774 | 0.696–0.724 | 0.887–1.030 | 0.798–1.544 |
| bare-control-first | 1000 | 1 | full-fresh | 0.777–0.784 | 0.735–0.756 | 1.072–1.790 | 0.757–0.878 |
| bare-control-first | 1000 | 1000 | rule | 3.003–3.443 | 2.679–2.743 | 4.731–4.828 | 4.496–4.555 |
| bare-control-first | 1000 | 1000 | full | 36.431–38.636 | 34.425–34.601 | 39.653–46.246 | 35.149–35.241 |
| bare-control-first | 1000 | 1000 | full-fresh | 39.355–40.211 | 38.950–39.215 | 40.600–44.718 | 39.536–64.008 |
| bare-control-last | 1000 | 1 | rule | 0.051–0.053 | 0.055–0.057 | 0.056–0.063 | 0.067–0.073 |
| bare-control-last | 1000 | 1 | full | 0.634–0.648 | 0.646–0.697 | 0.682–0.798 | 0.833–1.773 |
| bare-control-last | 1000 | 1 | full-fresh | 0.694–0.984 | 0.662–0.696 | 1.035–1.949 | 0.784–0.816 |
| bare-control-last | 1000 | 1000 | rule | 3.636–3.741 | 3.638–3.709 | 4.825–5.516 | 4.926–6.908 |
| bare-control-last | 1000 | 1000 | full | 35.082–36.648 | 35.744–36.878 | 37.344–52.696 | 38.333–40.104 |
| bare-control-last | 1000 | 1000 | full-fresh | 41.256–41.814 | 40.688–44.219 | 46.953–64.475 | 44.271–55.071 |
| qualified-control | 1000 | 1 | rule | 0.069–0.073 | 0.064–0.067 | 0.078–0.094 | 0.076–0.086 |
| qualified-control | 1000 | 1 | full | 0.727–0.757 | 0.711–0.734 | 0.858–1.036 | 0.758–1.563 |
| qualified-control | 1000 | 1 | full-fresh | 0.762–0.776 | 0.724–0.725 | 0.803–2.072 | 0.771–0.786 |
| qualified-control | 1000 | 1000 | rule | 4.521–4.560 | 3.580–3.605 | 9.646–9.733 | 5.072–5.642 |
| qualified-control | 1000 | 1000 | full | 38.741–40.957 | 38.082–39.253 | 39.865–49.254 | 45.289–48.414 |
| qualified-control | 1000 | 1000 | full-fresh | 45.635–46.790 | 43.908–44.965 | 47.962–51.722 | 46.795–48.904 |

With 1,000 unrelated types and 1,000 local references, rule time falls from 2.054–2.093 ms to 1.036–1.084 ms, while complete cached analysis is 33.029–33.702 ms versus 31.715–32.022 ms and fresh analysis is 38.327–38.443 ms versus 36.725–37.083 ms. The host-no-form rule also improves, but its whole-analyzer timings overlap. Small-project full-fresh local cases are modestly slower in these samples, and positive controls vary. No universal whole-analyzer or UI speedup is claimed.

## Validation

25 new regressions (12 baseline work-count failures, 13 behavioral controls), 216 focused tests, and type checks pass. The full suite on main baseline 328f46b0 plus this fix passes 648 files with 13,193 tests passed and 13 skipped. Complete analyzer outputs matched for 165,120 runs: all 8,256 oracle sources across Excel, Word, PowerPoint and Access, with five project contexts (absent metadata, known form without an active form, active F1, case-varied f1, and an explicitly empty form name). Cached significant statement tokens and arrays were frozen in both bundles; no internal errors occurred. This is semantic parity evidence, not a claim of complete analyzer/repository coverage.
