# Caret selection procedure lookup performance

The tracker cached procedure ranges per document/version, but each selection still searched them linearly. A caret in the last of 1,000 procedures therefore revisited every preceding range on every movement. The tracker now uses an ordered-range lookup with a first-procedure/declarations fast path. The standalone vbaProcedureAtLine helper retains arbitrary-input first-match semantics. No new cache or dependency is added; document edits and document lifetime still use the existing WeakMap.

The ordering precondition comes from vbaProcedureRanges: headers are visited in line order; a header's leading run cannot cross the previous header; each range ends one line before the next range begins. The last range ends at the source's last line. Only these tracker-owned generated ranges use binary lookup. The general helper continues to accept unsorted or overlapping inputs.

## Reproduction

```powershell
node scripts/benchmark-caret-selection-lookup.mjs --baseline=0d769761dc7dd11a0be7383752d680425e9a2088
node scripts/benchmark-caret-selection-lookup.mjs
```

The script bundles the actual tracker and scanner with esbuild, a lightweight event host and stubbed document ownership. Only vbaCaretProcedure.ts is substituted for the baseline. Fixture/source creation, expected-label reference lookup and initial document scan are outside timing. Each sample measures 1,000 synchronous selection events, or 200 in the edited-document control. Positions cycle within the last procedure, jump deterministically across procedures, stay in the first procedure or remain in declarations. Every label is checked after timing against the unchanged general helper. Edited control increments the document version each event, including rescanning in the clock. No range-getter instrumentation runs in timing samples.

v24.18.0, AMD Ryzen 7 9800X3D 8-Core Processor; 15 measured rounds after three warm-ups per row; before/after run sequentially with no other audit benchmarks/tests running. Milliseconds per sample, not per event.

| Procedures | Pattern | Events | Before median | After median | Before p95 | After p95 |
| ---: | --- | ---: | ---: | ---: | ---: | ---: |
| 2 | last | 1000 | 0.170 | 0.093 | 0.234 | 0.293 |
| 2 | jumps | 1000 | 0.094 | 0.098 | 0.220 | 0.182 |
| 2 | first | 1000 | 0.068 | 0.101 | 0.248 | 0.245 |
| 2 | declarations | 1000 | 0.056 | 0.073 | 0.145 | 0.214 |
| 1000 | last | 1000 | 0.684 | 0.107 | 4.642 | 0.156 |
| 1000 | jumps | 1000 | 0.345 | 0.215 | 2.554 | 0.352 |
| 1000 | first | 1000 | 0.031 | 0.043 | 0.150 | 0.124 |
| 1000 | declarations | 1000 | 0.438 | 0.034 | 0.566 | 0.177 |
| 5000 | last | 1000 | 2.814 | 0.076 | 3.042 | 0.242 |
| 5000 | jumps | 1000 | 1.452 | 0.150 | 1.524 | 0.192 |
| 5000 | first | 1000 | 0.030 | 0.030 | 0.339 | 0.063 |
| 5000 | declarations | 1000 | 1.989 | 0.022 | 2.104 | 0.036 |
| 1000 | edited-last | 200 | 26.349 | 25.664 | 30.245 | 30.610 |

A separate actual-tracker range-getter regression counts 400,000 bound reads for 200 selections within the last of 1,000 procedures on the baseline; the changed lookup meets a <=6,000 bound with one source read. Four regressions cover this work bound, reference-label parity across 120 generated sources (three line endings, comments, procedure kinds, every line/half-line boundary, empty modules and NaN/infinities), edit/version invalidation, and unchanged unsorted/overlapping general-helper semantics. The work regression fails on the baseline; the three compatibility controls pass there.

Large late/jump/declaration rows improve. Small-module/first-procedure controls are mixed: two-procedure first median 0.068 -> 0.101 ms and 1,000-procedure first 0.031 -> 0.043 ms. Edited-source work still scans the source on every edit; its p95 is mixed (30.245 -> 30.610 ms). Do not infer a universal speedup or measured VS Code/UI latency from these CPU/event-host samples.
