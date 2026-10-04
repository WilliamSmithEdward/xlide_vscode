# Caret procedure ranges while switching tabs

The status bar and explorer share VbaCaretProcedureTracker. It held ranges for only the active URI/version, so switching between unchanged documents repeatedly read and scanned the entire module. A document reopened with the same URI/version could also reuse a previous document's ranges.

Keep one current-version entry per TextDocument in a WeakMap. Unchanged documents reuse their ranges when tabs switch. A version change replaces only that document's ranges; a newly opened document object gets a fresh entry even when the URI/version repeat. Weak keys do not retain closed document objects. The cache lifetime is the tracker instance. Procedure lookup, event semantics and tree follow/reveal timing remain the same.

## Measurement

Run sequentially from the repository root:

```powershell
node scripts/benchmark-caret-tab-switching.mjs --baseline=88ca7c34
node scripts/benchmark-caret-tab-switching.mjs
```

The harness bundles the actual tracker and range scanner. A small VS Code event host and a stub document-to-module location isolate their work; no real UI rendering, backend reads, ownership decoding, diagnostics or tree reveals are timed. Each sample prepares stable document objects and visits each once outside timing, then times 200 active-editor events and captures their module/procedure answers. Verification and tracker disposal occur afterward. Text-read counters run separately from timing. Three warmup samples precede 15 measured samples; before/after commands run sequentially without other audit workloads.

Small modules have two procedures and 45 CRLF lines; large modules have 1,000 procedures and 22,001 CRLF lines (480,905 bytes). Each procedure contains 20 Debug.Print lines; the caret sits on the final line. The edited control increments a document's version on every activation, requiring a scan in both versions. Total reads include the initial document visits.

Node v24.18.0; AMD Ryzen 7 9800X3D 8-Core Processor; local date 2026-10-03. Times are milliseconds per 200-event batch, not per-switch editor latency.

| Fixture | Documents | Lines per document | Reads before | Reads after | Median before | Median after | p95 before | p95 after |
| --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| single-small | 1 | 45 | 1 | 1 | 0.042 | 0.015 | 0.133 | 0.052 |
| two-small | 2 | 45 | 202 | 2 | 0.387 | 0.017 | 0.500 | 0.028 |
| two-large | 2 | 22001 | 202 | 2 | 153.333 | 0.865 | 163.085 | 1.710 |
| eight-large | 8 | 22001 | 208 | 8 | 151.013 | 0.134 | 156.327 | 0.173 |
| edited-two-large | 2 | 22001 | 202 | 202 | 150.922 | 151.111 | 154.327 | 169.075 |

Unchanged tab alternation eliminates 200 full text reads/scans per batch. The edited control retains all required scans; its median and p95 are slightly worse (150.922 to 151.111 ms, 154.327 to 169.075 ms). These synthetic tracker measurements establish avoided scans, not a universal UI latency improvement. Range lookup still searches procedures and a changed document still scans its full text.

## Validation

Four added regressions cover unchanged tab alternation, an inactive document edited before return, reopen with repeated URI/version, and focus leaving/returning. Three fail on the baseline. The existing caret, procedure ownership, explorer follow, project/folder tree and tab-closing suites pass: 127 tests across six files. Type checking passed. Full suite: 628 files, 12,873 tests passed, 13 skipped. Real VS Code tree integration has not been run for this change. No dependency is added. The expand/collapse, closing-tab and tree redraw audit continues separately.
