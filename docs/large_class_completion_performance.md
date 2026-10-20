# ROneCOne large-class completion latency

The owner reported a visible delay while editing the ROneCOne class in ROneCOne_Delegates_Demo.xlsm. The actual workbook supplied through XLIDE_PERF_WORKBOOK has a 969,296-character class; the copied probe document has about 26,447 lines. Testing always reads the original and edits a disposable copy in the harness workspace. No original source or workbook is committed.

## Reproduce

Set `XLIDE_PERF_WORKBOOK` to the workbook path. Run `npm run compile`, then `npx vscode-test --code-version 1.139.1 --grep "Actual large class latency"`. The suite appends a probe procedure to the class in a copied workbook, types five actual characters and checks completion after each edit, then checks hover before and after changing a local declaration from Long to Double. It records edit-to-result timings and the extension's performance snapshot. Set `XLIDE_PERF_CPU_PROFILE=1` to additionally save an extension-host CPU profile to the disposable workspace. The clipboard is restored after reading the performance snapshot.

For analyzer phases run `npx vitest run tests/roneconeLatencyProbe.test.ts` with XLIDE_PERF_WORKBOOK set. The opt-in fresh-source probe reports lexer, parser, symbol, index, context and member costs separately. The keyword benchmark warms six different source/caret positions, then reports the median per-request time across 21 batches of six requests. Ordinary CI skips these workbook-specific probes.

## Findings and change

The first real editor run recorded five edit-to-completion samples of 679–1,252 ms. A later CPU-profiled run after startup recorded 154–499 ms, while the completion provider's own traces were 105–134 ms and semantic-token requests took 61–94 ms. These runs have different startup and profiling conditions; they are not a controlled speedup comparison.

The profile attributed about 173 ms across five requests to keyword completion's closing/branch helpers. Both scanned the entire class before filtering suggestions against an ordinary identifier prefix. Gate scans against the actual closer/branch catalog, share a single scan when those rows can match, and retain four enclosing-prefix answers across partial-word edits. A change before the statement chooses a different cache key. Compact dynamic closers such as Nextitem remain eligible.

The editor's external-symbol lists also queried the current module and discarded its rows. Add an external-only visibility option so those two queries skip that module before projecting its signatures/symbols. Other consumers retain full visibility, and qualified current-module surfaces remain available. The isolated fresh-projection benchmark showed no material overall timing improvement (1.5523 vs 1.5414 ms); this is a work elimination, not a claimed class-response speedup.

Compared with main at 29f8a231 (which already includes prefix windows), the matched actual-class keyword benchmark improved from 33.0141 to 0.1760 ms per request; slowest batch averages were 39.40085 and 0.30255 ms. Source lexing was warm, positions varied, and no other harness or unit run competed during this comparison. This measures a component, not full response or painting latency.

## Validation and remaining delay

Compilation passed. The full unit suite passed: 685 test files passed, one intentionally skipped; 13,632 tests passed and 25 skipped. Work-count regressions prove zero block scans for impossible prefixes, one shared scan for a matching request, reuse across partial-word edits, invalidation on earlier block changes, and no current-module declaration reads in the two external-list projections. Existing keyword, visibility, scope and designer-control behavior remains covered.

Seventeen real VS Code 1.139.1 checks passed, including the actual class completion/hover probes and the existing editor, immediate completion, prefix-window, type-lookback and hover-snapshot suites. The final class run's five edit-to-completion samples were 152.35, 167.21, 96.41, 79.99 and 158.29 ms. Warm hover took 6.15 ms; edit-to-hover took 240.35 ms. These observations include other editor/background activity and do not establish an isolated before/after speedup for the full response.

Issue #985 remains open: full lexing/parsing after edits and semantic-token/background work still leave noticeable delay. The fresh-source component probe measured lexing at roughly 27–68 ms and parsing at 37–58 ms. Menu and tooltip painting are not measured by these command timings; this patch does not claim nearly instantaneous behavior in the large class.
