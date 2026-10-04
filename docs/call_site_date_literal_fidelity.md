# Preserve date literals when collecting call sites

Issue #983. Baseline: 413f9cef (the #960 fix).

Introduce Parameter on Public Sub May previously treated May inside Debug.Print #May 1, 2000# as a call. Its generated edit changed the date to #May(3) 1, 2000#, even though the actual May statement was a separate call. All twelve full English month names reproduce this defect.

Hash-bearing logical lines now use the existing query-owned lexer index. The lexer recognizes complete dateLiteral tokens; tokenName rejects them as callees. File-number markers still retain real calls, and colon-separated statements retain their previous boundaries. This uses the existing date grammar and index rather than adding a second date parser or fragment cache. No signature or project binding policy changes here.

## Validation

The new date test file has 45 cases across LF, CRLF and CR, including all twelve month names, a real call on the same physical line, file-number syntax and complete Introduce Parameter output in primary/external modules. The baseline fails 39 and passes six controls. Three additional work guards at 1, 100 and 1,000 date/call pairs also fail on the baseline: they require exactly the real calls, at most one lex per date line, bounded fragment input, preserved module token-cache identity and exact rendered output. Across both files, the baseline fails 42 and passes 13 controls (seven are existing work tests).

All 202 focused tests pass. Type checking passes. The complete suite passes 687 files and 13,783 tests, with 22 skipped.

With owned and cached lexer tokens/trivia frozen, 19,814 complete queries over 8,256 corpus sources have identical metadata arrays and rendered source against the baseline, with no overlapping candidate edits or uncaught exceptions. All 216 generated complete refactors match the baseline and independently expected primary/external text. These compatible corpus cases do not expose the month-name defect; the dedicated independently expected regressions above do. No Office runtime execution is claimed.

## Cost measurement

The existing benchmark adds --date-literals, retaining its original 36 configurations without duplicating the harness. Candidate output and call count are enforced. Baseline date-mode counts are recorded rather than rejected because the defect creates extra call sites. All eighteen candidate configurations render correctly; twelve baseline configurations do. The original 36-configuration mode also completes with correct candidate output.

Commands: node scripts/benchmark-call-site-arguments.mjs --date-literals --baseline=413f9cef --rounds=9 and the same command without --baseline. Runs are sequential baseline/candidate/candidate/baseline, with three warmups and nine measured samples. Node 24.18.0, Windows, Ryzen 7 9800X3D. Each table cell lists the two median milliseconds; only call-site collection is timed, followed by complete expected-render checks. Fresh inputs vary a leading comment to miss the full-source stripped-text cache; this is not a claim that every cache is cold.

Correct date handling incurs additional bounded lexing. At 1,000 date/call pairs, cached time rises from 1.17–1.56 to 1.91–1.96 ms, while the erroneous 2,000 call sites become the correct 1,000. File-number calls also gain lexer cost: 1,000 cached calls rise from 0.60–0.69 to 1.31–1.54 ms. Ordinary numeric controls are mixed at these small timings. This is a correctness repair with documented costs, not a speedup claim.

| calls/layout/cache | before ms | after ms | before / after sites | before / after correct |
| --- | --- | --- | --- | --- |
| 1/month/cached | 0.0067 / 0.0068 | 0.0157 / 0.0157 | 2 / 1 | no / yes |
| 1/month/fresh | 0.0095 / 0.0088 | 0.012 / 0.012 | 2 / 1 | no / yes |
| 1/fileNumber/cached | 0.0024 / 0.0025 | 0.0107 / 0.0117 | 1 / 1 | yes / yes |
| 1/fileNumber/fresh | 0.0058 / 0.0041 | 0.0086 / 0.0088 | 1 / 1 | yes / yes |
| 1/number/cached | 0.0017 / 0.0017 | 0.0026 / 0.0025 | 1 / 1 | yes / yes |
| 1/number/fresh | 0.0048 / 0.0032 | 0.0034 / 0.0036 | 1 / 1 | yes / yes |
| 100/month/cached | 0.1889 / 0.192 | 0.298 / 0.2901 | 200 / 100 | no / yes |
| 100/month/fresh | 0.2575 / 0.1484 | 0.2584 / 0.251 | 200 / 100 | no / yes |
| 100/fileNumber/cached | 0.1188 / 0.0662 | 0.2264 / 0.2236 | 100 / 100 | yes / yes |
| 100/fileNumber/fresh | 0.105 / 0.0997 | 0.2483 / 0.2433 | 100 / 100 | yes / yes |
| 100/number/cached | 0.023 / 0.023 | 0.0237 / 0.0221 | 100 / 100 | yes / yes |
| 100/number/fresh | 0.0345 / 0.0344 | 0.0346 / 0.0333 | 100 / 100 | yes / yes |
| 1000/month/cached | 1.5642 / 1.1746 | 1.9068 / 1.9619 | 2000 / 1000 | no / yes |
| 1000/month/fresh | 1.5377 / 1.3872 | 2.3444 / 2.0935 | 2000 / 1000 | no / yes |
| 1000/fileNumber/cached | 0.6898 / 0.6006 | 1.5383 / 1.3088 | 1000 / 1000 | yes / yes |
| 1000/fileNumber/fresh | 0.8966 / 1.3598 | 2.4615 / 1.5331 | 1000 / 1000 | yes / yes |
| 1000/number/cached | 0.1834 / 0.3636 | 0.1894 / 0.1755 | 1000 / 1000 | yes / yes |
| 1000/number/fresh | 0.2845 / 0.5305 | 0.5257 / 0.2779 | 1000 / 1000 | yes / yes |
