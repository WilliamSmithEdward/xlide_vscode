# Expression procedure lookup performance

resolveExpressionType already caches parsed/bound modules, module signatures and procedure type environments. Its enclosing-procedure lookup nevertheless scanned every preceding module member for each selected span. A bound-module-owned index now serves repeated queries. The first query retains the original scan. The second collects procedure intervals and verifies that they are ordered and non-overlapping; malformed intervals keep the original first-match scan. Valid intervals use a binary search for the first end at or after the selection, retaining inclusive containment semantics. Index lifetime follows the existing bounded source/module-name/module-kind binding cache; no new global cache is introduced.

Four public-API regression tests include two baseline work failures and two behavior controls for overlapping intervals and source/module isolation. Twenty-nine focused resolver/Extract Variable tests and type checking passed. The full suite on main 21ef8cb1 passed 662 files and 13,357 tests (14 skipped). The overlap control supplies deliberately overlapping parsed intervals and verifies the first procedure's local type wins even after repeated queries. Complete baseline/fixed resolver results match in 501,085 comparisons across 8,256 corpus sources and 432 generated sources. Selections cover whole sources, all significant tokens, and nearby partial/boundary spans. Generated procedures include mismatched/missing terminators, module declarations, local scalar/object types and LF/CRLF/CR. Bound ASTs were frozen after the initial public query; no matching or unmatched exceptions occurred. These are pure resolver results, not Office execution.

Stable member-kind getter observations on frozen parsed procedure nodes show the old warm lookup performs 100,000 kind reads for 100 repeated queries to the final procedure of a 1,000-procedure module. The fixed index construction still reads all 1,000 kinds once, then the already-indexed queries read none. This counts the removed member-kind scan only: binary-search endpoint comparisons, tokenization, type inference and other work remain. One-shot and second-query construction costs are explicitly included in timing controls below.

Run scripts/benchmark-expression-procedures.mjs --baseline=21ef8cb1 and then without --baseline. Four separate processes ran A/B/B/A before this audit's full suite; all 36 complete-result hashes matched. The actual public resolver is timed for independent whole-number literal expectations, selecting first/last/spread procedures in modules of 1/100/1,000 procedures, with 1 or 100 queries, and warm or fresh binding. Three warmups and fifteen rounds. Fresh source comments change binding identity outside the timed source construction, so fresh timings include parsing, symbol binding and index construction where queried twice. Source/span/context generation and complete output assertions/hashing are outside timing. No IO, UI, Office or entire analyzer pass is measured. Node v24.18.0, shared Windows host. Scalar-literal timings do not establish the performance of every expression or context. All cases are reported, including small/first-procedure and single-query controls that can be mixed or slower.

| Procedures | Target | Queries | Binding | Baseline median ms | Fixed median ms | Baseline p95 ms | Fixed p95 ms |
| ---: | --- | ---: | --- | ---: | ---: | ---: | ---: |
| 1 | first | 1 | warm-binding | 0.0107–0.0133 | 0.0119–0.0125 | 0.0221–0.0225 | 0.0221–0.0242 |
| 1 | first | 1 | fresh-binding | 0.0529–0.0550 | 0.0526–0.0591 | 0.0972–0.1329 | 0.0952–0.0996 |
| 1 | first | 100 | warm-binding | 0.1317–0.1770 | 0.1343–0.1479 | 0.4820–0.7607 | 0.4851–0.5691 |
| 1 | first | 100 | fresh-binding | 0.1161–0.1951 | 0.1208–0.1252 | 0.3421–1.5560 | 0.2179–0.3821 |
| 1 | last | 1 | warm-binding | 0.0009–0.0011 | 0.0011–0.0011 | 0.0010–1.4229 | 0.0029–1.3719 |
| 1 | last | 1 | fresh-binding | 0.0132–0.0134 | 0.0137–0.0148 | 0.0260–0.0344 | 0.0299–1.6458 |
| 1 | last | 100 | warm-binding | 0.0819–0.0942 | 0.0812–0.0979 | 0.1057–0.1238 | 0.1117–0.1375 |
| 1 | last | 100 | fresh-binding | 0.0753–0.0795 | 0.0907–0.1093 | 0.2625–0.2762 | 0.2198–0.3340 |
| 1 | spread | 1 | warm-binding | 0.0007–0.0008 | 0.0007–0.0008 | 0.0030–0.0033 | 0.0008–0.0008 |
| 1 | spread | 1 | fresh-binding | 0.0103–0.0141 | 0.0105–0.0107 | 0.0166–0.0291 | 0.0124–0.0164 |
| 1 | spread | 100 | warm-binding | 0.0866–0.0925 | 0.0868–0.0896 | 0.1207–0.4239 | 0.3229–0.3413 |
| 1 | spread | 100 | fresh-binding | 0.0852–0.0935 | 0.0801–0.0848 | 0.1735–0.2834 | 0.1249–0.1526 |
| 100 | first | 1 | warm-binding | 0.0007–0.0007 | 0.0007–0.0007 | 0.0007–0.0007 | 0.0087–0.2912 |
| 100 | first | 1 | fresh-binding | 0.4551–0.5048 | 0.4488–0.4716 | 1.2713–1.8152 | 1.2495–2.2185 |
| 100 | first | 100 | warm-binding | 0.0603–0.0750 | 0.0743–0.0783 | 0.7500–0.9422 | 0.5720–0.6462 |
| 100 | first | 100 | fresh-binding | 0.3455–0.3686 | 0.3663–0.4119 | 1.1716–1.3237 | 1.1540–1.2904 |
| 100 | last | 1 | warm-binding | 0.0007–0.0012 | 0.0007–0.0008 | 0.0008–0.0015 | 0.0008–0.0008 |
| 100 | last | 1 | fresh-binding | 0.2909–0.3147 | 0.2901–0.2945 | 1.2618–1.3464 | 1.1859–1.2379 |
| 100 | last | 100 | warm-binding | 0.0550–0.0771 | 0.0434–0.0715 | 0.0840–0.0840 | 0.0552–0.0988 |
| 100 | last | 100 | fresh-binding | 0.5527–0.5697 | 0.5274–0.5382 | 1.0851–1.2149 | 0.9415–1.1581 |
| 100 | spread | 1 | warm-binding | 0.0006–0.0007 | 0.0007–0.0007 | 0.0009–0.0016 | 0.0010–0.0029 |
| 100 | spread | 1 | fresh-binding | 0.2971–0.3302 | 0.3151–0.4505 | 1.0010–1.1004 | 0.7817–1.2392 |
| 100 | spread | 100 | warm-binding | 0.0501–0.0516 | 0.0464–0.0481 | 0.0553–0.0565 | 0.0543–0.0605 |
| 100 | spread | 100 | fresh-binding | 0.4032–0.4229 | 0.4152–0.5391 | 0.6899–1.3952 | 1.0156–2.4088 |
| 1000 | first | 1 | warm-binding | 0.0006–0.0010 | 0.0006–0.0006 | 0.0006–0.0012 | 0.0007–0.0007 |
| 1000 | first | 1 | fresh-binding | 3.8272–5.8712 | 3.8866–4.2364 | 10.1895–10.2069 | 6.6421–8.2460 |
| 1000 | first | 100 | warm-binding | 0.0432–0.0433 | 0.0442–0.0467 | 0.0443–0.0617 | 0.0451–0.0815 |
| 1000 | first | 100 | fresh-binding | 2.5281–2.9631 | 2.4404–2.5002 | 6.0316–6.6842 | 5.4693–5.7000 |
| 1000 | last | 1 | warm-binding | 0.0019–0.0026 | 0.0006–0.0006 | 0.0022–0.0042 | 0.0009–0.0010 |
| 1000 | last | 1 | fresh-binding | 2.4883–2.4989 | 2.3897–2.4558 | 6.0907–6.5804 | 5.6805–6.2180 |
| 1000 | last | 100 | warm-binding | 0.1720–0.1825 | 0.0433–0.0470 | 0.1764–0.2098 | 0.0484–3.1084 |
| 1000 | last | 100 | fresh-binding | 2.7091–2.7362 | 2.4442–2.5028 | 6.6275–7.9213 | 5.8215–6.9641 |
| 1000 | spread | 1 | warm-binding | 0.0006–0.0006 | 0.0006–0.0006 | 0.0007–0.0007 | 0.0006–0.0007 |
| 1000 | spread | 1 | fresh-binding | 2.3725–2.4929 | 2.4277–2.5328 | 5.2816–9.7002 | 7.8832–7.9395 |
| 1000 | spread | 100 | warm-binding | 0.0535–0.0543 | 0.0483–0.0519 | 0.0549–0.0635 | 0.0549–0.0612 |
| 1000 | spread | 100 | fresh-binding | 2.5363–2.5449 | 2.4709–2.8507 | 5.1278–5.4547 | 5.3700–7.3607 |
