# Statement-ending call parenthesis performance

The empty-parentheses and multi-argument standalone-call helpers share a generator that used to match a parenthesis by rescanning its suffix for every candidate identifier. Rejected nested calls and incomplete input repeated overlapping scans. Matching pairs are now indexed locally on the second lookup. The first lookup still uses the original scan, preserving the allocation-free shallow-call path. The existing complete-receiver-chain checks remain in force. There is no persistent cache or mutation of caller tokens.

Deterministic rawText getter observations through the actual cached statement-token input (all tokens and the array frozen): 1,000 nested F calls (3,001 tokens) fell from 1,506,503 to 11,004 reads for the empty-call consumer, and from 1,509,501 to 14,002 for the multi-argument consumer. Unclosed input (2,001 tokens) fell from 1,005,002 to 8,003 in both consumers. This measures downstream rawText reads, excluding lexing and other token fields; it is not a count of all CPU instructions. Counts at 10, 100 and 1,000 depths are in the audit evidence.

Thirteen regression tests include four baseline work failures and nine passing behavior controls. Sixty-nine focused call/diagnostic tests and type checking passed. The final full suite on main 8ebe43b9 passed 658 files and 13,320 tests (13 skipped). Against baseline 46188294, 163,720 complete results from both public consumers match: 78,710 corpus spans (whole-source and physical-line spans) and 3,150 generated statements. Generated cases cover leading dots, numbered lines, explicit Call, assignments, nested calls, strings, comments, bracketed names, malformed parentheses and receiver-chain suffixes in LF, CRLF and CR contexts. Cached token objects and arrays were frozen before each comparison. These are pure source-text API comparisons, not Office execution claims.

Run node scripts/benchmark-call-parentheses.mjs --baseline=46188294 for baseline and omit the flag for the working tree. Four separate processes ran A/B/B/A, before this audit's full suite. All 64 complete-output hashes matched and every invocation was checked against an independently constructed expected complete result. Timers cover the actual public consumer; source/span generation, output assertions and hashes are outside. Cached tokens are frozen before timing. Fresh-source comments change cache identity outside the statement span, so fresh timings include lexing that statement. Three warmups and fifteen rounds per configuration; cached shallow cases use batches of 100. No file IO, Office execution, entire analyzer pass or end-user latency is measured. Node v24.18.0; shared Windows host, so absolute times remain host-dependent.

The depth refers to inner G calls inside F. Receiver inputs end in .Done(), and multi inputs give F two arguments. Module diagnostic speedups cannot be inferred directly from these helper timings. Small and already-cheap cases can be mixed or slower; the complete table includes these controls.

| Depth | Shape | Consumer | Tokens | Baseline median ms | Fixed median ms | Baseline p95 ms | Fixed p95 ms |
| ---: | --- | --- | --- | ---: | ---: | ---: | ---: |
| 0 | nested | empty | cached-tokens | 0.0003–0.0003 | 0.0003–0.0003 | 0.0011–0.0013 | 0.0016–0.0030 |
| 0 | nested | empty | fresh-tokens | 0.0049–0.0051 | 0.0058–0.0060 | 0.0116–0.0116 | 0.0116–0.0125 |
| 0 | nested | multi | cached-tokens | 0.0004–0.0004 | 0.0004–0.0004 | 0.0026–0.0028 | 0.0021–0.0024 |
| 0 | nested | multi | fresh-tokens | 0.0032–0.0036 | 0.0034–0.0035 | 0.0157–0.0192 | 0.0168–0.0188 |
| 0 | unclosed | empty | cached-tokens | 0.0001–0.0001 | 0.0001–0.0001 | 0.0009–0.0009 | 0.0009–0.0012 |
| 0 | unclosed | empty | fresh-tokens | 0.0026–0.0026 | 0.0027–0.0043 | 0.0067–0.0069 | 0.0074–0.0102 |
| 0 | unclosed | multi | cached-tokens | 0.0000–0.0001 | 0.0000–0.0000 | 0.0001–0.0001 | 0.0001–0.0001 |
| 0 | unclosed | multi | fresh-tokens | 0.0025–0.0026 | 0.0031–0.0031 | 0.0064–0.0065 | 0.0074–0.0212 |
| 0 | receiver | empty | cached-tokens | 0.0004–0.0005 | 0.0005–0.0005 | 0.0021–0.0021 | 0.0021–0.0022 |
| 0 | receiver | empty | fresh-tokens | 0.0047–0.0049 | 0.0049–0.0049 | 0.0846–0.1185 | 0.0138–0.0959 |
| 0 | receiver | multi | cached-tokens | 0.0002–0.0002 | 0.0002–0.0003 | 0.0003–0.0004 | 0.0012–0.0015 |
| 0 | receiver | multi | fresh-tokens | 0.0043–0.0043 | 0.0042–0.0043 | 0.0143–0.0146 | 0.0149–0.0164 |
| 0 | multi | empty | cached-tokens | 0.0003–0.0003 | 0.0002–0.0002 | 0.0012–0.0015 | 0.0004–0.0007 |
| 0 | multi | empty | fresh-tokens | 0.0034–0.0035 | 0.0033–0.0033 | 0.0037–0.0146 | 0.0876–0.1218 |
| 0 | multi | multi | cached-tokens | 0.0002–0.0002 | 0.0002–0.0002 | 0.0007–0.0010 | 0.0005–0.0010 |
| 0 | multi | multi | fresh-tokens | 0.0033–0.0036 | 0.0032–0.0032 | 0.0139–0.0305 | 0.0110–0.0127 |
| 5 | nested | empty | cached-tokens | 0.0003–0.0003 | 0.0005–0.0005 | 0.0012–0.0013 | 0.0013–0.0013 |
| 5 | nested | empty | fresh-tokens | 0.0060–0.0065 | 0.0060–0.0060 | 0.0096–0.0112 | 0.0120–0.0140 |
| 5 | nested | multi | cached-tokens | 0.0003–0.0004 | 0.0006–0.0006 | 0.0008–0.0010 | 0.0014–0.0019 |
| 5 | nested | multi | fresh-tokens | 0.0034–0.0059 | 0.0036–0.0038 | 0.0041–0.0067 | 0.0039–0.0040 |
| 5 | unclosed | empty | cached-tokens | 0.0002–0.0002 | 0.0002–0.0003 | 0.0010–0.0011 | 0.0004–0.0004 |
| 5 | unclosed | empty | fresh-tokens | 0.0021–0.0021 | 0.0023–0.0023 | 0.0022–0.0466 | 0.0024–0.0045 |
| 5 | unclosed | multi | cached-tokens | 0.0002–0.0002 | 0.0002–0.0002 | 0.0002–0.0008 | 0.0007–0.0012 |
| 5 | unclosed | multi | fresh-tokens | 0.0021–0.0021 | 0.0023–0.0023 | 0.0069–0.0081 | 0.0024–0.0027 |
| 5 | receiver | empty | cached-tokens | 0.0003–0.0003 | 0.0005–0.0005 | 0.0010–0.0011 | 0.0029–0.0030 |
| 5 | receiver | empty | fresh-tokens | 0.0034–0.0034 | 0.0037–0.0038 | 0.0035–0.0035 | 0.0041–0.0082 |
| 5 | receiver | multi | cached-tokens | 0.0003–0.0003 | 0.0005–0.0005 | 0.0004–0.0005 | 0.0011–0.0011 |
| 5 | receiver | multi | fresh-tokens | 0.0034–0.0034 | 0.0037–0.0037 | 0.0036–0.0038 | 0.0490–0.0816 |
| 5 | multi | empty | cached-tokens | 0.0003–0.0003 | 0.0005–0.0005 | 0.0006–0.0006 | 0.0011–0.0015 |
| 5 | multi | empty | fresh-tokens | 0.0034–0.0034 | 0.0042–0.0043 | 0.0052–0.0052 | 0.0059–0.0064 |
| 5 | multi | multi | cached-tokens | 0.0002–0.0002 | 0.0003–0.0003 | 0.0008–0.0008 | 0.0004–0.0005 |
| 5 | multi | multi | fresh-tokens | 0.0033–0.0033 | 0.0035–0.0035 | 0.0033–0.0114 | 0.0037–0.0158 |
| 100 | nested | empty | cached-tokens | 0.0260–0.0261 | 0.0057–0.0058 | 0.0303–0.0326 | 0.0098–0.0203 |
| 100 | nested | empty | fresh-tokens | 0.0520–0.0522 | 0.0327–0.0344 | 0.1312–0.1527 | 0.1195–0.1278 |
| 100 | nested | multi | cached-tokens | 0.0265–0.0266 | 0.0062–0.0062 | 0.0267–0.0519 | 0.0114–0.0139 |
| 100 | nested | multi | fresh-tokens | 0.0532–0.0535 | 0.0370–0.0374 | 0.1648–0.2848 | 0.1539–0.2144 |
| 100 | unclosed | empty | cached-tokens | 0.0122–0.0122 | 0.0031–0.0031 | 0.0123–0.0170 | 0.0032–0.0117 |
| 100 | unclosed | empty | fresh-tokens | 0.0289–0.0290 | 0.0207–0.0213 | 0.1002–0.1005 | 0.0965–0.1263 |
| 100 | unclosed | multi | cached-tokens | 0.0122–0.0122 | 0.0031–0.0031 | 0.0186–0.0205 | 0.0031–0.0031 |
| 100 | unclosed | multi | fresh-tokens | 0.0292–0.0294 | 0.0211–0.0221 | 0.0438–0.1067 | 0.1978–0.2861 |
| 100 | receiver | empty | cached-tokens | 0.0196–0.0196 | 0.0064–0.0064 | 0.0286–0.0292 | 0.0066–0.0185 |
| 100 | receiver | empty | fresh-tokens | 0.0448–0.0457 | 0.0344–0.0351 | 0.2611–0.3343 | 0.2066–0.2915 |
| 100 | receiver | multi | cached-tokens | 0.0195–0.0195 | 0.0086–0.0086 | 0.0196–0.0198 | 0.0378–0.1016 |
| 100 | receiver | multi | fresh-tokens | 0.0429–0.0451 | 0.0498–0.0598 | 0.2352–0.2484 | 0.1795–0.2085 |
| 100 | multi | empty | cached-tokens | 0.0194–0.0199 | 0.0046–0.0070 | 0.0270–0.0277 | 0.0049–0.0080 |
| 100 | multi | empty | fresh-tokens | 0.0599–0.0604 | 0.0579–0.0587 | 0.1542–0.2381 | 0.0746–0.0973 |
| 100 | multi | multi | cached-tokens | 0.0020–0.0020 | 0.0021–0.0022 | 0.0021–0.0027 | 0.0022–0.0022 |
| 100 | multi | multi | fresh-tokens | 0.0423–0.0457 | 0.0272–0.0274 | 0.0517–0.0589 | 0.0276–0.1321 |
| 1000 | nested | empty | cached-tokens | 1.6139–1.6139 | 0.0467–0.0469 | 1.6860–1.7002 | 0.0646–0.0651 |
| 1000 | nested | empty | fresh-tokens | 1.9646–1.9846 | 0.3872–0.4070 | 2.4968–2.7952 | 0.8616–1.1004 |
| 1000 | nested | multi | cached-tokens | 1.6231–1.6460 | 0.0497–0.0499 | 1.7825–1.9023 | 0.0537–0.0565 |
| 1000 | nested | multi | fresh-tokens | 1.8121–1.8415 | 0.3081–0.3983 | 2.5904–2.5939 | 0.9407–1.0758 |
| 1000 | unclosed | empty | cached-tokens | 1.0535–1.0602 | 0.0232–0.0234 | 1.7242–2.2528 | 0.0238–0.0266 |
| 1000 | unclosed | empty | fresh-tokens | 1.2753–1.2771 | 0.2501–0.2543 | 1.5593–1.8097 | 0.4894–0.5166 |
| 1000 | unclosed | multi | cached-tokens | 1.0469–1.0485 | 0.0232–0.0234 | 1.1189–1.3402 | 0.0258–0.0318 |
| 1000 | unclosed | multi | fresh-tokens | 1.1605–1.1652 | 0.1393–0.1444 | 1.2655–1.3157 | 0.4150–0.4425 |
| 1000 | receiver | empty | cached-tokens | 1.6253–1.6283 | 0.0508–0.0516 | 1.8036–2.0241 | 0.0593–0.0615 |
| 1000 | receiver | empty | fresh-tokens | 1.8122–1.8215 | 0.2317–0.2379 | 2.1883–2.1891 | 0.6445–0.6452 |
| 1000 | receiver | multi | cached-tokens | 1.6300–1.6448 | 0.0495–0.0497 | 1.7389–1.7714 | 0.0535–0.0684 |
| 1000 | receiver | multi | fresh-tokens | 1.8167–1.8340 | 0.2318–0.2424 | 2.3251–2.5396 | 0.6287–0.6391 |
| 1000 | multi | empty | cached-tokens | 1.6219–1.6285 | 0.0452–0.0453 | 1.6622–1.8118 | 0.0486–0.0513 |
| 1000 | multi | empty | fresh-tokens | 1.8083–1.8843 | 0.2382–0.4439 | 2.1842–2.2382 | 1.0165–1.3876 |
| 1000 | multi | multi | cached-tokens | 0.0170–0.0171 | 0.0193–0.0328 | 0.9440–0.9727 | 0.0225–0.0473 |
| 1000 | multi | multi | fresh-tokens | 0.2070–0.2355 | 0.3888–0.5943 | 0.3842–0.4010 | 0.6744–0.9884 |
