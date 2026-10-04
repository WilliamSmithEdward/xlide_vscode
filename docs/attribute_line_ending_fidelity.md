# Attribute writer line-ending fidelity

The writer stored one module-wide newline and serialized every original line with it. Mixed LF/CRLF source changed even when an accepted annotation already matched and the changes list was empty. Per-line endings now survive updates and serialization. A shared insertion helper uses the anchor's separator for new attributes; insertion at an unterminated EOF adds only the required separator and retains an unterminated final line. This also removes duplicated linked-list insertion logic.

Validation against b678a385: nine exact-text tests, five failing before the fix; 57 focused annotation tests on main 060bd166; type checking and the full suite (658 files, 13,305 passed tests, 13 skipped); 12,960 complete reader/writer comparisons (8,256 corpus and 4,704 generated LF/CRLF cases), with reader results frozen before writing. Mixed-ending output intentionally changes from baseline and has independent exact-text expectations. Pure CR writer input is tested with supplied metadata; this does not claim that the annotation reader recognizes CR-only source or that Office accepts arbitrary mixed-ending modules.

Run scripts/benchmark-attribute-line-endings.mjs with --baseline=b678a385 and --output=<path> for the old writer, then without --baseline for the fixed writer. Twenty-four configurations exercise the actual reader and writer together: 1/100/1,000 procedures, LF/CRLF, aligned attributes, changed attributes, inserted attributes and one already-aligned module annotation. Source preparation, assertions and hashing are outside the timer; each process has three warmups and fifteen measured rounds, using batches for small cases. Four separate processes ran in A/B/B/A order; every complete result hash matched. No Office or file IO is timed.

These measurements ran while test suites were active elsewhere on the shared machine (including initial overlap with this audit's suite). They are contention-affected observations, not an isolated benchmark or evidence of a speedup. The fix is for source fidelity; per-line separator storage adds work and all measured configurations are reported here. Node v24.18.0.

| Procedures | Ending | Shape | Baseline median ms range | Fixed median ms range | Baseline p95 ms range | Fixed p95 ms range |
| ---: | --- | --- | ---: | ---: | ---: | ---: |
| 1 | LF | aligned | 0.0050–0.0050 | 0.0046–0.0047 | 0.0072–0.0081 | 0.0066–0.0067 |
| 1 | LF | changed | 0.0045–0.0046 | 0.0043–0.0044 | 0.0096–0.0156 | 0.0079–0.0090 |
| 1 | LF | inserted | 0.0030–0.0030 | 0.0030–0.0031 | 0.0043–0.0056 | 0.0042–0.0072 |
| 1 | LF | module-only | 0.0033–0.0042 | 0.0033–0.0038 | 0.0332–0.1014 | 0.0104–0.2077 |
| 1 | CRLF | aligned | 0.0042–0.0043 | 0.0039–0.0039 | 0.0048–0.0122 | 0.0049–0.0055 |
| 1 | CRLF | changed | 0.0044–0.0047 | 0.0038–0.0039 | 0.0062–0.1811 | 0.1725–0.2419 |
| 1 | CRLF | inserted | 0.0033–0.0033 | 0.0030–0.0032 | 0.0034–0.0036 | 0.0036–0.0043 |
| 1 | CRLF | module-only | 0.0032–0.0034 | 0.0029–0.0030 | 0.0035–0.0035 | 0.0032–0.0051 |
| 100 | LF | aligned | 0.3331–0.3731 | 0.3138–0.3222 | 0.4394–1.7317 | 0.3989–2.9163 |
| 100 | LF | changed | 0.3229–0.3446 | 0.3657–0.3750 | 0.7330–2.9457 | 1.4719–2.9346 |
| 100 | LF | inserted | 0.2205–0.2211 | 0.2327–0.2382 | 0.3550–1.9377 | 1.6465–9.9184 |
| 100 | LF | module-only | 0.0482–0.0517 | 0.0547–0.0571 | 0.0861–0.9183 | 0.0857–0.0859 |
| 100 | CRLF | aligned | 0.3279–0.4859 | 0.3131–0.3227 | 0.4835–1.7542 | 0.4351–0.7435 |
| 100 | CRLF | changed | 0.3744–0.4638 | 0.3294–0.3585 | 1.5378–2.5200 | 0.3788–2.3658 |
| 100 | CRLF | inserted | 0.2433–0.3295 | 0.2311–0.2349 | 0.4154–1.6496 | 0.2411–0.2551 |
| 100 | CRLF | module-only | 0.0740–0.0827 | 0.0575–0.0587 | 0.1684–0.1982 | 0.0631–0.1192 |
| 1000 | LF | aligned | 10.8538–11.1550 | 7.7971–9.2252 | 30.9118–40.0812 | 53.5593–55.7533 |
| 1000 | LF | changed | 11.0050–11.0646 | 8.7639–10.9588 | 90.0491–100.9943 | 63.3204–73.9842 |
| 1000 | LF | inserted | 8.6149–9.9509 | 9.2239–9.9389 | 46.5344–187.1548 | 39.3955–158.8154 |
| 1000 | LF | module-only | 0.4743–0.5128 | 0.5494–0.6578 | 0.6056–1.0800 | 2.0289–29.3099 |
| 1000 | CRLF | aligned | 9.3411–12.2822 | 10.6265–11.5373 | 43.7532–53.5220 | 56.2901–83.7217 |
| 1000 | CRLF | changed | 9.1446–9.6878 | 9.5792–10.9473 | 11.6281–94.2483 | 46.1138–82.2304 |
| 1000 | CRLF | inserted | 7.2934–8.4525 | 7.1378–7.9152 | 11.3774–44.4361 | 9.7504–42.5045 |
| 1000 | CRLF | module-only | 0.5758–0.6394 | 0.5520–0.5907 | 1.9650–26.4740 | 1.2699–3.0020 |
