# Call-site physical bounds

Issue #959 records Introduce Parameter call-site corruption and repeated whole-module scans in CR-only source. LF-only declaration and argument bounds either skipped real calls as declarations or included following statements in a bare call replacement. Bracket matching could also reach a later physical line, including inside an unterminated string.

The helper now uses shared CR/CRLF/LF physical-line bounds and rejects a physical break while looking for a closing bracket. Its duplicate end-of-line function is removed. This preserves existing output on compatible sources while making CR-only edits match canonical LF.

## Work and timing

For 1,000 bare calls, measured text passed to the helper's imported stripVba falls from 5,034,000 characters to 1,000 on CR source. The LF and CRLF controls remain at 1,000. At 100 CR calls it falls from 53,400 to 100. This counter excludes the occurrence scanner's own source stripping and other work; it does not claim zero total scanning or allocation.

Run node scripts/benchmark-call-site-bounds.mjs --rounds=9, optionally with --baseline=306b8c74. Four sequential runs use ABBA order, three warmups and nine measured samples each. Cached mode reuses identical source; fresh mode changes a leading comment to miss the stripped-source cache. Timing covers the actual callSitesOf API; result serialization/hashing happens outside the measured interval. Node v24.18.0, AMD Ryzen 7 9800X3D 8-Core Processor           . These are local microbenchmarks, not UI or Office timing.

The 36 configurations cover 1/100/1,000 calls, three endings, bare/bracketed syntax and both cache modes. Thirty result hashes match; six bare CR hashes intentionally change because their replacement span previously included later code. For 1,000 bracketed CR calls with unchanged hashes, cached median drops from 45.14–45.19 ms to 0.22–0.23 ms. LF/CRLF and small cases are mixed; cached 1,000-call CRLF bracketed runs increase from 0.33–0.41 to 0.41–0.43 ms. All measured rows are retained below.

| Calls / ending / form / cache | Before median ms | After median ms | Result hash |
| --- | ---: | ---: | --- |
| 1/lf/bare/cached | 0.00600–0.00600 | 0.00600–0.00610 | same |
| 1/lf/bare/fresh | 0.00570–0.00580 | 0.00610–0.00630 | same |
| 1/lf/bracketed/cached | 0.00180–0.00190 | 0.00180–0.00190 | same |
| 1/lf/bracketed/fresh | 0.00470–0.00490 | 0.00470–0.00480 | same |
| 1/crlf/bare/cached | 0.00150–0.00150 | 0.00160–0.00160 | same |
| 1/crlf/bare/fresh | 0.00360–0.00370 | 0.00420–0.00440 | same |
| 1/crlf/bracketed/cached | 0.00230–0.00250 | 0.00180–0.00180 | same |
| 1/crlf/bracketed/fresh | 0.00430–0.00450 | 0.00540–0.00710 | same |
| 1/cr/bare/cached | 0.00210–0.00210 | 0.00150–0.00170 | corrected CR edit |
| 1/cr/bare/fresh | 0.00550–0.00580 | 0.00430–0.00430 | corrected CR edit |
| 1/cr/bracketed/cached | 0.00200–0.00200 | 0.00180–0.00200 | same |
| 1/cr/bracketed/fresh | 0.00390–0.00400 | 0.00540–0.00540 | same |
| 100/lf/bare/cached | 0.05310–0.05450 | 0.04000–0.04410 | same |
| 100/lf/bare/fresh | 0.03410–0.03440 | 0.03500–0.03730 | same |
| 100/lf/bracketed/cached | 0.02890–0.02980 | 0.03120–0.03730 | same |
| 100/lf/bracketed/fresh | 0.04750–0.04800 | 0.05130–0.05150 | same |
| 100/crlf/bare/cached | 0.02050–0.02160 | 0.02060–0.02120 | same |
| 100/crlf/bare/fresh | 0.03230–0.03350 | 0.03240–0.03240 | same |
| 100/crlf/bracketed/cached | 0.02850–0.02860 | 0.02930–0.03250 | same |
| 100/crlf/bracketed/fresh | 0.04640–0.04680 | 0.04770–0.04980 | same |
| 100/cr/bare/cached | 0.50760–0.51520 | 0.01720–0.01890 | corrected CR edit |
| 100/cr/bare/fresh | 0.38820–0.49340 | 0.02770–0.03060 | corrected CR edit |
| 100/cr/bracketed/cached | 0.48790–0.49700 | 0.02810–0.02900 | same |
| 100/cr/bracketed/fresh | 0.32940–0.33860 | 0.04610–0.04800 | same |
| 1000/lf/bare/cached | 0.18790–0.23930 | 0.18600–0.25930 | same |
| 1000/lf/bare/fresh | 0.46550–0.47100 | 0.27450–0.29820 | same |
| 1000/lf/bracketed/cached | 0.22990–0.24770 | 0.22500–0.24720 | same |
| 1000/lf/bracketed/fresh | 0.39070–0.40830 | 0.36760–0.39660 | same |
| 1000/crlf/bare/cached | 0.16220–0.17250 | 0.12780–0.14470 | same |
| 1000/crlf/bare/fresh | 0.25510–0.27590 | 0.23240–0.24560 | same |
| 1000/crlf/bracketed/cached | 0.32970–0.40800 | 0.40880–0.43390 | same |
| 1000/crlf/bracketed/fresh | 0.66100–0.68290 | 0.56590–0.65500 | same |
| 1000/cr/bare/cached | 42.01870–48.35470 | 0.14660–0.14840 | corrected CR edit |
| 1000/cr/bare/fresh | 33.22780–39.72830 | 0.23380–0.23650 | corrected CR edit |
| 1000/cr/bracketed/cached | 45.13680–45.19430 | 0.21950–0.23320 | same |
| 1000/cr/bracketed/fresh | 26.66370–26.69870 | 0.36630–0.39190 | same |

## Validation and remaining scope

Twenty-one regression tests include actual Introduce Parameter output, declaration skipping, qualified calls/comments, physical-break rejection and a work bound. Baseline fails eight and passes thirteen controls. The focused three-file suite passes 87 tests; the type check passes.

All 19,814 complete call-site queries over 8,256 corpus sources match baseline, including skip-span queries and three generated argument values. Thirty-six complete generated refactor results match LF/CRLF baseline, and eighteen CR-only results match canonical LF. Primary and external-module rendered source is compared exactly. The full suite passes on base 306b8c74: 673 files, 13,525 tests passed and 18 skipped.

Issue #960 separately records lost raw string arguments, colon-neighbor rewrites and bare parenthesized argument handling. Those remain unfixed in this PR; this boundary change does not claim completion of the call-site or analyzer audit.
