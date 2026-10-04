# Annotation physical-line parsing

The reader split only on LF and stripped one trailing CR from every line. CR-only bodies therefore returned no accepted annotations, and mixed CR/LF/CRLF sources lost bindings and miscounted placement lines. The save-path joinVbaSource helper adds a CRLF after the header while preserving body text, so an annotated CR-only body reaches this reader as mixed source. Recognize physical lines using CRLF, CR or LF, then read them directly without the redundant per-line replacement. The writer from #924 preserves each original line ending.

Nine exact metadata/output tests include five baseline failures and four LF/CRLF controls. They cover module and variable annotations, property Get/Let occurrence targeting, duplicate/placement/dangling problems at physical line numbers, exact rewrite output, idempotence, and actual save-path assembly. Sixty-six focused tests and type checking passed. Final full validation on main 946c64e4 passed 661 files and 13,349 tests (14 skipped). This is pure source assembly and reader/writer validation, not Office execution or a claim that all CR export/header handling elsewhere is correct.

Against baseline c3902b21, 8,256 corpus sources and 4,704 generated LF/CRLF cases retain complete reader and writer results. Another 2,352 generated CR cases are compared against the baseline reader/writer on the corresponding LF source, with only output line endings canonicalized: complete annotations, problems, changes, skipped messages and text must match. Reader result objects and arrays are frozen before writing. All 15,312 comparisons pass without uncaught exceptions. CR behavior is intentionally repaired, not baseline-compatible.

A deterministic observation through the actual reader counts only the removed trailing-CR String.replace calls: a module with one description and 1,000 unannotated procedures had 2,003 calls before and zero after. This excludes line splitting, declaration recognition, argument decoding and all other reader work; it does not claim that the reader now performs zero work.

Performance controls reuse scripts/benchmark-annotation-duplicates.mjs --baseline=c3902b21, then the same script without --baseline. Four separate processes ran A/B/B/A before this audit's full suite; all 42 complete output hashes match. Three warmups and fifteen rounds, counts 1/100/1,000 across annotated procedures, variables, repeated procedure/module annotations, one module annotation, unknown annotations and no markers, each for reader and guarded reader/writer pipeline. Source preparation, assertions and hashing are outside timing. No IO, Office or complete save operation is measured. The table below is LF. The benchmark also supports --eol=crlf; a second four-process A/B/B/A run covers those same 42 configurations. Its complete hashes also match; the LF default preserves all 42 prior fixed hashes. Both runs occurred after this audit’s benchmark inputs were prepared and outside its full-suite runs. All cases are reported, including already-cheap cases that can be mixed or slower. Node v24.18.0; shared Windows host, absolute timings are host-dependent.

| Count | Shape | Scope | Baseline median ms | Fixed median ms | Baseline p95 ms | Fixed p95 ms |
| ---: | --- | --- | ---: | ---: | ---: | ---: |
| 1 | procedures | reader | 0.0033–0.0045 | 0.0036–0.0037 | 0.0096–0.0097 | 0.0093–0.0094 |
| 1 | procedures | reader-writer | 0.0084–0.0085 | 0.0063–0.0083 | 0.0162–0.0163 | 0.0150–0.0164 |
| 1 | variables | reader | 0.0017–0.0018 | 0.0015–0.0015 | 0.0051–0.0051 | 0.0047–0.0051 |
| 1 | variables | reader-writer | 0.0037–0.0038 | 0.0030–0.0031 | 0.0058–0.0060 | 0.0044–0.0044 |
| 1 | procedure-duplicates | reader | 0.0020–0.0021 | 0.0018–0.0018 | 0.0042–0.0044 | 0.0031–0.0075 |
| 1 | procedure-duplicates | reader-writer | 0.0040–0.0049 | 0.0036–0.0037 | 0.0059–0.0063 | 0.0060–0.0073 |
| 1 | module-duplicates | reader | 0.0024–0.0024 | 0.0026–0.0027 | 0.0056–0.0082 | 0.0037–0.0046 |
| 1 | module-duplicates | reader-writer | 0.0058–0.0060 | 0.0050–0.0051 | 0.0122–0.0151 | 0.0100–0.0103 |
| 1 | module-only | reader | 0.0016–0.0017 | 0.0021–0.0021 | 0.0018–0.0040 | 0.0025–0.0039 |
| 1 | module-only | reader-writer | 0.0034–0.0036 | 0.0026–0.0029 | 0.0054–0.0064 | 0.0062–0.0071 |
| 1 | unknown-only | reader | 0.0013–0.0014 | 0.0014–0.0014 | 0.0034–0.0036 | 0.0020–0.0021 |
| 1 | unknown-only | reader-writer | 0.0014–0.0023 | 0.0012–0.0012 | 0.0018–0.0031 | 0.0035–0.0035 |
| 1 | no-annotations | reader | 0.0002–0.0003 | 0.0002–0.0002 | 0.0006–0.1876 | 0.0003–0.0005 |
| 1 | no-annotations | reader-writer | 0.0002–0.0002 | 0.0002–0.0002 | 0.0003–0.0026 | 0.0003–0.0003 |
| 100 | procedures | reader | 0.0853–0.0867 | 0.0763–0.0795 | 0.1889–0.2664 | 0.2442–0.2991 |
| 100 | procedures | reader-writer | 0.1503–0.1512 | 0.1429–0.1464 | 0.2719–0.2882 | 0.4147–0.4472 |
| 100 | variables | reader | 0.0703–0.0732 | 0.0639–0.0674 | 0.1869–0.2155 | 0.2176–0.2966 |
| 100 | variables | reader-writer | 0.1033–0.1045 | 0.0983–0.1077 | 0.2377–0.2480 | 0.1914–0.2014 |
| 100 | procedure-duplicates | reader | 0.0269–0.0274 | 0.0252–0.0265 | 0.0905–0.0940 | 0.0395–0.1223 |
| 100 | procedure-duplicates | reader-writer | 0.0278–0.0286 | 0.0260–0.0274 | 0.0907–0.0959 | 0.1117–0.1144 |
| 100 | module-duplicates | reader | 0.0959–0.1080 | 0.0909–0.1126 | 0.2320–0.2345 | 0.1948–0.2398 |
| 100 | module-duplicates | reader-writer | 0.1271–0.1377 | 0.1353–0.1355 | 0.2628–0.3112 | 0.2544–0.3084 |
| 100 | module-only | reader | 0.0249–0.0255 | 0.0174–0.0190 | 0.0976–0.2047 | 0.0261–0.0329 |
| 100 | module-only | reader-writer | 0.0273–0.0407 | 0.0336–0.0381 | 0.0467–0.1202 | 0.0477–0.0502 |
| 100 | unknown-only | reader | 0.0184–0.0287 | 0.0252–0.0272 | 0.0370–0.0696 | 0.0390–0.0402 |
| 100 | unknown-only | reader-writer | 0.0184–0.0374 | 0.0170–0.0171 | 0.0277–0.1825 | 0.0868–0.1026 |
| 100 | no-annotations | reader | 0.0004–0.0005 | 0.0005–0.0005 | 0.0005–0.0005 | 0.0005–0.0024 |
| 100 | no-annotations | reader-writer | 0.0005–0.0005 | 0.0005–0.0005 | 0.0005–0.0006 | 0.0005–0.0016 |
| 1000 | procedures | reader | 0.5855–0.6014 | 0.4878–0.5000 | 1.0176–1.4233 | 0.9416–1.2163 |
| 1000 | procedures | reader-writer | 1.0898–1.1482 | 1.0083–1.0091 | 1.6151–1.7110 | 1.1981–1.8175 |
| 1000 | variables | reader | 0.4755–0.4778 | 0.4067–0.4285 | 0.6864–0.6914 | 0.5885–0.5976 |
| 1000 | variables | reader-writer | 1.0007–1.0238 | 0.8636–1.0097 | 1.3576–1.4423 | 1.3518–1.6013 |
| 1000 | procedure-duplicates | reader | 0.1636–0.1742 | 0.1572–0.1639 | 0.3164–0.5463 | 0.3107–0.4135 |
| 1000 | procedure-duplicates | reader-writer | 0.2123–0.2253 | 0.2189–0.3653 | 0.3513–0.3883 | 0.4677–0.5682 |
| 1000 | module-duplicates | reader | 0.6345–0.7063 | 0.5370–0.5748 | 0.9190–0.9721 | 0.7658–0.7979 |
| 1000 | module-duplicates | reader-writer | 1.0925–1.2174 | 0.9994–1.0304 | 1.4791–1.7101 | 1.3430–1.7777 |
| 1000 | module-only | reader | 0.1727–0.1874 | 0.1460–0.1480 | 0.2649–0.2649 | 0.2208–0.2246 |
| 1000 | module-only | reader-writer | 0.2434–0.2647 | 0.2145–0.2162 | 0.3091–0.3455 | 0.2909–0.3229 |
| 1000 | unknown-only | reader | 0.1720–0.1947 | 0.1493–0.1495 | 0.1985–0.2334 | 0.1562–0.1614 |
| 1000 | unknown-only | reader-writer | 0.1723–0.1946 | 0.1473–0.1476 | 0.2657–0.2721 | 0.1501–0.1528 |
| 1000 | no-annotations | reader | 0.0038–0.0038 | 0.0037–0.0038 | 0.0038–0.0038 | 0.0038–0.0058 |
| 1000 | no-annotations | reader-writer | 0.0038–0.0038 | 0.0037–0.0038 | 0.0038–0.0038 | 0.0038–0.0038 |

## CRLF controls

| Count | Shape | Scope | Baseline median ms | Fixed median ms | Baseline p95 ms | Fixed p95 ms |
| ---: | --- | --- | ---: | ---: | ---: | ---: |
| 1 | procedures | reader | 0.0048–0.0071 | 0.0048–0.0067 | 0.0157–0.0169 | 0.0119–0.0150 |
| 1 | procedures | reader-writer | 0.0121–0.0134 | 0.0113–0.0119 | 0.0256–0.0282 | 0.0252–0.0307 |
| 1 | variables | reader | 0.0030–0.0035 | 0.0029–0.0031 | 0.0080–0.0084 | 0.0083–0.0097 |
| 1 | variables | reader-writer | 0.0066–0.0070 | 0.0054–0.0056 | 0.0104–0.0136 | 0.0083–0.0128 |
| 1 | procedure-duplicates | reader | 0.0036–0.0038 | 0.0031–0.0039 | 0.0052–0.0089 | 0.0042–0.0049 |
| 1 | procedure-duplicates | reader-writer | 0.0075–0.0087 | 0.0060–0.0081 | 0.0123–0.0131 | 0.0165–0.0184 |
| 1 | module-duplicates | reader | 0.0042–0.0045 | 0.0040–0.0045 | 0.0094–0.0122 | 0.0054–0.0082 |
| 1 | module-duplicates | reader-writer | 0.0103–0.0112 | 0.0093–0.0094 | 0.0191–0.0198 | 0.0176–0.0195 |
| 1 | module-only | reader | 0.0028–0.0028 | 0.0023–0.0034 | 0.0036–0.0071 | 0.0047–0.0104 |
| 1 | module-only | reader-writer | 0.0057–0.0065 | 0.0043–0.0053 | 0.0090–0.0129 | 0.0085–0.0133 |
| 1 | unknown-only | reader | 0.0026–0.0026 | 0.0020–0.0028 | 0.0114–0.0144 | 0.0120–0.0312 |
| 1 | unknown-only | reader-writer | 0.0027–0.0030 | 0.0022–0.0023 | 0.3595–0.5606 | 0.0049–0.0168 |
| 1 | no-annotations | reader | 0.0003–0.0004 | 0.0003–0.0003 | 0.0007–0.0008 | 0.0006–0.0006 |
| 1 | no-annotations | reader-writer | 0.0003–0.0004 | 0.0003–0.0003 | 0.0005–0.0008 | 0.0018–0.0049 |
| 100 | procedures | reader | 0.1603–0.1706 | 0.1422–0.1456 | 0.5117–2.0520 | 0.4156–0.8527 |
| 100 | procedures | reader-writer | 0.2872–0.3279 | 0.2699–0.2801 | 0.5430–0.5576 | 0.4749–1.1023 |
| 100 | variables | reader | 0.1284–0.1321 | 0.1004–0.1258 | 0.2918–0.5425 | 0.2608–0.3332 |
| 100 | variables | reader-writer | 0.1989–0.2026 | 0.1725–0.1741 | 0.5346–1.2844 | 0.2739–0.4151 |
| 100 | procedure-duplicates | reader | 0.0590–0.0621 | 0.0537–0.0539 | 0.1966–0.3887 | 0.1749–0.1773 |
| 100 | procedure-duplicates | reader-writer | 0.0619–0.0659 | 0.0529–0.0539 | 0.1720–0.7660 | 0.1791–1.0454 |
| 100 | module-duplicates | reader | 0.1690–0.1900 | 0.1334–0.1515 | 0.3438–0.3526 | 0.8207–2.2798 |
| 100 | module-duplicates | reader-writer | 0.2983–0.3225 | 0.2442–0.2547 | 0.8919–2.1724 | 1.1180–1.5807 |
| 100 | module-only | reader | 0.0537–0.0626 | 0.0324–0.0461 | 0.0901–0.1968 | 0.0500–0.0588 |
| 100 | module-only | reader-writer | 0.0642–0.0701 | 0.0498–0.0663 | 0.1786–0.2261 | 0.0796–0.1680 |
| 100 | unknown-only | reader | 0.0406–0.0416 | 0.0319–0.0419 | 0.0552–0.0569 | 0.0453–0.7123 |
| 100 | unknown-only | reader-writer | 0.0407–0.0460 | 0.0301–0.0326 | 0.0568–0.2139 | 0.0374–0.0521 |
| 100 | no-annotations | reader | 0.0008–0.0009 | 0.0008–0.0008 | 0.0009–0.1304 | 0.0009–0.0010 |
| 100 | no-annotations | reader-writer | 0.0009–0.0009 | 0.0009–0.0009 | 0.0010–0.0011 | 0.0009–0.0014 |
| 1000 | procedures | reader | 1.0367–1.0769 | 0.9172–0.9367 | 2.7041–56.7316 | 1.5577–13.3068 |
| 1000 | procedures | reader-writer | 2.6083–2.6704 | 2.1585–2.3651 | 72.3580–77.4889 | 6.4705–38.8697 |
| 1000 | variables | reader | 0.8610–1.1901 | 0.7568–0.7996 | 2.2592–3.2302 | 1.0935–2.4915 |
| 1000 | variables | reader-writer | 2.0957–2.6302 | 2.0752–2.4531 | 7.6378–19.1931 | 3.2287–4.5990 |
| 1000 | procedure-duplicates | reader | 0.3595–0.3920 | 0.3138–0.3557 | 0.3763–0.8337 | 0.6878–0.9825 |
| 1000 | procedure-duplicates | reader-writer | 0.4699–0.4922 | 0.3993–0.4179 | 0.9272–1.1139 | 0.7201–33.8410 |
| 1000 | module-duplicates | reader | 1.2518–1.2682 | 1.0373–1.0466 | 2.0394–9.5657 | 1.4015–1.6718 |
| 1000 | module-duplicates | reader-writer | 2.4133–2.8413 | 2.2262–2.4485 | 8.8804–32.7217 | 4.6271–12.2937 |
| 1000 | module-only | reader | 0.4093–0.4164 | 0.2837–0.3635 | 1.5154–1.8621 | 0.3499–2.7060 |
| 1000 | module-only | reader-writer | 0.5455–0.8101 | 0.4161–0.5504 | 1.0895–1.3765 | 0.6971–31.1779 |
| 1000 | unknown-only | reader | 0.4050–0.5322 | 0.2737–0.3121 | 1.3731–6.8374 | 0.4951–0.5114 |
| 1000 | unknown-only | reader-writer | 0.3800–0.4101 | 0.2914–0.3099 | 0.4153–0.4574 | 0.4927–0.7227 |
| 1000 | no-annotations | reader | 0.0061–0.0070 | 0.0070–0.0072 | 0.0077–0.0275 | 0.0077–0.0203 |
| 1000 | no-annotations | reader-writer | 0.0061–0.0071 | 0.0070–0.0071 | 0.0062–0.0079 | 0.0076–0.0224 |
