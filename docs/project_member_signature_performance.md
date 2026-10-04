# Project member signature lookup performance

Baseline: `d2cc3a881da2228042bacf8942e2033eb7ad9a13` (PR #1197 source head). Baseline builds restore only completion/memberAccess.ts; surrounding files are identical.

A completion row already holds its selected surface member, but an absent signature triggers a second project-member name scan. A pure project surface selects the first case-insensitive project member before controls/designer base members; this fallback cannot add a signature. Keep its absent signature absent. Host, combined, union and runtime fallback paths retain their existing behavior. No new cache or lifetime is introduced.

## Validation

Types and 137 focused tests pass. Eighteen new cases include nine work bounds for first/last named lookups and complete menus at 10/100/1,000 members, five receiver-kind controls, three first-duplicate signature controls and retained-metadata signature freshness across contexts. The baseline fails six work bounds and passes three first-hit bounds and all nine semantic controls. Tests compare entire independently expected completion rows, including undefined fields and first-member precedence with later duplicate signatures.

Full suite: 15,740 passed, 33 skipped across 798 passing files and seven skipped files, in 93.45 seconds.

Complete outputs match 10,000 generated completion/menu pairs, 20,000 generated rule bundles, 2,000 generated complete modules and 24,768 corpus runs over 8,256 sources with LF/CRLF/CR; zero differences. Completion cases vary signature presence/emptiness, duplicate casing, hidden members, surface kinds, cached/uncached contexts, direct/Me/code-name/host-union/bracketed access and designer classes. Diagnostics compare complete findings and internal-error lists.

## Work

Class1 has N Object properties without signatures. Named distinct queries share a fresh per-pass surface cache; menu cases build one full menu with/without a cache. Full diagnostics include N distinct object-property writes. Work counts use name getters only outside timing. Expectations are complete independent completion rows or exact N Set-required code/message/spans with no internal errors.

| Scope | N | Before member-name reads | After member-name reads |
| --- | ---: | ---: | ---: |
| named-distinct-pass | 1 | 5 | 3 |
| menu-cached | 1 | 5 | 3 |
| menu-uncached | 1 | 5 | 3 |
| complete-module-diagnostics | 1 | 11 | 7 |
| named-distinct-pass | 100 | 5,450 | 300 |
| menu-cached | 100 | 5,450 | 300 |
| menu-uncached | 100 | 5,450 | 300 |
| complete-module-diagnostics | 100 | 15,950 | 5,650 |
| named-distinct-pass | 1000 | 504,500 | 3,000 |
| menu-cached | 1000 | 504,500 | 3,000 |
| menu-uncached | 1000 | 504,500 | 3,000 |
| complete-module-diagnostics | 1000 | 1,509,500 | 506,500 |

This removes the redundant signature scan. Other project-member lookups and uncached surface construction remain separate costs; no universal linear-work claim.

## Timings

Node v24.18.0; AMD Ryzen 7 9800X3D 8-Core Processor. Sequential baseline/current/current/baseline trials; three warmups and nine measured rounds. Ranges are the two trial medians in milliseconds. Parsing is excluded from completion timing and included in complete diagnostics. Getter properties are replaced with plain names before timing. Output assertions run outside the clock. No editor/heap/cold-start claim.

| Scope | N | Before median ms | After median ms |
| --- | ---: | ---: | ---: |
| named-distinct-pass | 1 | 0.0126–0.0133 | 0.0111–0.0130 |
| menu-cached | 1 | 0.0046–0.0059 | 0.0039–0.0039 |
| menu-uncached | 1 | 0.0032–0.0032 | 0.0037–0.0040 |
| complete-module-diagnostics | 1 | 1.0925–1.1546 | 1.1930–1.2427 |
| named-distinct-pass | 100 | 0.3143–0.3672 | 0.2534–0.2570 |
| menu-cached | 100 | 0.1477–0.1488 | 0.0929–0.0958 |
| menu-uncached | 100 | 0.1453–0.2857 | 0.0773–0.0792 |
| complete-module-diagnostics | 100 | 4.3471–4.5337 | 4.2858–4.5009 |
| named-distinct-pass | 1000 | 6.0356–6.2385 | 0.9295–0.9694 |
| menu-cached | 1000 | 5.5063–5.5344 | 0.1692–0.1711 |
| menu-uncached | 1000 | 5.4357–5.5446 | 0.1605–0.1609 |
| complete-module-diagnostics | 1000 | 43.7410–45.4578 | 31.3576–32.4101 |

All 1,000-member cases improve in both trials. Single-member uncached menus and single-write complete diagnostics are slower in both trials; named single-member ranges overlap, as do 100-write complete-module ranges. These measurements do not establish an improvement for every small module. Complete diagnostics still read 506,500 member names at 1,000 distinct writes, including a separate raw project-member lookup; this change removes only the redundant signature scans.

Reproduce sequentially from the repository root:

```powershell
node scripts/benchmark-project-member-signatures.mjs --baseline=d2cc3a881da2228042bacf8942e2033eb7ad9a13
node scripts/benchmark-project-member-signatures.mjs
node scripts/benchmark-project-member-signatures.mjs
node scripts/benchmark-project-member-signatures.mjs --baseline=d2cc3a881da2228042bacf8942e2033eb7ad9a13
```
