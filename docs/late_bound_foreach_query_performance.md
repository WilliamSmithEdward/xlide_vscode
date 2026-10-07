# For Each compatibility queries

The runtime-member rule resolved project types separately for every held Collection item and every host collection loop. Shared-interface compatibility repeatedly traversed project Implements metadata. Even after caching those properties, direct and reverse casts still scanned long implemented-name lists per item, as did the generic Collection compatibility path.

Share query-owned type, common-implementer and direct-membership lookups across the complete public query. The direct-membership lookup indexes only consulted lists, keyed by their readonly array identity, and refreshes with each public query. Reuse one expected-alias helper for both indexed and original matching. The new optional callback covers direct, reverse and generic Collection checks; existing callers retain the original scan and Collection fast path. Class-member projection stays shared with its other consumers.

Host-item inference still uses its existing explicit Excel fallback. Compatibility resolution already defaults an omitted model to Excel, so both paths share one resolver while honoring supplied models. Preserve eligible duplicate ambiguity, excluded value kinds, all supplied shared-implementer kinds, source spelling/spans, first incompatible item, scalar refusal, early loop exit, conditional activity and fresh metadata.

## Reproduction

Run `node scripts/benchmark-late-bound-foreach-queries.mjs`, then with `--baseline=00e537640a668a5f3717437592f151b219d60014`. The pinned reachable baseline overrides only lateBoundMembers.ts and typeInference.ts; other dependencies are common.

Each fixture has N+2 project classes. Held fixtures fill a Collection with N items and iterate into a Class1 control; host fixtures have N separate procedures iterating Worksheets into Class1. Shared fixtures have a final bridge implementing both Class1/Class2, or every Class1 through ClassN+1 for distinct items. Long direct fixtures give Class2 N implemented names with Class1 last; reverse fixtures give Class1 the N distinct item types. The generic fixture iterates Class2 items into Collection, with Collection last in Class2's N-name list.

Direct and full queries independently require N exact host mismatch diagnostics (messages and original spans in the direct scope), or empty output for the held/interface fixtures. Complete queries also require no internal failures. Work counters count metadata properties and real string values behind array-index proxies, exclusively outside timing; all timed names and lists are plain properties and arrays.

Sequential before/after/after/before on Node 24.18.0 and Ryzen 7 9800X3D. Three warmups, nine measured calls, ranges of the two run medians in milliseconds. Parsed AST and symbols are outside direct-rule timing; complete calls include analyzer preparation and error checks. Complete output equality is outside the clocks.

| Fixture | Items / loops | Scope | Before ms | After ms |
| --- | ---: | --- | ---: | ---: |
| held-same | 1 | rule | 0.0938–0.0967 | 0.0973–0.1041 |
| held-same | 1 | complete-module-diagnostics | 1.4789–1.4846 | 1.6002–1.7548 |
| held-same | 100 | rule | 0.6306–0.6517 | 0.2156–0.2185 |
| held-same | 100 | complete-module-diagnostics | 5.4217–5.4733 | 4.9692–5.0071 |
| held-same | 1000 | rule | 33.8245–34.3799 | 0.9576–1.0358 |
| held-same | 1000 | complete-module-diagnostics | 64.2388–65 | 30.2685–30.4506 |
| host-loops | 1 | rule | 0.0534–0.0536 | 0.0516–0.0532 |
| host-loops | 1 | complete-module-diagnostics | 0.7057–0.7076 | 0.6667–0.8703 |
| host-loops | 100 | rule | 0.7332–0.7341 | 0.468–0.4749 |
| host-loops | 100 | complete-module-diagnostics | 10.9616–11.0452 | 11.2935–11.4951 |
| host-loops | 1000 | rule | 20.5148–21.5096 | 3.9338–3.975 |
| host-loops | 1000 | complete-module-diagnostics | 153.6323–156.9724 | 141.4146–142.2584 |
| shared-repeated | 1 | rule | 0.0379–0.0381 | 0.0387–0.0395 |
| shared-repeated | 1 | complete-module-diagnostics | 0.5127–0.5271 | 0.5319–0.5323 |
| shared-repeated | 100 | rule | 0.9024–0.9073 | 0.2037–0.2122 |
| shared-repeated | 100 | complete-module-diagnostics | 3.8777–3.8779 | 3.1219–3.1989 |
| shared-repeated | 1000 | rule | 66.1038–67.7173 | 1.2175–1.2556 |
| shared-repeated | 1000 | complete-module-diagnostics | 94.1551–97.6862 | 29.0669–29.8308 |
| shared-distinct | 1 | rule | 0.0372–0.0378 | 0.0383–0.066 |
| shared-distinct | 1 | complete-module-diagnostics | 0.5625–0.576 | 0.5596–0.6601 |
| shared-distinct | 100 | rule | 1.2401–1.2578 | 0.2406–0.2783 |
| shared-distinct | 100 | complete-module-diagnostics | 4.3815–4.486 | 3.4266–3.6276 |
| shared-distinct | 1000 | rule | 105.8448–108.2768 | 1.9534–2.032 |
| shared-distinct | 1000 | complete-module-diagnostics | 133.2497–135.7161 | 28.6258–29.1592 |
| direct-repeated | 1 | rule | 0.0368–0.0375 | 0.0372–0.0386 |
| direct-repeated | 1 | complete-module-diagnostics | 0.5328–0.5459 | 0.5219–0.5383 |
| direct-repeated | 100 | rule | 1.1225–1.1512 | 0.1388–0.1417 |
| direct-repeated | 100 | complete-module-diagnostics | 4.0635–4.0646 | 3.118–3.1497 |
| direct-repeated | 1000 | rule | 99.4488–99.753 | 1.0809–1.1066 |
| direct-repeated | 1000 | complete-module-diagnostics | 125.2702–126.7991 | 26.5445–26.5771 |
| reverse-distinct | 1 | rule | 0.0368–0.0369 | 0.0367–0.0373 |
| reverse-distinct | 1 | complete-module-diagnostics | 0.5403–0.5568 | 0.5124–0.5229 |
| reverse-distinct | 100 | rule | 0.8634–0.8771 | 0.2044–0.2065 |
| reverse-distinct | 100 | complete-module-diagnostics | 3.8498–3.8695 | 3.1437–3.1975 |
| reverse-distinct | 1000 | rule | 67.8989–70.2504 | 1.7784–1.8427 |
| reverse-distinct | 1000 | complete-module-diagnostics | 96.1547–99.0874 | 28.1895–28.2413 |
| generic-collection | 1 | rule | 0.0377–0.0394 | 0.0379–0.0383 |
| generic-collection | 1 | complete-module-diagnostics | 0.5597–0.5619 | 0.525–0.5474 |
| generic-collection | 100 | rule | 0.3759–0.3866 | 0.1469–0.1471 |
| generic-collection | 100 | complete-module-diagnostics | 3.2492–3.4134 | 3.0125–3.085 |
| generic-collection | 1000 | rule | 24.7764–25.2816 | 1.0345–1.0746 |
| generic-collection | 1000 | complete-module-diagnostics | 51.6071–53.0022 | 26.9504–27.5129 |

At N=1,000, public-query counts are:

| Fixture | Project-name reads | Implements property reads | Implemented-string reads |
| --- | ---: | ---: | ---: |
| held-same | 2006000 → 1003 | 2000 → 1 | 0 → 0 |
| host-loops | 1003000 → 1003 | 1000 → 1 | 0 → 0 |
| shared-repeated | 2006000 → 1004 | 1004000 → 1004 | 2000 → 2 |
| shared-distinct | 2006000 → 2003 | 1004000 → 2003 | 1001000 → 1001 |
| direct-repeated | 2006000 → 1004 | 2000 → 2 | 1000000 → 1000 |
| reverse-distinct | 2006000 → 2003 | 2000 → 1001 | 500500 → 1000 |
| generic-collection | 1003000 → 1003 | 1000 → 1 | 1000000 → 1000 |

Small results are mixed. One-item held-same rule and complete calls are slower; one-item shared-repeated rule and complete calls are slower, and the shared-distinct direct call is slower while its complete ranges overlap. One-loop host complete ranges overlap, and 100-loop host complete calls are slower. One-item direct-repeated complete ranges overlap. A first membership query now indexes a complete consulted list rather than stopping at an early match; unused lists remain untouched. No editor, cold-start or heap improvement is claimed.

## Validation

- Types and 81 focused tests pass.
- Forty-seven new tests: 22 work-bound tests fail against both baseline production files; 17 pre-existing-semantics controls pass there. Eight controls for the new lookup pin qualified/unqualified aliases, case, duplicate membership, sparse lists and fresh queries. Public controls cover repeated/distinct shared interfaces, direct/reverse/generic compatibility, all shared-implementer kinds, duplicate ambiguity, first failure, early exit, scalar items, host selection, activity and metadata replacement.
- Full suite: 15,442 tests across 788 files pass; 33 tests and seven files skipped. Final types/focused checks pass after completing the expected project-type fixture shape; runtime code is unchanged from the full run.
- 20,000 generated complete public-rule results, 2,000 generated complete-module results and 24,768 oracle module diagnostic/error results match the pinned baseline, with no differences. Generated cases vary project kind/duplicates/completeness/Implements facts, control/item/collection types, loop exits and branches, plus omitted, explicit Excel, Word and custom host models. Oracle comparison covers 8,256 sources across LF/CRLF/CR.
