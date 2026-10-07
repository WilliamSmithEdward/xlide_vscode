# Object assignment query performance

Baseline: `fb60195bc9d6810e093f1ee8ee1a02d5795f85b6` (PR #1178).

Variable and held-object Set compatibility already shared type resolution but still scanned implemented-name lists per statement. Property Set compatibility additionally resolved the same project types from scratch. Share one lazy implemented-name membership lookup in the variable rule and one type/interface/membership query across the complete property-assignment rule, including separate procedures. The same compatibility helper and alias semantics remain in use. Statement-specific held-object inference stays per statement, and each public query creates fresh metadata lookups. No dependency is added.

## Validation

Types and 252 focused tests pass. Eighteen new tests cover variable/held/property paths: nine bounded-work cases at 10/100/1,000 statements, three cross-procedure work bounds, three exact mismatch message/span controls and three next-query interface mutation controls. The baseline fails all 12 work bounds and passes the six semantic controls.

Full suite: 15,460 tests passed, 33 skipped across 789 passing files and seven skipped files, in 80.58 seconds. Complete comparisons match 20,000 generated public Set-rule results, 2,000 generated complete-module results and 24,768 corpus diagnostic/error results over 8,256 sources with LF/CRLF/CR; zero differences. Generated inputs vary classes, duplicate eligibility, completeness, implemented-name aliases, host/project/generic/scalar/unknown types, New/Nothing/scalar/ActiveSheet/held values and property accessor metadata. Full-module cases exercise property assignment through its actual caller; Set-rule comparisons exercise its visitor interface. No diagnostic changes are intended.

## Work

Each fixture has N assignments and one N-entry list whose last entry implements Class1. Compatible variable, held-object and writable property assignments independently produce no diagnostics and no internal errors. Untimed getter/proxy counters include the whole module query. Timed metadata uses only plain objects and arrays.

| N = 1,000 path | Before class-name reads | After class-name reads | Before interface-string reads | After interface-string reads |
| --- | ---: | ---: | ---: | ---: |
| variable | 10,035 | 10,035 | 1,000,000 | 1,000 |
| held | 8,023 | 8,023 | 1,000,000 | 1,000 |
| property | 2,021,052 | 14,057 | 1,000,000 | 1,000 |

## Timings

Node v24.18.0; AMD Ryzen 7 9800X3D 8-Core Processor           . Sequential baseline/current/current/baseline runs, three warmups and nine measured rounds per fixture. These are full-module warm diagnostic calls including analyzer setup and internal-error capture. Complete-output equality runs outside the clock. Table ranges are the two trial medians in milliseconds. No editor, cold-start or heap claim.

| Path | N | Before median ms | After median ms |
| --- | ---: | ---: | ---: |
| variable | 1 | 1.1209–1.1901 | 0.9561–0.9587 |
| variable | 100 | 3.9491–4.0290 | 3.4173–3.4800 |
| variable | 1000 | 85.7122–87.1328 | 20.2703–20.5983 |
| held | 1 | 0.6015–0.6233 | 0.5939–0.6063 |
| held | 100 | 3.1847–3.1920 | 2.4739–2.4800 |
| held | 1000 | 84.1677–87.6913 | 17.9990–18.1438 |
| property | 1 | 0.6225–0.6808 | 0.6106–0.6144 |
| property | 100 | 4.1714–4.3589 | 3.0633–3.0878 |
| property | 1000 | 121.8119–123.1793 | 21.1938–21.2277 |

One held-object assignment has overlapping trial ranges; do not infer a consistent improvement there. The fixtures use a late interface match. First membership lookup indexes the whole consulted list, whereas an old isolated early match could stop sooner. The experiment does not establish a benefit for that workload. Remaining linear class-name work includes the complete analyzer's other lookups.

Reproduce from the repository root:

```powershell
node scripts/benchmark-assignment-object-queries.mjs --baseline=fb60195bc9d6810e093f1ee8ee1a02d5795f85b6
node scripts/benchmark-assignment-object-queries.mjs
```
