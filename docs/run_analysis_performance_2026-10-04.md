# Run Analysis performance hunt — 2026-10-04

Investigated `ROneCOne_Delegates_Demo.xlsm` from the local testing directory,
without writing to the original. Its eight modules include a 969,296-byte
`ROneCOne` code module. Measurements use Node v24.18.0 and the cached VS Code
1.139.1 integration host on this Windows machine.

## Finding and fix

`checkDeletedObjects` counted `source.slice(0, statementOffset).split('\n')`
for every statement in every eligible procedure, even when nothing was deleted
or closed. This scans and allocates a growing source prefix repeatedly, making
line-number work quadratic in module size. Most of this workbook's statements
need no deletion location at all.

The rule now creates a line-start index lazily, only on an actual delete/close,
and uses binary search to locate that statement. It shares the index throughout
one rule invocation. The existing source-scanner helper also handles lone CR
line endings correctly. No new dependency or diagnostic suppression is involved.

An alternating comparison of separately bundled original/fixed versions against
the same workbook, after one warm-up each, gave these seven-round medians:

| Deleted-object check | Original | Fixed |
|---|---:|---:|
| ROneCOne | 437.60 ms | 2.49 ms |

Both isolated checks produced zero findings; the complete worker pass's 553
findings were also identical to the saved pre-fix result. Synthetic regressions
exercise actual deletion findings after 1,000 padding lines with LF, CRLF and CR
endings, checking both the message's line number and the finding's source span.

## Run Analysis surface

The opt-in integration probe copies the workbook into the disposable workspace
and invokes the real `xlide.analyzeProject` command twice. With the fix, the
initial observed run took 5,819 ms and the unchanged repeat took 19 ms. The
extension host's largest timer gap was 555 ms. The trace reported:

- Initial module read: 552 ms; cached read: 6 ms.
- ROneCOne worker request: 4,732 ms.
- Initial results panel render: 16 ms; model update: 5 ms.
- Cached project analysis: 18 ms.

A final verification run took 4,631 ms initially and 13 ms cached, with a
529 ms maximum host timer gap and an 11 ms panel render. Type checking,
90 focused tests across nine files, and the integration probe passed.

The other modules' request durations include waiting behind ROneCOne in the
single worker queue; they are not independent CPU costs. The original rule was
also run through the integration probe and passed, but that session's startup,
settings and loading were all slower. Do not infer an aggregate percentage
improvement from the two cold command measurements. Local timing varies under
other work on the machine.

Rendering and quick-fix preparation did not account for the reported wait.
The remaining cost is predominantly semantic analysis of the large module;
rule instrumentation identified overflow analysis at approximately 1.1 seconds
in one pass, including its shared dataflow setup. The worker kept the host
responsive, and the completed-result cache made unchanged repeats inexpensive.
Startup CPU profiling heavily distorted wall time (one profiled pass took
66 seconds versus roughly 5 seconds unprofiled), so profiler wall times are
excluded from performance claims.

## Follow-up: pause at 5 of 8

The count describes completed modules. The five empty workbook/sheet modules
finish first, then the single worker processes ROneCOne's 1,552 procedures.
The progress message now also lists up to three remaining module names, so
the count identifies the pending work during that pause.

Two additional repeated symbol scans were removed:

- The code-running/dataflow checks now retain safe-name sets per immutable
  module/procedure symbol snapshot, rather than rebuilding the sets for each
  statement. Weak maps allow old snapshots to be collected. Tests cover
  ByRef/ByVal distinctions and changes between symbol snapshots.
- Module-member checks now index property, type and external-value candidate
  names once per pass. Ordinary assignments bypass name resolution when the
  required declaration cannot exist; candidate names still undergo the original
  binding and shadowing checks.

An alternating isolated comparison on the actual ROneCOne source (three warm-up
rounds, then seven measured rounds per version) reduced the module-member check
from a median 443.76 ms to 19.44 ms, with identical rule findings. The repeatable
scope benchmark with seven measured rounds gave these medians:

| Constants / statements | Original scope scans | Indexed scopes |
|---|---:|---:|
| 100 | 0.651 ms | 0.070 ms |
| 500 | 12.239 ms | 0.159 ms |
| 2,000 | 159.664 ms | 0.489 ms |

The full native worker benchmark's three-round median was 3,354.67 ms, including
3,260.80 ms in ROneCOne. All 553 findings match the original full-pass digest:
`d125e68c915418a4f6c8d3589ec60db01830f5c4ec458cd16f94ee544ec232e8`.
The final real-command integration run took 4,777 ms initially and 15 ms cached;
ROneCOne's request took 3,699 ms, the maximum host timer gap was 573 ms, and the
panel render took 10 ms. These observations still show several seconds of real
semantic work; they do not establish an aggregate percentage speedup across
separate extension-host sessions. Instrumentation continues to identify overflow
analysis as the largest remaining rule cost, roughly one second in one pass.
Type checking, compilation, 121 focused tests across 13 files, and the real
command integration probe passed.

Reproduce the scope benchmark against the pinned pre-change revision:

```powershell
node scripts/benchmark-code-running-scopes.mjs --rounds=7 --baseline=35ec3574bb7501bb47aea29f7336f8683285b4ab
node scripts/benchmark-code-running-scopes.mjs --rounds=7
```

## Further follow-up: statement value facts

A runtime CPU profile identified repeated calls to the statement-value reader
and the per-procedure scan that selects followable module variables. The reader
now retains each statement's facts in a weak map within its existing
source/symbol/activity context. Module-variable eligibility is indexed once per
immutable symbol snapshot; each procedure still removes names hidden by its
own locals, parameters and result name. This changes neither inference nor
which checks run.

An isolated alternating comparison against the saved pre-change implementation
queried each of ROneCOne's 9,369 statements ten times, with fresh symbols each
round. After one warm-up, five-round medians were 2,178.88 ms before and
1,040.80 ms after. This is a repeated-query workload, not the full command.
Regressions cover differing values at successive writes, local shadowing,
conditional activity changes and changing declaration types on a reused parse.

The full workbook's 553 findings retain the original digest. The native worker
benchmark's three-round median was 3,062.55 ms; a separate five-round saved
baseline measured 3,244.02 ms. Machine contention and separate runs preclude a
reliable aggregate percentage claim. The real-command integration probe took
4,952 ms initially and 16 ms cached, with a 650 ms maximum host timer gap.
Type checking, compilation, 169 focused tests across 16 files and the integration
probe passed. The remaining
first-run pause still contains several seconds of semantic work.

## Further follow-up: flow starts and lazy cache keys

Two more setup costs were removed. The module's literal condition constants are
now indexed once per immutable symbol snapshot and overlaid with each
procedure's declarations, retaining local/parameter shadowing. The shared flow
cache retains its first start by identity and only serializes/sorts value keys
when a different start requires comparison. Separately allocated equal starts
still share results; source and conditional activity remain cache guards.

Alternating isolated comparisons, each with three warm-up and seven measured
rounds, gave these medians:

| Workload | Saved implementation | Current implementation |
|---|---:|---:|
| Literal condition constants for all 1,552 ROneCOne procedures | 138.51 ms | 12.07 ms |
| 1,552 empty bodies with 261 constants each and five repeated identity queries | 138.08 ms | 1.51 ms |

The second row isolates key construction; it is not a full analysis benchmark.
Before the cache change, an instrumented full pass generated 1,713 start keys
covering 385,806 name/value entries, across 7,333 walk requests and 1,594 bodies.

A full-workbook comparison alternated the two versions, used fresh worker state
each pass, and measured five rounds after one warm-up. Medians were 2,914.23 ms
before and 2,734.56 ms after. All rounds and both versions matched the same
diagnostic digest. During this investigation, a concurrent change to
`rules/declarations.ts` added implicit VBA/host library types; it reduced this
workbook's findings from 553 to 16 independently of these performance changes.
The current shared baseline digest is
`4b2d985d6c6f5a131a0097ed4ab7a53daf277a36266b633409cb6dea0cab9f5b`.

Tests cover equal starts with different insertion orders, different values,
source and activity changes on reused bodies, and constant shadowing by locals,
parameters and local constants. All 260 focused tests across 21 files passed.
Cold integration timings remain noisy: an observed run took 7,835 ms initially
and 20 ms cached, with an 803 ms maximum host timer gap. The integration probe
passed; this measurement does not establish a cold-command speedup.
After rebuilding against the current workspace, the final integration probe
took 4,695 ms initially and 15 ms cached, including 3,344 ms in the ROneCOne
worker request, a 769 ms maximum host timer gap, and a 7 ms panel render.
Compilation (including type checking) and the probe passed.

## Further hunt: variable scopes, Nothing scans and default-object setup

Four additional setup scans were reduced:

- Field variable lookup now shares the module table and retains only each
  procedure's local overlay. Explicit undefined entries still hide module
  variables when a local declaration is a constant rather than a variable.
- Object-state setup indexes module declaration candidates and skips its
  per-variable Nothing regexes unless the procedure contains both required
  pieces of assignment text. Exact candidate checks still run when possible.
- String constant scopes reuse a module index, then apply the procedure's
  declarations. Each public result remains an independent mutable map.
- Default-object setup indexes module arrays and relevant declared types once
  per pass, then checks procedure declarations, parameters, return names and
  implicit locals through the existing type environment. Ordinary scalar
  declarations no longer require two full environment iterations per procedure.

Alternating isolated comparisons on all 1,552 ROneCOne procedures, with three
warm-up and seven measured rounds, produced these medians:

| Setup workload | Saved implementation | Current implementation |
|---|---:|---:|
| First field-variable scope lookup per procedure | 27.07 ms | 1.22 ms |
| Local/module object candidate collection | 86.44 ms | 5.84 ms |
| String constant scope | 74.42 ms | 1.02 ms |
| Default-object procedure visitor construction | 185.76 ms | 7.94 ms |

These measure setup, not entire rule execution. Candidate/constant outputs
matched between versions. Full native worker passes retained the current
16-finding baseline digest. An observed three-round worker median was
1,881.67 ms, including 1,801.47 ms in ROneCOne, after an initial 3,095.11 ms pass.
An integration run after the first two changes took 3,639 ms initially and
14 ms cached, including a 2,559 ms ROneCOne request and a 628 ms host timer gap.

All 456 focused tests across 26 files passed, including field/local/parameter
shadowing, independent mutable string-scope results, symbol snapshot changes,
class arrays, DefType declarations, and Word default members. Compilation and
the integration probe passed. A later integration run overlapping CPU-heavy
benchmark work took 16,750 ms initially and 62 ms cached; concurrent measurements
are unsuitable for an end-to-end speedup claim.
The final integration probe, with this investigation's CPU-heavy benchmarks
stopped, took 3,674 ms initially and 16 ms cached, including 2,668 ms in the
ROneCOne request, a 496 ms maximum host timer gap and a 7 ms panel render.
The alternating full comparison matched the current digest in every round;
its measured medians were 2,353.50 ms before and 2,307.00 ms after, but early
rounds overlapped the integration run and varied widely. Use the isolated
comparisons above for the per-setup claims, not this noisy aggregate difference.

## Shared token keys, written names and statement-variable subsets

A further warm profile identified repeated start-key formatting and first
statement-value queries. This pass adds three small improvements:

- Weakly cache the text key of immutable constant/default token arrays shared
  by many procedure starts. Canonical map keys retain their existing format.
- Keep a four-entry source-only cache of written names, shared by project
  indexing and module diagnostics. Edited source gets a fresh scan; recent
  entries are promoted and older sources are evicted.
- Follow module-variable values by looking up eligible variables in the
  reaching state. Each statement previously iterated that entire state,
  including hundreds of module constants which cannot be eligible variables.

Alternating isolated comparisons, with three warm-up and seven measured rounds:

| Workload | Saved implementation | Current implementation |
|---|---:|---:|
| Keys for 1,552 starts with 261 shared constants | 40.47 ms | 34.08 ms |
| Repeated workbook written-name scan, including source reconstruction | 23.19 ms | 2.66 ms |
| First value queries for all 9,369 leaf statements | 63.75 ms | 36.22 ms |

Keys, all 2,399 written names and all statement-value facts matched their
respective saved implementations. The first-query comparison times queries
only, excluding reader construction and result hashing. These are component
measurements, not an end-to-end speedup claim.

All 137 focused tests across 15 files passed, covering cache eviction,
reconstructed and edited source, symbol/conditional activity changes, local
shadowing, default module values and separate numeric/variant assignments.
Compilation and type checking passed. The final three-round native benchmark
reported a 2,007.85 ms full-pass median, including 1,904.17 ms in ROneCOne;
its first pass was 3,303.83 ms. All rounds retained the current 16-finding
digest `4b2d985d6c6f5a131a0097ed4ab7a53daf277a36266b633409cb6dea0cab9f5b`.

The final sidebar integration probe, with benchmarks stopped, took 3,614 ms
initially and 15 ms cached, including 2,500 ms in ROneCOne, a 626 ms maximum
extension-host timer gap and an 8 ms panel render. An earlier probe in this
pass took 3,388 ms initially and 14 ms cached. Cold runs still take several
seconds; ROneCOne semantic analysis remains the largest part of step 5's wait.
The profile also showed statement token lookup, receiver/member resolution
and garbage collection as remaining costs. The member-resolution files are
being edited independently, so this pass leaves that work intact.

## Indexed private and class member queries

A fresh warm profile still attributed about 47 ms of sampled self time to the
private-name linear search across its separate call sites, plus 14 ms to
`projectClassMemberAt`. This pass indexes both immutable project member arrays
with weak maps. Public lookups retain the first case-insensitive match; private
lookups still defer to an accessible public surface with the same name.
Replacing an owner's member array gets a fresh index.

A saved implementation captured all 4,220 actual project-member queries from
one workbook analysis, including 2,104 private-member queries. Replaying those
queries alternately through the saved and current resolvers, after three
warm-up rounds and across seven measured rounds, took median 55.71 ms before
and 3.42 ms after. Every lookup result matched (digest
`7983eda362006bc9f4a859c738f9a087ab9af923409fd933ea901aaa01790e0f`).
The recorded contexts contain warmed receiver caches; this measures repeated
member queries, not complete analysis or initial index construction.

All 229 member-related tests across eight files passed, including mixed-case
duplicates, class versus document behavior, replaced public/private arrays,
public visibility winning over a private name, and bounded name reads across
repeated misses in thousand-member lists. Compilation and type checking passed.
All full native passes retained the current 16-finding diagnostic digest.

The final integration probe passed at 7,551 ms cold and 81 ms cached, with
5,318 ms in ROneCOne and an 853 ms maximum host timer gap. The full native
median was 5,216.52 ms. Both workloads and unrelated test setup slowed together
while other Node processes were active; these aggregate observations do not
establish an end-to-end improvement or regression. The alternating query
comparison above establishes the removed scan cost. The remaining declaration
lookup inspected in this pass only scans procedure-local names and parameters;
it was left unchanged.

## Equal reaching-start recognition

The warm profile continued to show canonical start-key sorting/serialization
as a substantial cost. Fresh symbol snapshots rebuild equal starts, each with
hundreds of module constants. The walk cache now first compares entries by
name and shared token identity or existing token-text keys. On equality it
adopts the new start, so later queries use the identity fast path. Previously
serialized starts retain the canonical lookup path; unequal states still use
the original sorted key. Source and conditional activity guards remain in
place.

One workbook pass captured 7,333 walk calls using 1,713 distinct start maps.
An alternating comparison replayed the same start/identity patterns, with
fresh map/token-array copies per round and empty bodies to isolate cache
recognition from walking code. After three warm-up rounds and across seven
measured rounds, median recognition time dropped from 75.32 ms to 12.21 ms.
This is a cache component measurement, not total analysis time.

All 127 focused tests across 12 files passed, including reordered equal starts,
changed values, source/activity changes, shared token ownership and adoption
without repeated entry comparisons. Full native analysis retained the current
16-finding digest in every round. Compilation and type checking passed.
The sidebar integration probe passed at 6,210 ms cold and 22 ms cached,
including 4,448 ms in ROneCOne and a 776 ms maximum host timer gap. The full
native median was 5,398.15 ms, with a 6,753.05 ms first pass. Full-run timings
remain noisy under concurrent runtime load and are not used for a speedup
claim. The remaining step-5 wait is still largely ROneCOne semantic analysis.

## Fixed-array return-function setup

The fixed-array rule rebuilt the complete parameterless-function list for
each procedure and reconstructed its local hiding set for each name. It now
collects active, known array-returning parameterless functions once per rule
pass and applies each procedure's local/parameter hiding set once. A module
without known array-returning functions skips this candidate scan entirely.
Return-shape lookup is also held once per pass, preserving its existing cache.

Alternating complete-rule runs on ROneCOne, after three warm-up rounds and
across seven measured rounds, dropped from median 60.05 ms to 19.76 ms with
identical emitted findings. Unlike the earlier setup-only comparisons, this
times the entire fixed-array rule on warmed analyzer facts; it does not time
the full analyzer.

All 283 tests across ten array/cache files passed, including direct
parameterless-function subscripts, explicit return-array subscripts on a
function with arguments, local/parameter shadows and changed conditional
declarations. Compilation and type checking passed. The full native benchmark
retained the current 16-finding digest in every round and reported median
1,877.13 ms overall, including 1,784.76 ms in ROneCOne; its first pass took
2,811.45 ms. This run followed the earlier runtime contention easing, so its
difference from the preceding noisy full runs cannot be attributed to this
change alone.

The final sidebar integration probe passed at 3,407 ms cold and 14 ms cached,
including 2,412 ms in ROneCOne, a 479 ms maximum extension-host timer gap and
a 6 ms panel render. The remaining nearby module-wide scans inspected were
already performed once per rule/pass and were left unchanged. Step 5 still
waits primarily for the large ROneCOne semantic pass.

## Shared overflow constant layers

The overflow rule copied the complete folded module/project constant map
for each procedure, then removed local and parameter names. Its consumers
only perform keyed reads. It now shares the folded constant layer and uses
a small hiding set for procedure-local names. Procedures without hidden names
read the existing map directly. Local constant folding still runs before
that hiding layer, retaining its access to module constant values.

Alternating complete overflow-rule runs on ROneCOne, after three warm-up
rounds and across seven measured rounds, dropped from median 104.63 ms to
84.79 ms with identical findings. This measures the warmed overflow rule,
not the full analyzer. All 183 tests across eleven files passed, including
module values, local/parameter shadows, unfolded local constants, local
constant expressions using module constants and isolation between procedures.
Compilation and type checking passed.

The fresh CPU profile in this pass was heavily distorted by concurrent Node
CPU/memory activity and is not used to rank hotspots or claim timing gains.
The final full native benchmark retained the current 16-finding digest in
every round, with a noisy median 5,124.25 ms and first pass 8,334.93 ms.
The subsequent sidebar integration probe passed at 3,045 ms cold and 14 ms
cached, including 2,102 ms in ROneCOne, a 492 ms maximum extension-host timer
gap and a 5 ms panel render. These differing full-run observations reinforce
that the component comparison, rather than previous full-run times under
different load, supports this change's performance claim.

## Reproduction

Read-only native worker pipeline, preserving project references and sheet
metadata, with diagnostic hashes checked across rounds:

```powershell
node scripts/benchmark-run-analysis.mjs 'F:\GitHub\xlide\xlide_vscode_testing\ROneCOne_Delegates_Demo.xlsm' --rounds=5
```

Each measured round uses fresh worker state for a full analysis; parser and
lexer caches warm in-process. This is not the completed-result cache benchmark.
The script reports first/full median times and per-module durations/counts.

Actual VS Code command, cached repeat and extension-host responsiveness:

```powershell
$env:XLIDE_ANALYSIS_PERF_WORKBOOK = 'F:\GitHub\xlide\xlide_vscode_testing\ROneCOne_Delegates_Demo.xlsm'
npm run compile
.\node_modules\.bin\vscode-test.cmd --code-version 1.139.1 --grep 'Run analysis performance'
```

The integration test skips unless the environment variable is set. It checks
successful performance traces, a cached repeat below two seconds, and no
extension-host timer gap above two seconds. These are broad regression guards,
not cross-machine first-run latency guarantees.
