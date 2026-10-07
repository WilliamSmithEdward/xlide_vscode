# Straight-line context cache ownership

The straight-line walk used a string made from initial value tokens as its memo key. It ignored the call effects and declaration facts attached to the start. Different contexts with equal initial values could therefore reuse old ByRef results or reachability. The table also strongly retained a complete walk for each distinct value string while its parsed body stayed live.

The table now uses weak initial-state keys. A hit requires the same source, conditional activity, call-effect reader and declaration-fact reader. Each retained start keeps only its latest walk. The analyzer already retains starts by bound symbol/procedure identity; declaration readers are now retained under that same symbol/procedure ownership with a source check so rules share the unchanged context without rebuilding declaration maps. The old serialized/sorted start-key map is removed.

Initial maps and attached readers describe immutable analysis facts; this does not create a cache contract for mutation behind an unchanged reader closure. Replacement readers and different start identities are distinct contexts.

Validation against 49ed5c90fc93893cc8f26beb0d9e5c299e3523e4:

- Three independent regressions fail before the fix: equal-valued starts with different ByRef effects, replacement effects on one retained start, and replacement declaration facts changing reachability. All pass after the fix, including 0/2/0 and 0/1/0 transitions.
- The retained-start value/reachability sharing control passes. Types and 25 focused tests across five files pass, including symbol/activity lifetimes, module-state reuse, held-object sharing and guarded branches.
- Frozen-input differential: 8,256 corpus sources plus 36 generated cases (three line endings, four string-comparison modes and 1/10/100 statements). All 8,292 complete diagnostic arrays match (14,095 findings), with zero internal errors. Original/equal-copy diagnostics also match independently within each version. ASTs, tokens and trivia are frozen. The independently expected changed-context regressions establish the intentional behavior difference; this corpus comparison establishes unchanged normal analysis behavior.

Full suite: 729 files passed, seven skipped; 14,459 tests passed, 32 skipped. No failures.

Retention probe: node --expose-gc scripts/benchmark-straight-line-contexts.mjs --baseline=49ed5c90fc93893cc8f26beb0d9e5c299e3523e4 and without baseline. Node 24.18.0 on Ryzen 7 9800X3D. Each argument start is independently checked at the first statement; repeated requests must return the identical memoized result. The procedure and parsed body stay live while caller references to starts/results are released. Five explicit GCs run across event-loop turns. Build-only instrumentation exposes the old strong table size; WeakRefs independently check whether starts survive.

| Released contexts | Before strong entries / surviving starts | After strong entries / surviving starts |
| --- | --- | --- |
| 100 | 100 / 100 | 0 / 0 |
| 1,000 | 1,000 / 1,000 | 0 / 0 |

This proves ownership of released starts, not a guaranteed GC schedule, heap byte reduction or total editor timing. Results a caller deliberately retains remain live. Unchanged retained starts still share one walk; distinct starts no longer share solely because their serialized values coincide.
