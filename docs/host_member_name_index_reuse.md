# Reuse the host member-name index

`hostModelIndex.memberNames` already contains every lowercase member name from
each type's `rawByLowerName`. `isHostMemberName` built a second Set from both
`byLowerName` and `rawByLowerName`, although object-access `byLowerName` keys are
a subset of the raw keys. This duplicated a whole-model sweep and retained the
same broad names again in a second WeakMap cache. Its caller is late-binding
Friend-member analysis; semantic tokens already use the shared broad lookup.

Remove the second WeakMap and export `isHostMemberName` as an alias of
`isHostMemberNameAnywhere`. Both existing export names remain callable with the
same arguments/default model and membership semantics. Events remain included;
globals alone are not members. Model ownership stays with the existing weak
index, with no source or project-context cache added.

## Validation

Baseline `b3494c2262071ca45e7e1bc4b989f54b8cb90efd`: two failing work
regressions, two passing result/ownership controls. After the shared index is
primed, a 1,000-type frozen fixture receives 5,000 extra `Set.add` calls on the
baseline and zero with the repair. Four new tests cover work, all member kinds,
unknown/global names, separate model ownership and either initialization order.

Type checking and 37 focused tests passed. Full suite: 748 files / 14,647 tests
passed; 7 files / 33 tests skipped (71.64 seconds).

Differential: 155,346 membership query/order comparisons across nine real host
models and 25,879 summed unique names, plus 700 duplicate-case/event/method
custom-model queries, match the baseline. Actual models and metadata are deeply
frozen; either lookup can initialize first. Excel, Word, PowerPoint, Access, VB6,
unmodelled-host empty metadata and three merged models are covered. Complete
analyzer outputs and failure reports match across 8,256 oracle sources with a
project Friend-member context.

## Measurement

Run `node scripts/benchmark-host-member-name-reuse.mjs`, optionally with
`--baseline=b3494c2262071ca45e7e1bc4b989f54b8cb90efd`.
Separate baseline/fixed/fixed/baseline processes on Node 24.18.0 / Ryzen 7 9800X3D
used three warmups and nine rounds. A positive and negative lookup are timed as a
pair. Real metadata loads and shallow model construction are outside the timer.
Fresh model identity makes each cold or after-shared-index round independent;
the shared index is primed outside the timer only for the latter. Warm pairs
are averaged over 10,000 batches. Independent complete pair assertions are
outside the timer.

| Host, first pair after shared index | Baseline median ms | Fixed median ms |
| --- | --- | --- |
| Excel (3,414 unique names) | 0.3638–0.3795 | 0.0010–0.0030 |
| Word (3,273) | 0.3103–0.3204 | 0.0005–0.0006 |
| Access (2,318) | 0.3121–0.3137 | 0.0005–0.0006 |
| VB6 (437) | 0.0287–0.0324 | 0.0002–0.0005 |
| Excel + Word (5,091) | 0.5977–0.6156 | 0.0007–0.0008 |

Fully cold Excel pairs improved 3.2849–3.3858 to 2.7822–2.7863 ms; Access
2.2098–2.2661 to 1.9000–1.9643; merged Excel/Word 4.3909–4.5637 to
3.8174–3.9471. Word was mixed (2.9674–3.0577 to 2.6491–3.1844), and VB6
was slower in this control (0.2489–0.3483 to 0.3780–0.3900).

Warm pairs show small mixed changes: Excel was 65.76–69.96 ns to 88.61–97.60 ns;
Word roughly 62 to 64–65 ns; Access roughly 63–73 to 64–65 ns; VB6 52–64 to
36–37 ns; merged Excel/Word 61–63 to 64–66 ns. No universal warm speedup is
claimed. The largest fixed after-shared-index Excel pair was 0.0278 ms. Host
index construction remains; model loading, heap bytes and whole-editor latency
were not measured. The structural saving is one redundant Set of broad names
and its WeakMap entry per used model.
