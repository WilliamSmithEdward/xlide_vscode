# Reuse class value contributions after unrelated edits

`ProjectIndex.projectClassMembers` cached the whole project snapshot but called `classMemberValues` again for every class after any indexed module changed. Editing a standard module therefore rescanned unchanged class bodies even though their value facts depend only on their own source and symbols.

Reuse each class's value map through the existing per-module contribution cache. `setModule` and `removeModule` already invalidate that module's contributions, while unrelated edits retain them. The maps remain lazy and are only requested by diagnostic surfaces. Member rows are still built for each project snapshot, so editor surfaces omit value annotations and old snapshots retain their values. No new invalidation mechanism or cross-revision project snapshot is introduced.

## Validation

Eight new tests cover deterministic scan counts for 10/1,000 unchanged classes across LF/CRLF/CR, single-class edits, editor/diagnostic separation, removal, module-kind replacement and old snapshot preservation. Seven cases fail on the baseline. Type checking and 15 focused tests passed.

An independent baseline comparison against `bed0c445cfe658db9958640d563debc1e089a141` matched 102,672 complete surface queries across 8,556 sources. Each source was queried with and without value facts during initial indexing, caller edits, class edits, removal, restoration as a UserForm, and restoration as a class. Previously returned snapshots stayed unchanged.

## Measurement

Run `node scripts/benchmark-class-value-contributions.mjs`, optionally with `--baseline=bed0c445cfe658db9958640d563debc1e089a141`. Complete `projectClassMembers` calls include surface construction, value facts and returned array copying. Parsing, project setup, edits and independent output assertions are outside the timer. The lexer is warmed; cold means a fresh index contribution cache. Separate before/after/after/before processes use three warmups and nine samples. Node 24.18.0, AMD Ryzen 7 9800X3D. Each class has two fields and twenty functions.

| Classes / mode | Before medians (ms) | After medians (ms) |
| --- | --- | --- |
| 1 / caller edit | 0.0565–0.0739 | 0.0282–0.0309 |
| 100 / caller edit | 2.7082–2.7424 | 1.0792–1.0819 |
| 1,000 / caller edit | 27.4498–28.1003 | 11.3644–11.6669 |
| 1 / cold contribution cache | 0.1198–0.1262 | 0.1428–0.1470 |
| 100 / cold contribution cache | 3.1728–3.4676 | 3.4410–3.6765 |
| 1,000 / cold contribution cache | 33.0261–33.5020 | 33.2045–33.2290 |
| 1,000 / editor caller edit | 9.8663–10.5052 | 10.1344–10.1683 |

This removes repeated class scans after unrelated edits. Cold queries pay contribution-cache overhead; the one-class cold control was slightly slower. Member construction still scales with project size. Retained value maps add memory per queried class; heap bytes and live-editor latency were not measured. No universal speedup is claimed.

Full repository suite: 753 files passed, seven skipped; 14,692 tests passed, 33 skipped (75.62 seconds).
