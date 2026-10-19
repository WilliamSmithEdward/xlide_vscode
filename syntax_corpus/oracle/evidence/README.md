# Gap-contexts measurements

`gap_contexts_historical_20261016.json` preserves the complete rq17/rq18 input and raw log text with SHA-256 hashes. Its 44 selected records are the gap-contexts compile observations; the canonical corpus asserts that phase only. Secondary run results remain evidence, without changing compile outcomes. The historical runner revision, reference versions and VBE selections were not recorded. An accepted `Is Nothing` constraint does not establish an exact Variant declaration.

The `../probes/gap_contexts_review_20261019.json` file retains the original 405 observe-only compile inputs. `../probes/gap_contexts_designer_20261019.json` retains ten corresponding observe-only designer/predeclared inputs. The successful 405-case and 43-case Windows batches are retained in `gap_contexts_review_20261019.json` and `gap_contexts_validation_20261019.json`; their actual outcomes are promoted in the canonical corpus. Original inputs remain `observe` to preserve the checkpoint hashes. `../probes/gap_contexts_completion_20261019.json` adds fresh triplets for the 53 historically covered builtin names, declaring-module DefType controls and invalid-call neighbors. Its 199 successful observations are retained in `gap_contexts_completion_20261019.json` and promoted in the canonical corpus.

The original ten setup descriptions are retained in `../probes/gap_contexts_manual_20261019.json`; that file is a planning document, not runner input.

The worker creates actual UserForm controls and imports class/procedure attributes. Setup readback failures are infrastructure failures, never accepted compilation evidence. Compile selection capture occurs after dismissing the compile dialog and includes the actual code-pane source; it is supplementary evidence, not a claim of exact live-dialog highlighting.

Run from the repository root on Windows:

```text
node syntax_corpus/oracle/run_excel_vbe_oracle.mjs --cases syntax_corpus/oracle/probes/gap_contexts_review_20261019.json --checkpoint test-results/gap-contexts-review-20261019.jsonl --json
node syntax_corpus/oracle/run_excel_vbe_oracle.mjs --cases syntax_corpus/oracle/probes/gap_contexts_designer_20261019.json --checkpoint test-results/gap-contexts-designer-20261019.jsonl --json
```

Create `test-results` before these commands. Checkpoint records retain actual outcomes, messages, stages and measured host builds. Infrastructure failures must be resolved before fixture promotion.

`gap_contexts_harness_20261019.json` preserves the exact coordinator and two worker snapshots used by the first two batches. The first snapshot predates selection RCW cleanup; the second includes it. Contract tests reproduce every checkpoint hash from the exact harness bytes, host build and original input, and compare promoted module sources and outcomes.

The measured parity test checks compile verdicts and messages across the retained fresh inputs. Compact and whole-line post-dismissal selections are compared for diagnostic codes changed here. Unchanged diagnostics can retain a narrower range than Excel; no capture here establishes live-dialog highlighting.

`gap_contexts_builtin_constraints_20261019.json` links all 168 current runtime functions to fresh ordinary, Is and TypeOf compile observations on build 16.0.20430.20182. The coverage test requires every catalog entry and checks the matrix against retained checkpoint outcomes/messages/stages. These constraints are deliberately separate; several intrinsic forms have different TypeOf grammar from Is.

`gap_contexts_qualified_20261019.json` retains all 168 VBA-qualified triplets and eight reserved-name designer controls. The matrix links these separately: qualified compiler productions differ from bare intrinsic productions, and VBA.LBound/VBA.UBound are unavailable members. Their ordinary rejection cannot establish a return constraint.

`gap_contexts_comparison_neighbors_20261019.json` retains 90 bare/qualified InStr, InStrB and StrComp controls with String, Variant, Null and Object arguments. `gap_contexts_receiver_shadows_20261019.json` retains 14 source-module/library receiver controls. The first six fresh batches contain 1,263 successful compile observations, with no infrastructure failures or skipped cases.

Existing duplicate-source fixtures remain as named matrix controls so historical IDs and audit references stay stable. The original worktree scratch manifests `oracle_added_cases.json` and `oracle_recheck.json` are not part of this change.

`gap_contexts_ci_neighbors_20261019.json` retains eight exact neighbors from older assertions: seven VBA library member selections and a qualified source call missing its required caption argument. These confirm the dot/suffix selections and `Argument not optional` wording on the same Microsoft 365 build. Across all seven batches there are 1,271 successful compile observations (549 rejected, 722 accepted).
