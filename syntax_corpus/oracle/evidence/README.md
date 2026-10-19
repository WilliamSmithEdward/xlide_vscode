# Gap-contexts measurements

`gap_contexts_historical_20261016.json` preserves the complete rq17/rq18 input and raw log text with SHA-256 hashes. Its 44 selected records are the gap-contexts compile observations; the canonical corpus asserts that phase only. Secondary run results remain evidence, without changing compile outcomes. The historical runner revision, reference versions and VBE selections were not recorded. An accepted `Is Nothing` constraint does not establish an exact Variant declaration.

The `../probes/gap_contexts_review_20261019.json` file retains the original 405 observe-only compile inputs. `../probes/gap_contexts_designer_20261019.json` retains ten corresponding observe-only designer/predeclared inputs. The successful 405-case and 43-case Windows batches are retained in `gap_contexts_review_20261019.json` and `gap_contexts_validation_20261019.json`; their actual outcomes are promoted in the canonical corpus. Original inputs remain `observe` to preserve the checkpoint hashes. `../probes/gap_contexts_completion_20261019.json` adds fresh triplets for the 53 historically covered builtin names, declaring-module DefType controls and invalid-call neighbors. Those observations remain pending until their successful results are retained.

The original ten setup descriptions are retained in `../probes/gap_contexts_manual_20261019.json`; that file is a planning document, not runner input.

The worker creates actual UserForm controls and imports class/procedure attributes. Setup readback failures are infrastructure failures, never accepted compilation evidence. Compile selection capture occurs after dismissing the compile dialog and includes the actual code-pane source; it is supplementary evidence, not a claim of exact live-dialog highlighting.

Run from the repository root on Windows:

```text
node syntax_corpus/oracle/run_excel_vbe_oracle.mjs --cases syntax_corpus/oracle/probes/gap_contexts_review_20261019.json --checkpoint test-results/gap-contexts-review-20261019.jsonl --json
node syntax_corpus/oracle/run_excel_vbe_oracle.mjs --cases syntax_corpus/oracle/probes/gap_contexts_designer_20261019.json --checkpoint test-results/gap-contexts-designer-20261019.jsonl --json
```

Create `test-results` before these commands. Checkpoint records retain actual outcomes, messages, stages and measured host builds. Infrastructure failures must be resolved before fixture promotion.

`gap_contexts_harness_20261019.json` preserves the exact coordinator and two worker snapshots used by the first two batches. The first snapshot predates selection RCW cleanup; the second includes it. Contract tests reproduce every checkpoint hash from the exact harness bytes, host build and original input, and compare promoted module sources and outcomes.

The measured parity test checks compile verdicts and messages across the retained fresh inputs. Compact post-dismissal selections are compared for diagnostic codes changed here. Unchanged diagnostics can retain a narrower range than Excel; whole-line captures do not establish a narrow live-dialog highlight.
