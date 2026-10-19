# Gap-contexts measurements

`gap_contexts_historical_20261016.json` preserves the complete rq17/rq18 input and raw log text with SHA-256 hashes. Its 44 selected records are the gap-contexts compile observations; the canonical corpus asserts that phase only. Secondary run results remain evidence, without changing compile outcomes. The historical runner revision, reference versions and VBE selections were not recorded. An accepted `Is Nothing` constraint does not establish an exact Variant declaration.

The `../probes/gap_contexts_review_20261019.json` file contains 405 pending compile observations. `../probes/gap_contexts_designer_20261019.json` contains ten corresponding designer/predeclared observations. Expectations in these files remain `observe` until a successful measured result is promoted. The original ten setup descriptions are retained in `../probes/gap_contexts_manual_20261019.json`; that file is a planning document, not runner input.

The worker creates actual UserForm controls and imports class/procedure attributes. Setup readback failures are infrastructure failures, never accepted compilation evidence. Compile selection capture occurs after dismissing the compile dialog and includes the actual code-pane source; it is supplementary evidence, not a claim of exact live-dialog highlighting.

Run from the repository root on Windows:

```text
node syntax_corpus/oracle/run_excel_vbe_oracle.mjs --cases syntax_corpus/oracle/probes/gap_contexts_review_20261019.json --checkpoint test-results/gap-contexts-review-20261019.jsonl --json
node syntax_corpus/oracle/run_excel_vbe_oracle.mjs --cases syntax_corpus/oracle/probes/gap_contexts_designer_20261019.json --checkpoint test-results/gap-contexts-designer-20261019.jsonl --json
```

Create `test-results` before these commands. Checkpoint records retain actual outcomes, messages, stages and measured host builds. Infrastructure failures must be resolved before fixture promotion.
