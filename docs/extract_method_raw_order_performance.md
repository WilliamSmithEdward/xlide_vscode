# Extract Method raw local ordering performance

Baseline: published main `fff67474250947c76f9574e79bd1500ae3f3a27c`.

Extract Method had already moved raw source searches out of its sorting comparator, but still ran one independent `source.indexOf(displayName)` for every touched local. Large prefixes were searched repeatedly. The raw behavior matters: matching is case-sensitive, includes comments and strings, matches substrings, and preserves stable ties. A token/whole-name occurrence index cannot replace it.

Large ordering queries now share a literal matcher and prefix dictionary. The matcher advances from the previous match, retaining overlapping matches; the dictionary records every name starting at that position, including shorter prefixes. Resolved names leave the active matcher at geometrically spaced rebuilds, preventing common short names from matching throughout unrelated padding. Each query owns its data. There is no retained source/name cache.

The direct path remains for fewer than 64 touched locals, or a workload estimate below 16,000,000 (local count × maximum declared-name end). A declared spelling occurs by that offset, so a large unrelated suffix does not inflate the estimate. These are measured crossover heuristics, not a claim of an optimal threshold for every input. Single-local extraction still skips ordering searches. Dictionaries exceeding 65,536 escaped-pattern code units use native searches before allocating the prefix dictionary. SyntaxError/RangeError from the regexp engine also falls back to native searches for unresolved names.

## Reproduction

```powershell
node scripts/benchmark-extract-method-raw-order.mjs --baseline=fff67474250947c76f9574e79bd1500ae3f3a27c
node scripts/benchmark-extract-method-raw-order.mjs
```

Run sequentially in baseline/fixed/fixed/baseline order. Node 24.18.0, AMD Ryzen 7 9800X3D, esbuild 0.28.2. Three warmups and nine timed runs per fixture. Ranges below span two medians. Public extraction includes the existing parse/lookup cache behavior; the second scope also applies all edits. Repeated identical sources are used. Work instrumentation is untimed and restored before timing.

The selection writes locals before reading them, so their declarations move into a zero-parameter helper. Input physical lines are below 1023 characters, including name-order comments and long-name controls. The small case has no large padding. Medium/large prefixes contain approximately 93/930 KB of comments; common-short uses repeated `a` with a local named `a`. Large-suffix puts padding after the procedure. The dictionary-limit control has 400 long local names.

## Measurements

| Placement | Locals | Scope | Baseline ms | Fixed ms |
| --- | ---: | --- | ---: | ---: |
| small | 1 | public-extraction | 0.0340–0.0391 | 0.0339–0.0361 |
| small | 1 | extraction-and-apply-edits | 0.0131–0.0131 | 0.0128–0.0129 |
| small | 10 | public-extraction | 0.0397–0.0405 | 0.0429–0.0437 |
| small | 10 | extraction-and-apply-edits | 0.0602–0.0672 | 0.0578–0.0591 |
| small | 64 | public-extraction | 0.1990–0.2067 | 0.1752–0.2686 |
| small | 64 | extraction-and-apply-edits | 0.1190–0.1192 | 0.1311–0.1837 |
| small | 256 | public-extraction | 0.6050–0.6150 | 0.6112–0.6874 |
| small | 256 | extraction-and-apply-edits | 0.6006–0.6513 | 0.7195–0.7535 |
| small | 1000 | public-extraction | 4.1360–4.3337 | 2.7215–2.8499 |
| small | 1000 | extraction-and-apply-edits | 3.5230–3.5529 | 2.5805–2.6138 |
| medium-prefix | 1 | public-extraction | 0.0843–0.1559 | 0.0812–0.1033 |
| medium-prefix | 1 | extraction-and-apply-edits | 0.0860–0.0875 | 0.0855–0.1059 |
| medium-prefix | 10 | public-extraction | 0.1079–0.1215 | 0.1045–0.1069 |
| medium-prefix | 10 | extraction-and-apply-edits | 0.1092–0.1258 | 0.1085–0.1103 |
| medium-prefix | 64 | public-extraction | 0.2261–0.2425 | 0.2549–0.3279 |
| medium-prefix | 64 | extraction-and-apply-edits | 0.2544–0.2586 | 0.2417–0.3999 |
| medium-prefix | 256 | public-extraction | 0.8409–0.8571 | 0.5473–0.7487 |
| medium-prefix | 256 | extraction-and-apply-edits | 0.8267–0.8586 | 0.5346–0.6595 |
| medium-prefix | 1000 | public-extraction | 5.0889–5.2837 | 2.2436–2.2600 |
| medium-prefix | 1000 | extraction-and-apply-edits | 4.8012–4.8666 | 2.0217–2.4164 |
| large-prefix | 1 | public-extraction | 0.8031–0.8697 | 0.8304–0.8814 |
| large-prefix | 1 | extraction-and-apply-edits | 0.9095–0.9382 | 0.9142–0.9345 |
| large-prefix | 10 | public-extraction | 0.9373–1.1282 | 0.9796–1.1022 |
| large-prefix | 10 | extraction-and-apply-edits | 1.0247–1.0770 | 1.0454–1.0573 |
| large-prefix | 64 | public-extraction | 1.7452–1.9168 | 1.0168–1.1436 |
| large-prefix | 64 | extraction-and-apply-edits | 1.7770–2.2640 | 1.0574–1.1576 |
| large-prefix | 256 | public-extraction | 4.4641–4.5999 | 1.2705–1.3083 |
| large-prefix | 256 | extraction-and-apply-edits | 4.5786–4.6176 | 1.4378–1.5316 |
| large-prefix | 1000 | public-extraction | 17.0160–17.3919 | 2.9296–3.4812 |
| large-prefix | 1000 | extraction-and-apply-edits | 16.8925–17.1943 | 2.8833–3.0230 |
| common-short | 1 | public-extraction | 0.8333–0.8644 | 0.8587–0.9732 |
| common-short | 1 | extraction-and-apply-edits | 0.9306–0.9364 | 0.9305–0.9781 |
| common-short | 10 | public-extraction | 0.9685–0.9787 | 0.9745–1.0406 |
| common-short | 10 | extraction-and-apply-edits | 1.0523–1.0633 | 1.0731–1.0835 |
| common-short | 64 | public-extraction | 1.8434–1.8739 | 1.3079–1.5396 |
| common-short | 64 | extraction-and-apply-edits | 1.7986–1.8026 | 1.3696–1.5099 |
| common-short | 256 | public-extraction | 4.4935–4.5609 | 1.6931–1.8523 |
| common-short | 256 | extraction-and-apply-edits | 4.5536–4.6052 | 1.8640–1.8859 |
| common-short | 1000 | public-extraction | 16.6129–16.9199 | 3.1852–3.2522 |
| common-short | 1000 | extraction-and-apply-edits | 16.7860–16.8484 | 3.2290–3.2524 |
| large-suffix | 1 | public-extraction | 0.7805–0.8008 | 0.7619–0.7938 |
| large-suffix | 1 | extraction-and-apply-edits | 0.8740–0.8928 | 0.8601–0.8855 |
| large-suffix | 10 | public-extraction | 0.8227–0.9283 | 0.7976–0.8335 |
| large-suffix | 10 | extraction-and-apply-edits | 0.8652–1.0565 | 0.8393–0.9513 |
| large-suffix | 64 | public-extraction | 0.8540–0.8635 | 0.8283–0.8507 |
| large-suffix | 64 | extraction-and-apply-edits | 1.0914–1.1045 | 0.9191–0.9293 |
| large-suffix | 256 | public-extraction | 1.1838–1.2373 | 1.1676–1.4215 |
| large-suffix | 256 | extraction-and-apply-edits | 1.2426–1.2850 | 1.2860–1.3714 |
| large-suffix | 1000 | public-extraction | 4.1934–4.3552 | 2.7731–3.0292 |
| large-suffix | 1000 | extraction-and-apply-edits | 4.2891–4.2928 | 2.7717–2.7846 |
| long-name-limit | 400 | public-extraction | 7.9056–7.9115 | 7.8065–7.8795 |
| long-name-limit | 400 | extraction-and-apply-edits | 7.8945–8.0214 | 7.8776–8.2318 |

At 1,000 locals with a large prefix, independent full-source `indexOf` searches fall from 1,000 to zero; forward matching supplies the positions. Public extraction improves from 17.0160–17.3919 ms to 2.9296–3.4812 ms. Extraction plus text application improves from 16.8925–17.1943 ms to 2.8833–3.0230 ms. Common-short, large-suffix, and several other large cases also improve. Smaller cases and dictionary fallback overlap or can be slower. No universal, cold-start, heap, native-editor, tree, or tab-switch latency improvement is claimed.

## Validation

15 dedicated tests: seven batching work bounds fail on the baseline, eight controls pass; all 15 pass after the repair. Exact public edited output is checked against native first-substring ordering. Coverage includes comments/strings, case-sensitive Unicode, same-position prefixes, internal overlaps, common short names, single/small queries, unrelated large suffixes, long-name dictionary limits, and forced SyntaxError/RangeError fallbacks.

20,000 private raw-position comparisons against native `indexOf` match exactly for nonempty unique names, including literal regexp metacharacters and UTF-16 surrogate pieces. Complete public result and applied text comparisons match for 500 generated batched modules and 24,768 corpus environments (8,256 sources under LF/CRLF/CR). Every corpus extraction succeeds. Four ABBA trials independently assert the exact replacement, helper declaration order/body, title and rename span, then compare complete result/applied-text snapshots and digests. Zero differences.

Final type checking passed. Full suite: 15,983 tests across 810 files passed, with 33 tests and seven files skipped, in 82.19 seconds. This validates this repair; it is not completion evidence for the analyzer/repository audit.
