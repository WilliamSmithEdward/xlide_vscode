# Indexed class-member token copies

Each indexed class-member read copied the entire statement token array before calling matchParenFrom. The helper accepts readonly tokens and only reads them. N sibling indexed reads in one statement therefore copied approximately 7N squared token references, although each parenthesis search examines only its short argument list. Pass the original token array directly.

## Reproduction and results

Run `node scripts/benchmark-class-indexed-token-copy.mjs`, then `--baseline=77c4bf53e523dec521a493e5a67013af5f4cc40d`. The baseline overrides only this rule. Each fixture has one local As New Class1 and N `actor.M(1)` reads in a logical Debug.Print statement, separated by semicolons. M independently holds Nothing. Thirty reads per physical line keep line widths below 1,023; the largest fixture uses 20 lines and 19 continuations.

Sequential before/after/after/before on Node 24.18.0, Ryzen 7 9800X3D. Three warmups and nine measured calls; ranges of the two run medians in milliseconds.

| Reads | Scope | Before ms | After ms |
| ---: | --- | ---: | ---: |
| 1 | rule | 0.0057–0.0072 | 0.0055–0.0064 |
| 1 | complete-module-diagnostics | 1.2008–1.2651 | 1.2037–1.224 |
| 100 | rule | 0.8574–0.8617 | 0.0757–0.0768 |
| 100 | complete-module-diagnostics | 4.3238–4.688 | 3.5224–3.6368 |
| 600 | rule | 27.3176–27.7994 | 0.1605–0.1663 |
| 600 | complete-module-diagnostics | 49.8594–49.8866 | 22.1609–22.9928 |

Token-array iteration counts through the actual public rule: 9/70,200/2,521,200 token references for 1/100/600 reads before, zero after. The tests bound token-array iteration while independently expecting every exact diagnostic, message and original span, preventing an early-return shortcut from satisfying the work bound.

One-read ranges overlap. AST/symbol preparation is outside direct-rule timing. Complete-module calls include analyzer setup and internal-error assertions; full-result equality checks are outside timing. The harness independently requires N object-variable-not-set diagnostics and no internal errors. No editor, cold-start or heap speedup is claimed. Nested argument scanning and diagnostic label construction are unchanged; this removes the redundant whole-statement copies.

## Validation

- Type checking and 86 focused tests pass.
- Eight new tests: three work regressions fail on baseline; five controls pass on both (nested argument spans across LF/CRLF/CR, unmatched lists, indexed assignment suppression).
- Full suite: 15,307 tests across 782 files pass; 33 tests and seven files skipped.
- 20,000 generated indexed-read rule results and 24,768 complete module diagnostic/error queries over 8,256 oracle sources across LF/CRLF/CR match the previous rule.
