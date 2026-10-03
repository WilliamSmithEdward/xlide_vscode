# Undeclared-variable declaration setup

An undeclared assignment can offer a `Dim` insertion edit. Previously each edit searched the same procedure body for its insertion point, found indentation again, and scanned the source to detect line endings. A long declaration section followed by many missing assignments repeated the body search for every finding.

The rule now lazily prepares one insertion site per procedure and detects line endings once per rule call. Declared targets and read-only findings still do not prepare edits. The cache lives only within the current invocation; subsequent calls use their own source and AST.

## Measurement

Run `node scripts/benchmark-undeclared-fix-sites.mjs --rounds=15`. Compare with `--baseline=00d4dd4e`. The baseline loader substitutes only the previous undeclared rule; parsing and symbol binding happen outside the timed region. Each fixture has N object declarations followed by N `Set` assignments. Missing targets produce N declaration fixes; declared targets produce none. The object assignment avoids expression type inference, isolating insertion setup within the rule. LF source also exercises full-source line-ending detection.

Node 24.18.0, AMD Ryzen 7 9800X3D, three warmups, 15 samples. Values are median milliseconds:

| N | Missing before | Missing after | Declared before | Declared after |
| --- | ---: | ---: | ---: | ---: |
| 100 | 0.319 | 0.322 | 0.191 | 0.225 |
| 1,000 | 3.141 | 1.770 | 1.062 | 1.283 |
| 3,000 | 17.871 | 3.890 | 3.788 | 3.870 |

At 3,000 missing assignments the isolated rule is about 4.6 times faster. This stress fixture is not an end-to-end editor timing. Small measurements fluctuate; the declared control retains the existing guarded path.

## Validation

An operation-count regression limits procedure-body reads for 1,000 declarations and missing assignments. Output regressions cover multiple procedures, LF/CRLF, comments, indentation and distinct inferred types. A guard test verifies that declared assignments and read-only diagnostics do not detect line endings for edits.

Validation: 67 targeted tests and the type check passed. The complete suite passed 578 files / 12,196 tests (13 skipped). Complete diagnostic messages, spans and edit data matched the previous rule for 1,000 generated LF/CRLF fixtures with multiple procedures, varying declarations and object/scalar/read-only references.
