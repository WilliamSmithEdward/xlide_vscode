# Referenced-library own-entry lookup

`libraryTypeNames` indexed ordinary objects using a project's reference name.
Unknown names `constructor` / `CONSTRUCTOR` returned `Object`, and `__proto__`
returned `Object.prototype`, through the inherited cache entries. The declaration
rule then attempted `names.has(...)` and reported an isolated internal failure
instead of finishing its type-name checks.

Normalize the spelling once and require an own entry in the generated names
table before consulting its cache. These names now return undefined like other
unread libraries. Known sets and their cache identity remain unchanged. Update
`dump-library-type-names.py` to emit the same lookup so regeneration preserves
the repair; type-library metadata is unchanged.

Validation on branch base `b3494c2262071ca45e7e1bc4b989f54b8cb90efd`:

- Six baseline failures: three inherited-name lookups, two actual declaration
  rule failures (captured via `onInternalError`), and duplicate normalization.
  One complete known-library cache control passed. All seven now pass.
- Type checking and 17 focused tests passed.
- Full suite: 748 files / 14,650 tests passed; 7 files / 33 tests skipped
  (78.72 seconds).
- Python AST syntax validation and exact generator lookup-tail comparison
  passed, without COM calls or regeneration of installed-library metadata.
- Differential: all 24 complete known-library set/case/cache-identity checks
  matched; 16,512 complete diagnostic and failure-report comparisons across
  8,256 oracle sources matched for previously valid reference inputs. Seventy-two
  corrected inherited-name queries matched the ordinary unknown-library control
  with no internal failures, across three EOLs and eight declared types.
  The component baseline `6059d89dc95f54d986a16eadea77d360859b2a20` has the
  identical lookup file as the validated branch base.

A successful lookup now normalizes its name once rather than twice. This is a
work-count cleanup alongside the rule-failure fix; no wall-clock speedup or
whole-analyzer latency claim is made.
