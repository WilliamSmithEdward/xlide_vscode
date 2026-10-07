# Procedure hover signature reuse

Procedure hover formatting filtered every child symbol to find parameters on every request. A procedure's children include local declarations, so repeated hovers repeated body-sized work even when the read-only editor snapshot was unchanged.

The immutable signature string now follows its owned `VbaSymbol` through a WeakMap. Only symbols returned by the read-only `editorModuleSymbols` snapshots reach this private formatting path. A source, module-name or module-kind change selects a different graph. Entries do not keep discarded symbols alive. Hover details, spans, documentation and result objects remain request-local. Existing local hover formatting is preserved, including its treatment of Optional/ByVal/ParamArray and array parameters; this does not substitute the richer external declaration formatter.

Eight new tests cover deterministic child reads, all three newline styles, changed source, renamed modules, result mutation, property kinds and parameters without explicit types. Six work-count cases fail on the baseline: ten requests with 1,000 locals read 10,010 children; the repair reads zero after the first signature. Seventy-six focused tests pass, with two skipped; types pass.

## Measurement

Run `node scripts/benchmark-hover-procedure-signatures.mjs` independently in baseline/current/current/baseline order, supplying `--baseline=95dd31be71e9ee95dda0ce7d2b8a0e6b50a1e47d` for baseline runs. Node 24.18.0 on AMD Ryzen 7 9800X3D; three warmups and nine measured rounds. Source construction, symbol construction, tokenization and hover index construction are outside the clock. Warm rows measure 100 complete hover requests per round and validate each full result after timing. First-procedure rows include first signature formatting, with the other caches warmed by a local hover.

Milliseconds per complete query; ranges below are medians from the two independent runs:

| Locals | Query | Baseline | Repair |
| --- | --- | --- | --- |
| 0 | Warm procedure | 0.00095–0.00108 | 0.00070–0.00088 |
| 100 | Warm procedure | 0.00135–0.00140 | 0.00060–0.00085 |
| 10,000 | Warm procedure | 0.05934–0.06302 | 0.01388–0.01463 |
| 0 | First procedure | 0.00170 | 0.00200–0.00390 |
| 100 | First procedure | 0.00240–0.00260 | 0.00290–0.00340 |
| 10,000 | First procedure | 0.07330–0.07630 | 0.06750–0.07530 |
| 10,000 | Warm local control | 0.01501–0.02030 | 0.01640–0.02112 |

The maximum 10,000-local warm-procedure batch-average sample was 0.07806–0.13307 ms before and 0.02641–0.03195 ms after. First-query/tiny and local controls show overhead or mixed results. The string cache adds a small retained value per hovered procedure; heap bytes are not measured. Complete query work elsewhere still varies with source size. This is a component improvement, not a claim about cold lexing, sustained edits, tooltip painting or resolution of #964/#985.
## Complete validation

Full suite: 14,732 tests pass, 33 skipped, 757 passing files and seven skipped files (89.00 s). Types pass. Baseline/current comparison matches 228,240 complete hover results across 8,556 sources and 25,668 newline-varied documents. It checks an injected function, up to six original procedure/Declare positions, repeated queries and module name/kind changes. Only resolveHover.ts is replaced with baseline code in the comparison, so lexer/parser/symbol builders are identical.
