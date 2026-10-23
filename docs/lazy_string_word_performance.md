# Lazy string-word facts during index construction

Index construction eagerly scanned every cached token and extracted words from
all string literals, even for completion/hover consumers that never query
string references. The extension-host profile showed this scanner among its
largest self-cost entries during large-class typing.

Build literal-word facts only when `stringLiteralWords()` is called. Retain
facts for unchanged module source, invalidate replacement/removal, and keep the
project union in the existing revision cache. The union follows module source
order, preserving results/order after case-insensitive module replacement.
Runtime reachability through strings still uses the complete literal-word set.

Regressions assert zero extraction across 80 module insertions before the query,
exact per-module work on the first query, no repeated work, retention across
unrelated edits and metadata changes, removal/re-addition, empty results,
Unicode, escaped quotes, attributes and exclusion of comments. Existing symbol
index and dead-code tests continue to exercise the query's consumers.

An alternating warm component comparison on the 969,328-character private
snapshot, 21 batches of 30 fresh index constructions, measured median batch
means of 4.309277 ms before and 2.318747 ms after. Maximum batch means were
4.951013/2.895490 ms. Module symbols/lexing were warmed beforehand. First-demand
literal extraction still has its own cost; this is not an overall editor
speedup claim.

Both actual-workbook/renderer runs passed all eight cases, with the second run
CPU-profiled. The unprofiled run had command typing/Backspace medians 2.37/2.79 ms
and maxima 18.54/137.66 ms, with a 76.28 ms maximum host heartbeat delay. Across
24 visible recovery cycles, typing/Backspace maxima were 40/36 ms and menu
updates had a 2 ms median, with 109 ms on the first show. Updated declaration
hover took 49.46 ms. The profiled run had command medians 2.42/1.81 ms and maxima
19.37/11.61 ms; visible maxima were 47/32 ms and updated hover took 22.66 ms.
Every recovery displayed Cells and every following miss hid it. The latest host
profile had no self samples in `stringLiteralWordsIn`.

Compilation, 59 focused index/dead-code tests and all 14,500 unit tests across
732 files passed, with 33 tests and seven files intentionally skipped. The
137.66 ms command outlier remains recorded; issues #964/#985 and the intermittent
latency goal remain open.
