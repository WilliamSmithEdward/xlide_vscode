# Bound Dim-initializer quick-fix context scanning

The `dim-initializer` quick fix walked every preceding physical line to decide whether the declaration was inside a procedure. Its existing heuristic changes state only at a procedure opener or closer. An edit near the end of a large module therefore repeated unrelated prefix work whenever quick fixes were requested.

Search backward for the nearest preceding boundary instead. The same opener/closer patterns decide the result; malformed nested declarations retain their existing behavior. Blank physical lines are skipped. LF/CRLF/CR and BOM handling remain compatible, including the existing LF edit EOL for CR-only source. No source, token or result cache is added.

## Validation

Eight new tests cover deterministic boundary-test counts after 10/1,000 unrelated procedures across three newline forms, independent complete edits, malformed boundary precedence, BOMs, blank lines, unsafe declaration refusals and fresh context after edits. Six work-count cases fail on the baseline: 3,001 opener tests after 1,000 procedures become one. Types and all 53 focused quick-fix tests passed.

Independent comparison against `58e867cd20bcab3dc0a05dc5f91a82019420c8df` matches 385,020 complete action results across 8,556 oracle/generated sources. Each prefix was queried with three newline forms, three procedure contexts and five safe/unsafe declaration variants, with suppression actions enabled.

## Measurement

Run `node scripts/benchmark-dim-initializer-context.mjs`, optionally with `--baseline=58e867cd20bcab3dc0a05dc5f91a82019420c8df`. Complete quick-fix queries include context scanning, statement tokenization, EOL detection and edit construction. Inputs and independent complete expected actions are outside the timer. Three warmups, nine rounds, twenty queries per sample; separate before/after/after/before processes. Statement lexer warmed. Node 24.18.0, AMD Ryzen 7 9800X3D.

| Workload | Before medians (ms) | After medians (ms) |
| --- | --- | --- |
| After 100 procedures | 0.01057–0.01078 | 0.00144–0.00146 |
| Middle of 10,000 procedures | 0.46923–0.47172 | 0.00585–0.00587 |
| After 10,000 procedures | 0.93642–0.93820 | 0.00592–0.00609 |
| Module-level refusal after 10,000 procedures | 0.92079–0.92921 | 0.00010 |
| First target before 10,000 procedures | 0.00587 | 0.00594–0.00610 |
| 30,000-line first procedure body | 0.89847–0.91870 | 0.88209–0.89310 |

Tiny/early controls are mixed or slightly slower. The last-position fixed maximum batch-average sample was 0.00925 ms, and the long-body maximum was 1.18180 ms. Work still scales with the distance to the nearest boundary; a large first procedure can require a long scan. These are returned-action measurements, not complete editor latency or cold lexer measurements.

Full repository suite: 755 files passed, seven skipped; 14,715 tests passed, 33 skipped (80.98 seconds).
