# Interface refactor parsing preflight

`implementInterface` parsed the entire implementing class before finding its `Implements` names or validating the requested interface. No class AST is needed to report no interface, ambiguous choice, an unimplemented requested interface, missing interface source or an interface without public members.

Move the class parse after those checks. Public interface members are still parsed before the empty-interface check. Generation and already-complete checks still parse both relevant modules and retain the same signatures, edits and refusal messages. There is no new cache or response wait. The project-wide command still gathers its project sources; these measurements cover the pure analyzer refactor only.

Eight new tests include six no-interface work-count cases at 10/1,000 procedures across LF/CRLF/CR, the remaining refusal paths and generation/already-implemented controls. Seven cases fail on baseline. Types, 23 focused tests and the full suite pass: 14,740 tests, 33 skipped, 758 passing files and seven skipped files, 84.98 seconds.

Complete baseline comparisons match 179,676 refactor results across 8,556 sources and 25,668 newline-varied documents. Seven input variants include the original corpus as both implementing and interface source, injected interface choice/membership checks, missing/empty interfaces and generation. Only implementInterface.ts is replaced with baseline code; parser and lexer implementations stay identical. This preserves existing behavior rather than claiming every generated refactor is semantically valid for every possible interface.

## Reproduction and limits

Run `node scripts/benchmark-interface-refactor-preflight.mjs` independently in baseline/current/current/baseline order; use `--baseline=b162fab19aca1ef7879c8271c504193f28c9b036` for the baseline. Add `--full-reparse` to change every procedure name on each round so the body must be parsed again. Default trailing-comment edits permit incremental parser reuse. The warmed no-interface control parses the class outside the clock.

Node 24.18.0, AMD Ryzen 7 9800X3D; three warmups, nine measured rounds. Source construction is excluded; the complete refactor call is measured, with exact result validation after timing. Stable interface source is shared between rounds. No editor command, popup paint or actual private-workbook latency claim. Scanning for Implements still scales with source size.

### Changed procedure bodies

10,000 procedures; milliseconds per complete result. Ranges are the medians of two independent runs.

| Result | Baseline | Repair |
| --- | --- | --- |
| noImplements | 35.66450–41.48190 | 0.13720–0.14240 |
| ambiguous | 38.20700–38.93510 | 0.12410–0.12570 |
| wrongInterface | 36.37550–39.98280 | 0.13050–0.13090 |
| missingSource | 37.03260–37.12690 | 0.13810–0.15520 |
| emptyInterface | 36.75090–37.30880 | 0.12400–0.13390 |
| success | 40.04250–42.40340 | 38.68010–41.14820 |
| already | 43.02780–45.79700 | 39.65560–40.17730 |
| warmNoImplements | 0.09110–0.09200 | 0.09110–0.09170 |

### Trailing-comment edits with incremental reuse

10,000 procedures; milliseconds per complete result. Ranges are the medians of two independent runs.

| Result | Baseline | Repair |
| --- | --- | --- |
| noImplements | 2.25800–2.28140 | 0.10940–0.11560 |
| ambiguous | 2.20520–2.65510 | 0.09840–0.10050 |
| wrongInterface | 2.18840–2.24220 | 0.09850–0.10400 |
| missingSource | 2.20620–2.28480 | 0.10260–0.14320 |
| emptyInterface | 2.20270–2.25030 | 0.09550–0.09940 |
| success | 2.76960–2.79350 | 2.94610–2.96910 |
| already | 2.81720–2.89250 | 2.84660–2.90710 |
| warmNoImplements | 0.04920–0.04930 | 0.04530–0.04570 |

Successful generation, already-complete and fully warmed controls show run-to-run variation; parser deferral does not remove their required work. Tiny results can be mixed. Maximum changed-body no-interface samples were 55.6820–62.9744 ms before and 0.1529–0.1711 ms after. Incrementally reused no-interface medians were 2.2580–2.2814 ms before and 0.1094–0.1156 ms after. No heap-byte measurement or general analyzer speedup is claimed.
