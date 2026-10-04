# Shared diagnostic held-object facts

Several diagnostic rules request the same procedure's callback-free object/Collection state. A whole-module probe on f006e788 records three requests with identical source, procedure, symbols and activity. Previously each ran the full dataflow walk.

The cache is weakly owned by symbol snapshot and parsed procedure. Each procedure retains only its most recent source/activity pair, with no older-version chain. Facts are read-only snapshots by the existing HeldObjects contract. A changed source, symbol snapshot or activity identity collects fresh facts. Callback-driven queries always collect again, preserving side effects and answers from changing callback contexts.

Validation:

- Seven new regressions: four baseline work-count failures and three passing isolation controls. Fixed: all 18 focused checks across shared facts, historical snapshots and argument laziness pass; types pass.
- Repeated equivalent consumers (2/10/100) perform one body walk. Activity off/on/off, with two consumers each, performs three walks and preserves A/B/A facts. Source changes preserve A/B/A; separate symbol snapshots collect independently. The same callback returning A then B is called twice, and default facts do not acquire callback results.
- All 8,298 complete diagnostic and internal-error arrays match f006e788: 8,256 corpus sources plus 42 generated cases across LF/CRLF/CR, activity on/off, collection additions, aliasing, changing members, labels, branches and calls. 13,485 diagnostics; zero internal errors. AST nodes and lexer tokens/trivia are frozen.

Benchmark: node scripts/benchmark-held-object-facts.mjs --baseline=f006e788 --rounds=9 and the same command without baseline. Three warmups, nine measured rounds, baseline/candidate/candidate/baseline order, Node 24.18.0 on Ryzen 7 9800X3D. Every statement's classes and Collection items are independently checked outside timing in all three consumers. Symbol identities are fresh each round; parsing, lexing and symbol construction are outside this component measurement.

| Statements | Before median ms | After median ms |
| --- | --- | --- |
| 100 | 0.111–0.129 | 0.060–0.077 |
| 1,000 | 0.573–0.763 | 0.209–0.332 |
| 10,000 | 6.378–8.063 | 2.185–2.325 |

These measurements show equivalent default-fact reuse, not a claim about total diagnostic or editor latency. Rules with different value-resolution callbacks remain separate.
