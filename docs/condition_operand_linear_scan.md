# Linear condition operand discovery

The helper searched its growing output for every bare name, although the forward token loop visits each index once. Only the mutually exclusive statement-head forms can seed an earlier operand. Retain that head index and compare it directly, preserving head precedence, operand order and all existing forms. No cache or extra index collection is needed.

## Validation

Against baseline `ad47e91f3540b51f0f5f75e0a2498d953b7419f2`, three work regressions fail and 21 semantic controls pass. After repair, type checking, 61 focused tests and the full suite pass (774 files / 15,133 tests; 7 files / 33 tests skipped; 73.59 seconds). Controls cover all statement heads, Not/logical/IIf operands, bracketed and qualified names, array access, head precedence, malformed input and frozen tokens. A seeded 20,000-expression comparison preserves exact operand outputs; 8,256 oracle sources across LF, CRLF and CR preserve 24,768 complete diagnostic/failure objects.

## Measurement

Duplicate comparisons for 100/1,000 independent operands fall from 4,950/499,500 to zero. Run `node scripts/benchmark-condition-operand-scan.mjs`, optionally with `--baseline=ad47e91f3540b51f0f5f75e0a2498d953b7419f2`. Node v24.18.0 on AMD Ryzen 7 9800X3D 8-Core Processor           . Sequential ABBA trials use three warmups and nine measured samples. Tokens and module AST warming are outside the clock. Every helper output is independently specified; the complete module uses a Boolean flag plus a used array local to enable the condition-values caller, and must return exactly no diagnostics or internal failures. Continued lines stay within physical limits. Counters run before timings, affecting JIT state.

| Scope | Operands | Before median ms | After median ms |
|---|---:|---:|---:|
| helper | 1 | 0.0024–0.003 | 0.0032–0.0033 |
| complete-module-diagnostics | 1 | 1.4739–1.4869 | 1.4875–1.5079 |
| helper | 100 | 0.037–0.0477 | 0.0175–0.0202 |
| complete-module-diagnostics | 100 | 1.7273–1.8013 | 1.7553–1.762 |
| helper | 1000 | 2.2057–2.2101 | 0.0761–0.0769 |
| complete-module-diagnostics | 1000 | 8.6626–8.8835 | 7.1026–7.2687 |

The large helper and 1,000-operand complete fixture improve. One-operand helper/module timings are slightly slower, and the 100-operand complete-module ranges overlap. This is a narrow component and fixture improvement; no general editor responsiveness, cold-parser or heap improvement is claimed.
