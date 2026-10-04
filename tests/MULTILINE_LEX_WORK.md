# Nearby edit lexer work

Two nearby edits can share one small changed span that contains an unchanged
physical line break. The old incremental lexer rejected any line break in that
span, causing completion, indexing or semantic requests to lex the entire large
module again.

The guard now preserves the ordered physical line-break sequence and limits
the changed span to 128 characters. It compares CRLF, CR and LF explicitly and
includes one unchanged character at each boundary, so edits that split or join
a CRLF pair retain the full-lexer fallback. The 16,384-character logical-window
limit and unchanged trailing-boundary checks still apply. The cached suffix
keeps its line coordinates because the physical break sequence is unchanged.

Fifteen public regressions cover eight coupled nonce/member revisions with
LF/CRLF/CR, immutable prior snapshots, exact token/trivia coordinates and prefix
identity, deterministic nearby edit histories, added/removed/converted breaks,
and CRLF boundary fission/fusion. The existing code fails the three bounded-work
cases by lexing the whole roughly 8 KB synthetic module. The candidate passes
121 focused lexer/semantic cases, including all fifteen new cases (one existing skip).

An ignored differential compared complete cached-token arrays against full
lexing for 16,000 generated public edits and 160 sampled edits in the private
large class. Both the unchanged baseline and candidate matched every complete
token array and preserved the previous snapshots. Source text was suppressed
on any mismatch.

Two interleaved complete-request trials timed indexing, project-type context,
completion's actual macro-name lookup and all four semantic collectors together.
They used identical dependencies with separate old/candidate lexer bundles,
40 warm and 50 fresh requests per variant. Complete results matched for each
revision across variants. All setup code inside the stated request path was
included in the total time.

| Nearby edits | Original | Revised |
| --- | --- | --- |
| Trial 1 median / p95 / maximum | 46.15 / 61.49 / 92.53 ms | 20.13 / 26.16 / 27.30 ms |
| Trial 2 median / p95 / maximum | 46.86 / 63.03 / 90.82 ms | 20.37 / 27.59 / 44.58 ms |
| Trial 1 indexing median | 30.27 ms | 6.77 ms |
| Trial 2 indexing median | 29.81 ms | 6.84 ms |
| Maximum freshly lexed input | 969,467 characters | 77 characters |
| Freshly lexed input over 50 revisions | 48,473,325 characters | 3,825 characters |

Single-line controls remained similar: total medians 20.58/20.20 versus
21.25/20.03 ms. Actual line-break changes retained full lexing and similar
48.11/49.19 versus 49.39/50.51 ms medians. Individual control tails do not all
improve; the measured gain applies to nearby edits preserving physical breaks.

The matching diagnostic workbook/editor run passed all twelve integration cases
and 64 fresh-source recovery cycles. Compared with the split-cache baseline,
host full-class misses fell from 37 to two parse requests. Completion's 35 full
misses disappeared, and semantic requests continued to avoid full-class lexing.
Diagnostics logged only source lengths and static stacks. Their timings are
not controlled UI-speed evidence.

The ordinary build passed type checking and all 15,410 unit tests in 788 files
(33 existing skips across seven files). Editor validation initially passed
38 of 39 cases: the polling acceptance test for `.cez` to `.ce` recovery failed.
That small fixture does not reach the changed large-module lexer path. A
build-only diagnostic run showed correct pending-caret recovery and passed
all 39 cases; the subsequent ordinary build also passed all 39. No unrelated
recovery change was made, and this first failure remains an open investigation
if it reproduces.

The unprofiled native workbook run passed all twelve cases, including 64
fresh-source recovery cycles and four mouse hovers. Median / p95 / maximum:

| Visible operation | Latency |
| --- | --- |
| Backspace | 17 / 35 / 48 ms |
| Suggestion menu recovery | 112 / 125 / 136 ms |
| Typing | 16 / 46 / 48 ms |
| Unmatched-prefix dismissal | 1 / 12 / 13 ms |
| Mouse hover | 380 / 406 / 406 ms |

The final ordinary-code capture passed all twelve cases, 1,000 fresh-source
recovery cycles and 62 mouse hovers while profiling both processes. Median /
p95 / maximum:

| Visible operation | Latency |
| --- | --- |
| Backspace | 17 / 34 / 82 ms |
| Suggestion menu recovery | 114 / 136 / 184 ms |
| Typing | 16 / 47 / 59 ms |
| Unmatched-prefix dismissal | 1 / 12 / 53 ms |
| Mouse hover | 378 / 388 / 413 ms |

Both profiles, timing metadata, complete observations and the exact bundle/map
were retained locally and SHA-256 verified. The host capture lasted 518.47
seconds. Lexer-cache work took 3.32 seconds inclusive and macro-name lookup
2.72 seconds; semantic tokens still took 14.63 seconds, module-symbol
construction 14.72 seconds, project-member surfaces 11.85 seconds and casing
19.08 seconds. Inclusive categories overlap and must not be added together.
The worst Backspace window approximately overlapped 19.7 ms of renderer garbage
collection and 13.5 ms of completion work. Alignment uses the recorded profile
stop time and duration; these overlaps do not establish causation.

The native runs used pinned VS Code 1.139.1, the disposable-copy workbook
harness, fresh sources, word suggestions disabled and navigation diagnostics.
CPU profiling was disabled for 64 cycles and enabled for 1,000. The measured
build contains this guard over base commit `00e53764`; the source map and bundle
were preserved before any further code changes. The original workbook was not
modified, and private source and captures remain in ignored local artifacts.

These runs verify recovery and responsiveness under the recorded workload;
they are not a controlled UI comparison with earlier builds. The stalling goal
remains open for the remaining slow frames, repeated context/symbol work and
the initial small-fixture recovery failure.
