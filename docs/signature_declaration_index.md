# Signature declaration indexing

A signature request looked up the local procedure and Declare separately for documentation, external details and the signature itself. Each lookup walked every parsed module member until its match. A target at the end of a 1,000-procedure module caused 40,040 member visits across ten requests despite the parse already being cached.

One weak parsed-module index now builds separate case-insensitive procedure and Declare maps in one pass. The first declaration of each kind wins, preserving the old searches and procedure-versus-Declare precedence. Source changes produce another parsed module/index. Only declaration nodes are retained; documentation is extracted from the current moduleSource and project/registry/host metadata is still resolved on each request. Returned signature results are not cached.

Memory is proportional to the indexed module's declaration names. The weak module key does not create a source history beyond the parser's existing ownership. This improves repeated local declaration lookup; parsing, call-site resolution, documentation and host/project lookup costs remain.

Validation against af326a9227d507a489627bfb270ac4a9c32e2e66:

- Types and 41 focused tests across two files pass.
- Three work-count tests fail on the baseline (440/4,040/40,040 member visits for 10/100/1,000 unrelated procedures over ten requests); the changed-source/first-match control passes. After the fix, each module is indexed once; all complete expected signature fields match.
- Frozen-input differential over 8,256 corpus sources plus 36 generated cases: all 33,278 complete SignatureInfo/undefined results match, including 20,023 resolved signatures. Generated cases cover LF/CRLF/CR, Long/String parameters, procedures, Declares, duplicates, runtime fallback, project fallback and alternate moduleSource. ASTs and lexer tokens/trivia are frozen. The harness rethrows the resolver's normally caught internal errors; no exceptions occurred.

Full suite: 730 files passed, seven skipped; 14,463 tests passed, 32 skipped. No failures.

Benchmark: node scripts/benchmark-signature-declarations.mjs --baseline=af326a9227d507a489627bfb270ac4a9c32e2e66 --rounds=9 and without baseline. Baseline/candidate/candidate/baseline order, three warmups/nine measured rounds; Node 24.18.0 on Ryzen 7 9800X3D. Each round uses a unique source, with the target after the unrelated procedures. Both versions warm the initial request/parse/index. Timing covers 100 repeated requests, including call-site/signature work; construction and priming are excluded. Every complete SignatureInfo is independently checked.

| Unrelated procedures | Before median ms / 100 requests | After median ms / 100 requests |
| --- | --- | --- |
| 1 | 0.331–0.338 | 0.280–0.296 |
| 100 | 0.391–0.422 | 0.207–0.226 |
| 1,000 | 2.532–2.558 | 0.182–0.223 |

These are repeated signature resolver measurements, not cold-index or total editor latency. One index pass is still required per new parsed module, and large bodies/documentation/member metadata can cost more than this synthetic fixture.
