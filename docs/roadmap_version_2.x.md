# XLIDE Roadmap Version 2.x

Unified development roadmap for the current XLIDE workstream: realtime VBA
linting, language service hardening, type analysis, corpus coverage, and
Excel/VBE compatibility validation.

## North Star

XLIDE should make VS Code feel like a serious VBA IDE for real Excel projects:
fast while typing, deterministic in its diagnostics, honest about uncertainty,
and close to VBE behavior where VBE behavior is the relevant authority.

The goal is not to make VBA stricter than VBA. The goal is to catch real compile
errors early, offer useful guidance where appropriate, and avoid noisy or
heuristic diagnostics.

## Operating Principles

- Deterministic logic only. No heuristic analyzer behavior unless the operator
  explicitly approves it before implementation.
- Prefer no diagnostic over a guessed diagnostic.
- Red/error means a deterministic compile blocker or a construct XLIDE can prove
  invalid.
- Yellow/warning means XLIDE guidance, maintainability advice, or a soft risk
  that may still compile.
- Diagnostic language must match severity. Red/error diagnostics use
  authoritative wording such as "will fail", "will raise", or "is invalid" when
  the behavior is proven. Yellow/warning diagnostics use non-authoritative
  wording such as "may", "can", "risk", or "consider" because they represent
  guidance or uncertainty.
- For complicated work, create or update a roadmap before implementation.
- The Excel/VBE oracle is a discovery, debugging, and corpus-coverage tool, not
  a routine per-change test.
- The syntax corpus is evidence, not authority. Corpus cases may be incomplete
  or wrong and should be verified against the oracle when they drive analyzer
  behavior.
- Deterministic runtime-error rules may be red even when
  `vbeCompileEquivalent` is false, but only with focused runtime oracle evidence
  or an equally deterministic local proof.

## Current Baseline

### Landed

- Realtime diagnostics are split between structural linting and analyzer rules.
- Active diagnostics are catalogued in `src/analyzer/diagnostics/ruleMetadata.ts`.
- Rule metadata includes category, VBE compile-equivalence, and diagnostic-kind
  fields so compile-time and deterministic-runtime red squiggles stay separate.
- Same-module callable signatures are parsed for parameters and return types.
- High-confidence argument count and type mismatch diagnostics exist.
- Non-callable call statements are flagged, including bare variable statements
  that VBE Compile rejects.
- Runtime metadata is curated explicitly; parameter types are not inferred from
  parameter names.
- Inline documentation comments support descriptive metadata plus optional
  `type`, `unit`, and `value` hints.
- Excel/VBE oracle harness exists under `syntax_corpus/oracle/`.
- The wider unverified type-analysis backlog is tracked in
  `docs/type_analysis_corpus_coverage.md`.
- Oracle cleanup is coordinator-owned: the recorded disposable Excel PID is
  killed after each oracle case.

### Current Validation Layers

- Fast routine checks:
  - `npm.cmd test`
  - `npm.cmd run check-types`
  - `npm.cmd run compile`
- Optional oracle checks:
  - `npm.cmd run test:oracle:vbe`
  - Use only for VBE behavior discovery, corpus validation, or oracle harness
    changes.

## Workstream A: Diagnostic Policy and Metadata

Purpose: make every diagnostic self-describing and severity-safe.

- [x] Add `category` to every rule.
- [x] Add `vbeCompileEquivalent` to every rule.
- [x] Add `diagnosticKind` to every rule.
- [x] Document red/yellow/no-diagnostic severity policy.
- [x] Document diagnostic language policy: red is authoritative, yellow is
  advisory.
- [x] Add tests that assert every rule declares category and VBE equivalence.
- [x] Add a short diagnostic policy table to architecture docs.
- [ ] Use category/equivalence fields when reporting workbook lint summaries.

Definition of done:

- Every emitted diagnostic can be classified as VBE compile-equivalent,
  deterministic runtime error, runtime-risk, XLIDE guidance, or
  project/model-specific.
- New diagnostics cannot be added without category, VBE equivalence, and
  diagnostic-kind metadata.

## Workstream B: Syntax Corpus and Oracle Coverage

Purpose: keep corpus fixtures useful without letting stale assumptions become
truth.

- [x] Add Excel/VBE oracle harness.
- [x] Capture VBE compile popups as the primary oracle signal.
- [x] Keep oracle out of routine per-change verification.
- [x] Kill the recorded disposable Excel PID after each oracle case.
- [x] Mark corpus fixtures with provenance:
  - spec-derived
  - VBE-oracle-verified
  - observed but not asserted
  - pending verification
- [x] Add an oracle-result update workflow for promoted corpus cases.
- [x] Review current syntax corpus for cases that predate the oracle and mark
  them pending until verified.
- [x] Add focused oracle fixtures for newly debated syntax edges only.
- [x] Add a type-analysis corpus coverage matrix for pending, partial, missing,
  and verified type-analysis areas.
- [x] Complete full `syntax_corpus` digestion into managed backlog categories
  (syntax, realtime recovery, type analysis, runtime resolution, project
  binding, host behavior, completion context, UserForm/designer symbols,
  limits, and legacy edges). Treat every Markdown case as raw material until
  promoted through spec, oracle, or deterministic XLIDE-owned evidence. See
  `syntax_corpus/managed_backlog.md`.

Definition of done:

- Corpus-driven analyzer behavior has traceable provenance.
- If a corpus case and the oracle disagree, the case is treated as suspect until
  resolved.

## Workstream C: Phase 1 Type System Cleanup

Purpose: finish the conservative first slice before broadening inference.

- [x] Same-module procedure parameter and return metadata.
- [x] Literal and declared-variable argument type inference.
- [x] Named argument mapping before validation.
- [x] Runtime function participation from curated metadata.
- [x] High-confidence nonnumeric string literal mismatch diagnostics.
- [x] Add fixture coverage for every current compatibility edge:
  - [x] numeric literal to numeric parameter
  - [x] numeric string to numeric parameter
  - [x] nonnumeric string to numeric parameter
  - [x] Variant suppressing hard mismatch
  - [x] named argument order
  - [x] omitted required argument
  - [x] omitted optional argument behavior
- [x] Separate compile-equivalent argument errors from deterministic runtime
  errors in tests and metadata.
- [x] Split object argument mismatches into a compile-equivalent red diagnostic
  after focused oracle verification.
- [x] Add red deterministic-runtime-error diagnostic for nonnumeric string
  literals in numeric arithmetic expressions after focused oracle verification
  showed VBE Compile accepts the representative case but execution raises
  runtime error 13.
- [x] Promote deterministic nonnumeric string-to-number/Boolean assignment and
  argument coercion failures to red runtime-error diagnostics after focused
  runtime oracle verification.
- [x] Record the runtime type mismatch experiment matrix in
  `docs/vba_runtime_type_mismatch_oracle_matrix.md`.
- [x] Add compile-equivalent diagnostics for runtime functions used as `As` type
  names and `Set` assignments to known intrinsic scalar variables.

Definition of done:

- Phase 1 diagnostics have valid/invalid/unknown fixtures.
- Every hard error is deterministic and either VBE-equivalent or explicitly
  justified as XLIDE-invalid.

## Workstream D: Phase 2 Expression Return Types

Purpose: infer simple expression result types without guessing.

First vertical slice:

- [x] Same-module function call return type.
- [x] Nested same-module call return type.
- [x] Curated runtime conversion function return types:
  - [x] `CStr`
  - [x] `CDbl`
  - [x] `CCur`
  - [x] `CLng`
  - [x] `CBool`
- [x] Arithmetic result family inference for obvious numeric expressions.
- [x] String concatenation result inference for `&`.
- [x] Unknown or `Variant` expression operands suppress hard diagnostics.
- [x] Assignment type mismatch diagnostics after deterministic expression proof
  for scalar assignments.

Out of scope until proven:

- Locale-sensitive string coercion.
- Full operator overload/coercion lattice.
- Object member receiver chains.
- Flow-sensitive variable narrowing.

Definition of done:

- Expression inference only returns a concrete type when the source provides
  deterministic proof.
- Inference feeds argument validation and assignment validation without adding
  noisy false positives.

## Workstream E: Project-Wide Binder

Purpose: move from same-module checks to workbook-aware analysis.

- [x] Resolve unique exported standard-module `Sub`/`Function` signatures for
  cross-module argument count and type diagnostics.
- [x] Resolve public procedures across standard modules for module-qualified
  calls and ambiguous duplicate behavior in the first deterministic slice:
  ambiguous bare exported names stay silent, while `ModuleName.ProcedureName`
  resolves through the named standard module only.
- [ ] Model module-level visibility and shadowing.
- [ ] Resolve `As` type names against project classes, UDTs, enums, and host
  object types before flagging broad unknown type names.
- [ ] Resolve enums and enum members across modules.
- [ ] Resolve UDT names across modules.
- [ ] Add workbook-level fixture builder for project analysis tests.
- [x] Keep current project-signature diagnostics stable under module order
  changes.

Definition of done:

- Cross-module procedure calls are validated deterministically.
- Shadowing behavior is explicit and tested.

## Workstream F: Object and Member Types

Purpose: validate Excel/VBA object use where receiver type is known.

- [ ] Track `Set` assignments to known object types.
- [ ] Resolve class/document module member calls.
- [ ] Resolve curated Excel object model receiver chains.
- [ ] Add `member-not-found` only when receiver type is known.
- [ ] Add `set-required` and `set-forbidden` only where deterministic.

Definition of done:

- Object diagnostics do not guess from names.
- Host metadata has auditable provenance.

## Workstream G: Realtime Experience

Purpose: keep the live editor useful while the user is mid-keystroke.

- [ ] Suppress hard errors for incomplete expressions where VBE behavior is not
  yet deterministically knowable.
- [ ] Make diagnostic ranges precise and stable.
- [ ] Use metadata categories to tune Problems output and future filters.
- [ ] Keep signature help, hover, completion, and diagnostics sharing the same
  symbol/type model.

Definition of done:

- Typing incomplete code does not create avoidable noise.
- Completed invalid code produces deterministic, useful diagnostics.

## Workstream H: Developer VBA Test Runner

Purpose: let workbook developers write and run deterministic tests for their own
VBA code from XLIDE, using Excel COM as the execution host.

- [ ] Define an explicit VBA test discovery contract. Discovery must be
  annotation-driven or manifest-driven, not naming-heuristic-driven.
- [ ] Add a developer-facing `xlide.runVbaTests` command that runs selected
  tests through Excel COM.
- [ ] Run tests against a disposable workbook/session by default so test runs do
  not mutate the developer's open workbook unexpectedly.
- [ ] Reuse the workbook close/reopen/reset discipline from macro execution and
  warn when a workbook cannot be safely reopened in XLIDE's context.
- [ ] Add a small VBA assertion/support module or equivalent injected test
  runtime for assertions such as equality, truth, expected error, and expected
  no error.
- [ ] Support rich explicit test metadata:
  - tags/categories
  - skip reason
  - expected failure (`xfail`) reason
  - per-test timeout
  - owner or requirement id
  - expected error metadata
  - output/state assertions
- [ ] Support developer-friendly test selection and execution modes:
  - run all
  - run current test
  - run current module
  - include/exclude tags
  - rerun failed
  - fail fast
  - machine-readable automation mode
- [ ] Support setup/teardown patterns:
  - per-test setup and teardown
  - per-module setup and teardown
  - workbook/session setup and teardown
  - deterministic cleanup failure reporting
- [ ] Capture deterministic test results:
  - pass/fail/skip
  - expected failure (`xfail`)
  - unexpected pass (`xpass`)
  - compile errors
  - runtime errors, including error number and description
  - assertion failures
  - explicit test log/output written through the XLIDE test API
  - timeout and teardown failures
- [ ] Support tests that assert expected output, expected state, expected thrown
  error, and expected absence of errors.
- [ ] Return machine-readable JSON results for automation and render a concise
  Problems/Test Results view in VS Code.
- [ ] Keep the product test runner separate from the Excel/VBE oracle. The
  oracle validates XLIDE behavior; the test runner validates user VBA projects.
- [ ] Add fixture tests before enabling broad adoption.
- [ ] Fully document the downstream developer workflow before calling the test
  runner shipped:
  - how to author tests
  - how to mark test procedures explicitly
  - metadata for tags, skips, expected failures, timeouts, owners, and
    requirement ids
  - assertion API reference
  - expected-error and expected-output patterns
  - setup/teardown and test data patterns
  - filtering, rerun failed, and fail-fast workflows
  - workbook/session lifecycle
  - COM/Excel trust requirements
  - timeout and cleanup behavior
  - command palette and automation usage
  - result JSON schema
  - troubleshooting and known host limitations

Definition of done:

- A developer can author VBA tests, run them through Excel COM from XLIDE, and
  receive deterministic pass/fail/error/output results.
- The runner is opt-in, visible to the user, timeout-bounded, and safe against
  silent workbook mutation.
- The full workflow is documented for downstream workbook developers, including
  examples they can copy into real projects.

## Workstream I: Lint Suppression Directives

Purpose: give developers deterministic, VBA-comment-compatible control over
XLIDE lint diagnostics without changing VBA execution behavior.

- [x] Document the proposed suppression syntax in
  `docs/xlide_vba_lint_suppression_comments.md`.
- [ ] Parse suppression directives only from explicit XLIDE comment directives.
- [ ] Support module-level suppression.
- [ ] Support next-member suppression for `Sub`, `Function`, `Property`,
  `Type`, and `Enum` blocks.
- [ ] Support line-level and next-line suppression.
- [ ] Support paired arbitrary block suppression.
- [ ] Support optional diagnostic-code lists so developers can suppress one rule
  without hiding unrelated diagnostics.
- [ ] Add directive diagnostics for malformed, unbalanced, or unknown-code
  directives without guessing the user's intent.
- [ ] Preserve a suppressed-diagnostic count so ignored problems can be audited.
- [ ] Add unit tests for every directive scope, nesting edge, and malformed
  directive case.

Definition of done:

- Suppression behavior is lexical, deterministic, auditable, and fully
  compatible with VBA because directives are comments.
- Suppressions hide XLIDE diagnostics only; they never make invalid VBA valid
  and never affect COM/test/oracle execution.

## Workstream J: XLIDE Activity Bar and Sidebar Panel

Purpose: make XLIDE feel like a first-class VS Code extension with a polished
Activity Bar icon and a full sidebar command/status surface for workbook
development.

- [x] Add the planned sidebar design/spec to `docs/xlide_sidebar_panel.md`.
- [ ] Add a dedicated XLIDE Activity Bar container using standard VS Code
  contribution points.
- [ ] Create a slick monochrome SVG Activity Bar icon that follows VS Code's
  icon style: simple line geometry, mask-friendly, readable at 24px, and
  theme-neutral.
- [ ] Expand the current workbook/module explorer into a full XLIDE sidebar
  with deterministic status sections and command surfaces.
- [ ] Add setup health checks with pass/warn/fail/unknown states for:
  - active XLIDE workbook/project context
  - workbook source sync/export mapping
  - Excel COM availability on Windows
  - trust access to the VBA project object model
  - macro/security prerequisites relevant to run/test workflows
  - workbook open/reopen safety
  - lint engine readiness
  - VBA test runner readiness once implemented
  - optional metadata/doc-comment support
- [ ] Add primary action buttons:
  - lint current module
  - lint workbook
  - run current test
  - run all tests
  - export/sync modules
  - open/reopen workbook
  - refresh project
- [ ] Add polished secondary panels for:
  - Problems summary by severity
  - Test summary by pass/fail/skip/xfail/xpass
  - Recent XLIDE operations and logs
  - Setup recommendations and quick fixes
  - Workbook/module metadata
- [ ] Keep status checks deterministic. If XLIDE cannot prove a requirement is
  met or missing, show `Unknown` with a concrete action instead of guessing.
- [ ] Add telemetry-free local persistence for collapsed sections and selected
  workbook context.
- [ ] Add VS Code UI/integration tests or fixture-backed provider tests for the
  sidebar model before broadening the UI.

Definition of done:

- XLIDE has a recognizable Activity Bar icon and a sidebar that surfaces the
  main workbook workflows without needing command-palette spelunking.
- Setup health, lint, test, sync, and workbook actions are available from one
  coherent place.
- Sidebar status is deterministic, accessible, theme-safe, and does not depend
  on heuristic project guesses.

## Workstream K: Safety, Trust, and Recovery

Purpose: make workbook mutation, COM execution, and macro/test workflows safe,
auditable, and recoverable for real user projects.

- [ ] Define XLIDE's trust model for workbook read, write, run, and test
  workflows.
- [ ] Require explicit user action for operations that mutate workbook contents
  or execute VBA.
- [ ] Add backup/snapshot strategy before workbook mutation where practical.
- [ ] Add an audit trail for XLIDE writes:
  - module changed
  - command/tool that changed it
  - timestamp
  - source path/workbook path
  - success/failure
- [ ] Add crash/timeout recovery for COM operations.
- [ ] Add workbook lock/open-state checks before write/run/test operations.
- [ ] Warn when a workbook is open outside XLIDE's controllable context.
- [ ] Surface "what changed?" summaries after sync/write operations.
- [ ] Document trust-center, macro security, and VBA project access
  requirements without hiding them behind vague failures.

Definition of done:

- A developer can see when XLIDE is about to mutate or execute something, what
  it changed afterward, and how to recover from failed COM/workbook operations.
- Safety behavior is explicit and deterministic; XLIDE does not infer consent.

## Workstream L: Settings and Profiles

Purpose: give individuals and teams controlled configuration without weakening
the deterministic analyzer contract.

- [ ] Define workspace/project configuration for XLIDE.
- [ ] Separate team-shared settings from local-machine settings.
- [ ] Add rule severity overrides with guardrails:
  - error to warning only when the rule permits downgrade
  - warning to off
  - no override that converts unknown behavior into a red diagnostic
- [ ] Add enabled/disabled rule-set profiles.
- [ ] Add COM/test-runner settings:
  - Excel visibility
  - timeouts
  - workbook reset behavior
  - trusted test folders
  - result output path
- [ ] Add sidebar/profile UI for active configuration.
- [ ] Add configuration validation diagnostics for malformed settings.
- [ ] Document precedence: defaults, workspace, user/local, command override.

Definition of done:

- Teams can share stable XLIDE behavior through project settings while each
  developer can keep local COM/test paths and machine-specific choices local.
- Settings cannot make analyzer behavior heuristic or unverifiable.

## Workstream M: Code Actions and Quick Fixes

Purpose: turn deterministic diagnostics into useful repairs where XLIDE can
prove the edit is safe.

- [ ] Add code-action infrastructure for analyzer diagnostics.
- [ ] Add deterministic quick fixes for:
  - add missing block closer
  - insert required `Call` parentheses
  - add `Set` when the assignment target is known object type
  - remove `Set` when the assignment target is known scalar type
  - add missing required argument placeholder only when explicitly requested
  - add lint suppression comment for a selected diagnostic
  - create procedure stub for unresolved calls only when the target location is
    explicit
- [ ] Add source actions:
  - lint current module
  - export/sync current module
  - run current test when test runner exists
- [ ] Add tests for every generated edit, including formatting and range
  stability.
- [ ] Avoid quick fixes for uncertain, host-dependent, or incomplete-code cases.

Definition of done:

- Code actions are precise, previewable, and backed by tests.
- XLIDE offers no quick fix when the correct edit would require guessing intent.

## Workstream N: Performance and Scale

Purpose: keep XLIDE responsive on large workbooks and during active typing.

- [ ] Define performance budgets for:
  - keystroke diagnostics
  - module parse/analyze
  - workbook-wide lint
  - project index rebuild
  - sidebar health refresh
- [ ] Add incremental parsing/indexing where deterministic and measurable.
- [ ] Add cache invalidation rules for source text, workbook state, metadata,
  and project symbol graph changes.
- [ ] Ensure live diagnostics never block typing.
- [ ] Add large-workbook fixture coverage.
- [ ] Add stress tests for many modules, large modules, and many diagnostics.
- [ ] Add cancellation for long-running lint/test/sync operations.
- [ ] Add status/progress reporting for work that exceeds user-visible latency
  thresholds.

Definition of done:

- XLIDE remains responsive on large real-world workbooks, and slow paths have
  measurable budgets, cancellation, and progress.
- Performance improvements remain deterministic and do not skip diagnostics
  silently.

## Workstream O: Release and Documentation Polish

Purpose: make XLIDE understandable and trustworthy for users who were not part
of development.

- [ ] Add marketplace-ready assets:
  - icon
  - screenshots
  - feature bullets
  - short walkthrough media or GIFs
- [ ] Add getting-started documentation.
- [ ] Add a sample workbook/project repo for demos and regression examples.
- [ ] Add feature walkthroughs:
  - opening a workbook
  - editing modules
  - linting
  - running macros
  - running VBA tests
  - using doc comments and metadata
  - interpreting setup health
- [ ] Add troubleshooting playbooks for common Excel/COM/security failures.
- [ ] Add versioned changelog and release checklist.
- [ ] Add compatibility notes for Windows, macOS/Linux, WSL, and remote
  containers.

Definition of done:

- A new developer can install XLIDE, understand what works on their platform,
  run the main workflows, and troubleshoot common setup issues without reading
  source code.

## Workstream P: Support Bundle and Diagnostics

Purpose: make user support and self-debugging possible without guessing or
asking for sensitive workbook contents.

- [ ] Add "Export XLIDE Support Bundle" command.
- [ ] Include non-sensitive diagnostic data:
  - extension version
  - VS Code version
  - platform
  - enabled XLIDE settings
  - setup health states
  - recent XLIDE command log
  - lint/test summary counts
  - COM availability and failure categories
  - workbook metadata summary without source code by default
- [ ] Add opt-in inclusion of anonymized lint/test reports.
- [ ] Add opt-in inclusion of selected logs.
- [ ] Redact workbook paths or allow path redaction.
- [ ] Add a local "Copy Diagnostics" quick action for setup failures.
- [ ] Document what the support bundle contains before export.

Definition of done:

- Users can produce a useful local support snapshot without exposing workbook
  source by default.
- The bundle helps diagnose setup, COM, lint, test, and sync issues
  deterministically.

## Immediate Next Steps

1. Use `syntax_corpus/managed_backlog.md` and
   `docs/type_analysis_corpus_coverage.md` to choose the next verified
   corpus additions for the project-wide binder.
2. Start the project-wide binder vertical slice for public procedures across
   standard modules.
3. Promote small `CANARY_*` cases through observe-only oracle fixtures when
   they become relevant to analyzer behavior.
4. Keep the VBA test runner and lint-suppression directives as planned
   workstreams until their specs and fixture coverage are ready.
5. Treat the XLIDE sidebar as the future product shell for lint, test, setup,
   sync, and workbook actions.
6. Track safety, settings, code actions, performance, release polish, and
   support diagnostics as product-maturity gates before a broad release.

## Files To Keep In Sync

- `docs/roadmap_version_2.x.md`
- `docs/type_analysis_corpus_coverage.md`
- `docs/xlide_vba_type_system_roadmap.md`
- `docs/xlide_vba_linting_test_strategy.md`
- `docs/xlide_vba_lint_suppression_comments.md`
- `docs/xlide_vba_com_test_runner.md`
- `docs/xlide_sidebar_panel.md`
- `docs/xlide_development_principles.md`
- `src/analyzer/diagnostics/ruleMetadata.ts`
- `syntax_corpus/README.md`
- `syntax_corpus/corpus_provenance.json`
- `syntax_corpus/diagnostic_influence_audit.json`
- `syntax_corpus/managed_backlog.md`
- `syntax_corpus/oracle/README.md`
- `syntax_corpus/oracle/run_excel_vbe_oracle.py`
- `syntax_corpus/oracle/vbe_oracle_cases.json`
- `tests/corpusProvenance.test.ts`
