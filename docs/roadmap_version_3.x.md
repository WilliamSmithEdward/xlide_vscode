# XLIDE Roadmap Version 3.x

Forward backlog for work intentionally moved out of Version 2.x closeout.
Version 2.x remains the launch-hardening track; Version 3.x is for deeper
language-modeling, object-metadata, and cross-workbook workflows that should not
block the current v2 completion push.

## North Star

V3 should extend XLIDE's deterministic VBA understanding without weakening the
core rule: no hard diagnostics from guesses. Every new binder, metadata, object,
or workflow surface must be backed by source facts, generated metadata with
provenance, the Microsoft specification, or focused VBE/Excel oracle evidence.

## Workstream A: Binder Shadowing and Expression Binding

Purpose: refine the broad v2 backlog item into implementable binder slices.

- [ ] Define the exact bare-identifier precedence ladder for expression
  contexts:
  - procedure locals and parameters
  - procedure labels where syntactically relevant
  - procedure return variables
  - same-module private/public declarations
  - exported standard-module members
  - enum members and UDT/type names where expression use is valid
  - class/document/UserForm code names
  - runtime, host, and external reference symbols
- [ ] Model shadowing separately for declaration contexts, call contexts,
  assignment targets, member receivers, type positions, and `New` positions.
- [ ] Preserve no-diagnostic behavior for ambiguous exported names and unknown
  external references until the resolver can prove the target.
- [ ] Add external-reference constants and globals as an explicit metadata-backed
  identifier source, with provenance and ambiguity behavior.
- [ ] Add deterministic arbitrary-expression binding only for proven expression
  shapes:
  - parenthesized expressions
  - known function/property return values
  - known member chains
  - collection/default-member calls where metadata proves the returned type
  - `With` receivers and nested `With` receiver stacks
- [ ] Add workbook fixture scenarios covering shadowing, duplicates,
  multi-workbook isolation, live-source overlays, and unresolved external names.

Definition of done:

- The binder can explain why a name resolves, why it is ambiguous, or why XLIDE
  intentionally stays silent.

## Workstream B: Object Member Model

Purpose: broaden object-member authoring and diagnostics beyond the v2 shipped
source-backed and generated Excel slices.

- [ ] Add designer-backed UserForm metadata ingestion sufficient to prove form
  controls, control types, and control event surfaces.
- [ ] Offer UserForm form/control event handlers only when designer-backed
  metadata proves the control/event surface.
- [ ] Model declared `Event` members with declarations, signatures, visibility,
  and source spans.
- [ ] Model `WithEvents` bindings and their event handler authoring surface.
- [ ] Verify default-member/direct object-expression behavior, including
  `VB_UserMemId = 0`, before inferring direct object usage such as
  `textValue = p`.
- [ ] Extend source member signatures with richer parameter metadata,
  declaration spans, return/write types, docs, and provenance.
- [ ] Add document/UserForm designer-backed members to the same object-member
  contract used by source-backed classes and host metadata.

Definition of done:

- UserForm/document member and event surfaces are deterministic and do not invent
  controls, events, or members from names alone.

## Workstream C: External Object/Member Metadata

Purpose: let downstream developers describe referenced libraries, add-ins, and
host extensions that XLIDE cannot parse from workbook source.

- [ ] Define a versioned external object/member metadata schema.
- [ ] Support member names, kinds, signatures, parameter docs/types, return
  types, examples, mutability, exhaustiveness, and provenance.
- [ ] Define reload behavior and validation diagnostics for malformed metadata.
- [ ] Define deterministic precedence:
  - workbook source symbols win for workbook-owned members
  - inline docs enrich source symbols
  - external metadata describes explicitly declared external/extension members
  - curated host/runtime metadata remains the built-in fallback
- [ ] Add completion, hover, signature help, member-call diagnostics, assignment
  diagnostics, and no-diagnostic controls for external metadata.
- [ ] Ship downstream developer documentation with schema examples,
  troubleshooting, provenance rules, and verification steps.

Definition of done:

- A downstream developer can author metadata, reload it, verify `object.`
  completion, and troubleshoot missing members without reading XLIDE source.

## Workstream D: Host Metadata Completeness

Purpose: make generated host metadata broad, auditable, and safe enough to power
hard diagnostics type by type.

- [ ] Normalize the repo-local `reference/` dump corpus into generated metadata
  that production extension code can consume without reading `reference/` at
  runtime.
- [ ] Preserve Office/version/source provenance for every type, member,
  signature, return type, enum, event, default member, and writable/read-only
  fact.
- [ ] Diff generated dumps against curated metadata and official documentation
  where available.
- [ ] Add oracle spot checks for behavior a reference dump cannot answer.
- [ ] Produce coverage reports by host/library type so gaps are visible.
- [ ] Resolve remaining Excel object receiver chains beyond the v2 simple
  return-type and collection-default `Item` paths.
- [ ] Promote generated host types into hard `member-not-found` only
  type-by-type after coverage reports and representative oracle controls prove
  the surface complete enough for red diagnostics.
- [ ] Require every host metadata expansion to add coverage for completion,
  hover/signature docs where applicable, member-call arity/type, assignment
  validation, `member-not-found`, and no-diagnostic controls for incomplete or
  non-exhaustive surfaces.

Definition of done:

- Host metadata has auditable provenance and hard absence diagnostics only come
  from exhaustive receiver surfaces.

## Workstream E: Workbook-To-Workbook Transfer

Purpose: support explicit module transfer between workbooks without crossing
analysis scopes or guessing user intent.

- [ ] Add source workbook and destination workbook selection.
- [ ] Add module/class selection with a side-by-side preview.
- [ ] Add conflict handling for existing destination modules.
- [ ] Add backup/snapshot hooks before destination workbook mutation where
  practical.
- [ ] Preserve multi-workbook analysis isolation; transfer previews must not
  imply cross-workbook project binding.
- [ ] Record write-audit entries and changed/skipped/failed summaries.

Definition of done:

- Workbook-to-workbook transfer is explicit, previewed, auditable, and
  recoverable.

## Files To Keep In Sync

- `docs/roadmap_version_2.x.md`
- `docs/roadmap_version_3.x.md`
- `docs/xlide_external_member_metadata.md`
- `docs/excel_reference_coverage.md`
- `src/analyzer/symbols/projectIndex.ts`
- `src/analyzer/completion/memberAccess.ts`
- `src/analyzer/host/excelReferenceMembers.ts`
- `src/analyzer/host/excelObjectModel.ts`
- `tests/fixtures/vbaProjects/`
