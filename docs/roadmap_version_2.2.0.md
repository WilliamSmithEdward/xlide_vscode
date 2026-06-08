# XLIDE Roadmap Version 2.2.0

Forward backlog moved out of Version 2.1.0 so that 2.1.0 can stay focused on
red-squiggly completeness, binding/name resolution, host metadata hardening, and
source-backed object/event authoring.

Version 2.2.0 begins with the lower-priority developer-experience tracks that
were previously listed as Version 2.1.0 Priority 4 and Priority 5. These items
remain important, but they should follow the core language-service hardening
unless customer demand changes.

## Priority 1: External Metadata Authoring and Reload

Purpose: let downstream developers describe referenced libraries, add-ins, and
host extensions that XLIDE cannot parse from workbook source.

Developer-experience impact:

- Lets teams add proven API surfaces that can safely power hard diagnostics for
  private or third-party dependencies when the metadata is exhaustive enough.
- Gives teams a practical path to improve completion, hover, signature help, and
  diagnostics for private or third-party dependencies.
- Lets advanced users close gaps without waiting for a full XLIDE release.
- Makes missing or malformed metadata visible through validation and reload
  feedback instead of silent failure.

Scope:

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

## Priority 2: Workbook-To-Workbook Transfer

Purpose: support explicit module transfer between workbooks without crossing
analysis scopes or guessing user intent.

Developer-experience impact:

- Improves a useful project-maintenance workflow, especially when moving modules
  between real workbooks.
- Keeps workbook mutation explicit, previewed, auditable, and recoverable.
- Has lower red-squiggle impact than binding and metadata, so it should follow
  the core authoring improvements unless customer demand changes.

Scope:

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
- `docs/roadmap_version_2.1.0.md`
- `docs/roadmap_version_2.2.0.md`
- `docs/xlide_external_member_metadata.md`
- `docs/xlide_development_principles.md`
- `docs/xlide_performance_budgets.md`
- `docs/excel_reference_coverage.md`
- `src/vbaLanguageProviders.ts`
- `src/vbaModuleAnalysis.ts`
- `src/vbaProjectAnalysis.ts`
- `src/vbaWorkbookAnalysis.ts`
- `tests/fixtures/vbaProjects/`
- `syntax_corpus/README.md`
- `syntax_corpus/managed_backlog.md`
- `syntax_corpus/corpus_provenance.json`
- `syntax_corpus/diagnostic_influence_audit.json`
