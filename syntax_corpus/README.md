# Syntax Corpus

This folder is a development corpus for XLIDE's VBA diagnostics and realtime
editing behavior. It is evidence, not authority.

Before a corpus case drives a hard analyzer diagnostic, it must be backed by one
of these sources:

- MS-VBAL or another Microsoft primary source.
- The Excel/VBE oracle under `syntax_corpus/oracle`.
- Deterministic project metadata that XLIDE owns and tests.

If that evidence is missing, keep the case as pending or observational and do
not use it to justify a red diagnostic.

## Provenance States

- `spec-derived`: verified directly against a Microsoft specification or
  documentation source.
- `vbe-oracle-verified`: accepted/rejected behavior has been verified with the
  Excel/VBE oracle and is asserted by the fixture.
- `observed-not-asserted`: behavior has been observed, but the fixture is still
  collecting evidence and should not fail the suite.
- `pending-verification`: existing corpus material that has not yet been
  rechecked against a primary source or the oracle.

## Metadata Convention

Every active corpus-bearing file must be listed in `corpus_provenance.json`.

Each entry also carries a `status`:

- `mining`: a planning corpus still being mined for promotable cases.
- `reference`: ongoing reference material, not a promotion queue (recommendations,
  limits, the audit, the backlog, the oracle fixtures).
- `reconciled`: fully dispositioned — every case has a recorded verdict.

When a file becomes `reconciled` it is moved to `syntax_corpus/Archive/` and its
record is relocated from `files` to the `archivedFiles` array. This keeps the
active `files` list (which the provenance test validates against the top-level
directory) clean while preserving the audit record. The corpus is a living index:
some veins (e.g. binder-independent compile errors) mine out and get archived;
others (host behavior, completion, realtime) stay as permanent reference.

Machine-readable fixture files should also add a per-case `provenance` field.
The per-case value is the source of truth when it exists. File-level provenance
is only a coarse audit marker for Markdown planning corpora and legacy fixture
collections.

`diagnostic_influence_audit.json` maps every active diagnostic code to its
current evidence source. It distinguishes diagnostics backed by asserted oracle
cases from diagnostics that are spec-derived and diagnostics that only mention
observe-only oracle cases for future discovery.

`managed_backlog.md` digests the Markdown corpus into backlog categories such as
syntax, realtime recovery, type analysis, runtime resolution, project binding,
host behavior, completion context, UserForm/designer symbols, limits, and
legacy edges. It is a planning index, not diagnostic-driving evidence.

When promoting a Markdown example into an executable fixture, copy the case into
a machine-readable fixture file and assign one of the provenance states above.
Hard error expectations require `spec-derived` or `vbe-oracle-verified`
evidence.
