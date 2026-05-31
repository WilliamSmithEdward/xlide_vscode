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

Every corpus-bearing file must be listed in `corpus_provenance.json`.

Machine-readable fixture files should also add a per-case `provenance` field.
The per-case value is the source of truth when it exists. File-level provenance
is only a coarse audit marker for Markdown planning corpora and legacy fixture
collections.

When promoting a Markdown example into an executable fixture, copy the case into
a machine-readable fixture file and assign one of the provenance states above.
Hard error expectations require `spec-derived` or `vbe-oracle-verified`
evidence.
