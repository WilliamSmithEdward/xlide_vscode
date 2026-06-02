# XLIDE Development Principles

## Deterministic Logic Only

All code and analyzer logic must be deterministic. The same inputs must produce
the same outputs for tokens, parse trees, symbols, completions, diagnostics,
workbook operations, and generated metadata.

Do not implement heuristic, fuzzy, probability-based, or "looks like" behavior
unless the operator explicitly and unwaveringly agrees before implementation.

When behavior is uncertain:

- Prefer no diagnostic over a guessed diagnostic.
- Mark unresolved grammar or API behavior as pending instead of inventing it.
- Use explicit parsers, curated metadata, or documented rules.
- Keep provenance clear for native/runtime/host metadata.
- Treat ambiguous resolution as unknown, not as a best guess.

## Shoot For The Moon

For complicated work, create or update a roadmap Markdown file first. The
roadmap should name the intended complete system, phases, risks, and definition
of done.

Implementation should then proceed iteratively through production-quality
vertical slices. Each slice should fit the larger roadmap and avoid throwaway
MVP shortcuts or half measures.

## Compatibility Patches

Backwards compatibility patches, aliases, legacy paths, migration shims, and
parallel old/new code paths are anti-patterns by default.

Use them only when there is a very good, well-defined reason and the operator
explicitly agrees to the reason, scope, and retirement plan.

When working in the codebase, actively look for unnecessary compatibility hacks
and remove them instead of preserving them. Prefer one supported path, one set
of names, and one business rule implementation.

## Diagnostic Language

Diagnostic text must match the certainty of the rule.

- Red/error diagnostics use authoritative language: "is invalid", "will fail",
  or "will raise" when the behavior is proven.
- Yellow/warning diagnostics use advisory language: "may", "can", "risk",
  "consider", or similar wording for guidance and soft risks.
- Uncertain behavior should not be hidden behind hedged red text. Prefer no
  diagnostic until the behavior is deterministic enough to state plainly.

## Configuration Scoping

Use discernment when deciding where settings live.

- Global extension settings are defaults and environment/user preferences.
- Workbook-scoped settings are choices that are meaningful for one workbook and
  may differ from another workbook in the same workspace.
- Workbook-facing GUIs should write workbook-scoped settings for workbook
  decisions and should inherit global defaults only when the workbook has no
  explicit value.
- Some workbook settings have no global equivalent. Workbook-specific facts
  such as a selected sync folder should live only with the workbook.
- A workbook GUI action must not silently mutate global defaults.
- Sidecar path resolution, schema validation, and reads/writes should live in
  the shared workbook settings owner, not in individual feature modules.
- Displayed configuration and runtime behavior must come from the same resolver.
  Do not create separate GUI-only, command-only, or compatibility-only settings
  paths for the same business rule.

## Excel/VBE Oracle Usage

The Excel/VBE oracle is a discovery, debugging, and corpus-coverage tool. Use it
when investigating real VBE behavior, validating new corpus cases, or changing
the oracle harness itself.

Do not run the oracle as routine verification for every code change. Prefer fast
local checks such as targeted unit tests, `npm.cmd test`,
`npm.cmd run check-types`, and `npm.cmd run compile` unless the change depends
on fresh Excel/VBE behavior.
