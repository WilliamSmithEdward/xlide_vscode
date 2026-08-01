# XLIDE VBA Realtime Analysis Test Strategy

## Purpose

This document defines the high-level test strategy for validating realtime Excel VBA syntax analysis in a VS Code extension.

The goal is not only to prove that the parser accepts or rejects snippets. The goal is to prove that XLIDE behaves correctly across three different layers:

1. The standalone VBA analysis engine.
2. The live VS Code diagnostic surface.
3. The real Excel/VBE compiler behavior, when available.

These layers should not be treated as equal. The pure analysis engine should be the primary regression target. VS Code and Excel COM should be integration and oracle layers.

---

## Core Principle

Do not make live VS Code or Excel COM the primary test layer.

Use them as validation layers around a deterministic, fixture-driven analyzer test suite.

```text
Spec corpus
   |
Fixture runner
   |
Pure analysis facade tests
   |
VS Code diagnostic integration tests
   |
Optional Excel COM / VBE oracle tests
```

The core analyzer must be testable without launching VS Code, Excel, or COM.

---

## Corpus Accuracy Policy

The syntax corpus is a development asset, not an infallible specification.
Existing corpus cases may be incomplete, stale, or based on assumptions made
before the Excel/VBE oracle existed.

For the broader type-analysis backlog, use
`docs/type_analysis_corpus_coverage.md` as the planning matrix. It can list
pending and missing areas, but those entries are not diagnostic authority until
promoted through this provenance workflow.

When a corpus case affects analyzer behavior, diagnostic severity, or VBE
compile-equivalence metadata, treat it as pending until one of these is true:

- it is traced to an explicit MS-VBAL source rule;
- it is verified with the Excel/VBE oracle;
- it is intentionally marked `observe` or pending and does not drive a hard
  diagnostic.

If the corpus and oracle disagree, the corpus case is suspect. Update the corpus
or mark the discrepancy before using it to justify analyzer behavior.

This does not mean running the oracle for every change. It means using the oracle
for new VBE-behavior evaluation, debugging, discovery, and corpus coverage work.

---

## Current XLIDE Shape

XLIDE already has the right separation for this strategy. Do not rewrite it into
one large parser/analyzer module.

Current pure pieces:

- `src/vbaStructuralDiagnostics.ts` owns fast structural block-balance diagnostics;
  `src/vbaSmartEnter.ts` owns the smart-enter helpers.
- `src/analyzer/diagnostics/analyzeModule.ts` owns high-confidence semantic and
  syntax-adjacent diagnostics over one module.
- `src/analyzer/diagnostics/ruleMetadata.ts` owns the diagnostic rule catalogue.
- `src/vbaWorkbookAnalysis.ts` flattens module diagnostics across a workbook for the
  command and agent-tool path.
- `src/vbaLanguageProviders.ts` adapts pure diagnostics to VS Code ranges,
  severities, debounce, and `DiagnosticCollection`.

The next test layer should add a small pure facade, not collapse these modules:

```ts
export function analyzeVbaModule(
  source: string,
  options: VbaAnalysisOptions
): NormalizedVbaDiagnostic[];
```

That facade should merge `analyzeVbaStructure(source)` and `analyzeModule(source,
options)`, normalize locations, apply mode-specific filtering, and keep the VS
Code adapter thin.

---

## Layer 1: Pure Analyzer Engine Tests

### Purpose

Pure analyzer tests answer this question:

```text
Did the language engine classify this VBA source correctly?
```

This is the main regression suite.

It should be fast, deterministic, and runnable on every platform supported by the extension development workflow.

### What This Layer Tests

This layer should test:

- tokenization
- parsing
- tolerant realtime parsing
- strict background validation
- declaration rules
- semantic checks
- module-kind rules
- Excel-host-aware warnings
- diagnostic codes
- diagnostic severity
- diagnostic ranges
- recovery behavior for incomplete code

### Example Engine Shape

This is a facade over the existing pure modules, not a replacement for them.

```ts
export type AnalysisMode = "realtime" | "strict";

export interface VbaAnalysisOptions {
  mode: AnalysisMode;
  moduleKind: "standard" | "class" | "form" | "worksheet" | "workbook";
  host: "excel";
  settings?: {
    optionExplicit?: "off" | "information" | "warning" | "error";
  };
  knownProcedures?: ReadonlySet<string>;
}

export interface NormalizedVbaDiagnostic {
  code: string;
  severity: "error" | "warning" | "information";
  category: DiagnosticCategory;
  vbeCompileEquivalent: boolean;
  message: string;
  range: {
    startLine: number;
    startCol: number;
    endLine: number;
    endCol: number;
  };
  source: "xlide-vba";
}

export function analyzeVbaModule(
  source: string,
  options: VbaAnalysisOptions
): NormalizedVbaDiagnostic[] {
  return [];
}
```

The existing `analyzeVbaStructure` name should remain reserved for the current
structural block-balance function unless that API is intentionally migrated.

### Why This Layer Comes First

If the analyzer cannot be tested outside VS Code, then every small grammar change becomes slow and fragile.

The extension should adapt the analyzer to VS Code. The VS Code extension should not be the analyzer itself.

---

## Layer 2: Live VS Code Diagnostic Tests

### Purpose

VS Code tests answer this question:

```text
Does the editor surface the right diagnostic in the right place?
```

This layer verifies extension behavior, not VBA correctness.

Keep this layer intentionally small at first. It should be a smoke/integration
suite that proves the VS Code adapter is wired correctly; the full grammar and
semantic matrix belongs in Layer 1.

### What This Layer Tests

This layer should test:

- document activation
- language mode registration
- diagnostic publishing
- diagnostic clearing after edits
- debounce behavior
- range mapping
- severity mapping
- diagnostic code preservation
- command-triggered analysis
- behavior while the user is still typing

### Recommended Test Command

Expose an internal or development-only command for deterministic tests:

```text
xlide.analyzeDocument
```

This command should force analysis for one document URI.

Avoid making tests wait on debounce timers whenever possible.

### Example Flow

```text
1. Open an in-memory VBA document in VS Code.
2. Run xlide.analyzeDocument.
3. Read vscode.languages.getDiagnostics(uri).
4. Normalize the diagnostics.
5. Compare against the fixture expectation.
```

### What Not To Test Here

Do not use VS Code tests to prove the full grammar.

That belongs in the pure analyzer tests.

VS Code tests should prove that the analyzer result is correctly translated into the editor experience.

---

## Layer 3: Excel COM / VBE Oracle Tests

### Purpose

Excel COM / VBE oracle tests answer this question:

```text
Would real Excel/VBE accept or reject this code?
```

This is a compatibility audit layer.

It should be optional, Windows-only, and not required for normal CI.

Important product boundary: XLIDE should remain no-COM/no-Office at runtime.
Oracle tests are a developer validation tool only; they must never become a
runtime dependency or a prerequisite for normal local development.

### Why It Should Be Optional

Excel COM tests are useful, but they are also:

- slow
- machine-dependent
- Windows-only
- dependent on Excel being installed
- dependent on Trust Center settings
- sensitive to Office version differences
- awkward to run in cloud CI
- not a clean compiler API

Use them to audit compatibility, not to drive every local test run.

### Suggested Command

```text
npm run test:oracle:vbe
```

This should be separate from the normal test command.

Initial implementation:

- `syntax_corpus/oracle/vbe_oracle_cases.json` stores empirical fixtures.
- `syntax_corpus/oracle/run_excel_vbe_oracle.py` coordinates per-case timeouts.
- `syntax_corpus/oracle/excel_vbe_oracle_worker.ps1` owns the Excel COM calls.
- `expected: "observe"` records behavior without asserting it yet.
- The runner is observational by default; pass `--strict` to fail on expectation
  mismatches once a local Excel/VBE automation path is stable.

### Oracle Flow

```text
1. Generate a temporary .bas, .cls, or .frm file from a fixture.
2. Launch Excel through COM.
3. Create a temporary .xlsm workbook.
4. Import the module into the VBProject.
5. Attempt to compile, run, or otherwise trigger validation.
6. Capture whether Excel/VBE accepts or rejects the code.
7. Store the oracle result as metadata.
8. Compare XLIDE's syntax classification against real VBE behavior.
```

### Important Constraint

Excel/VBE is the oracle for compile validity.

It is not the oracle for every XLIDE diagnostic.

For example, this may compile:

```vba
Range("A1").Value = 123
```

XLIDE may still warn:

```text
VBA_EXCEL_UNQUALIFIED_RANGE
```

That is a host-safety warning, not a syntax error.

---

## Diagnostic Categories

Every diagnostic should declare what kind of rule it belongs to.

Suggested categories:

```text
syntax
lexer
parser
realtime-recovery
declaration
semantic
project-symbol
module-kind
excel-host
style
```

This matters because not every diagnostic maps to a VBE compile failure.

In XLIDE, these fields should live in
`src/analyzer/diagnostics/ruleMetadata.ts` alongside the existing rule code,
title, severity, MS-VBAL reference, and confidence:

```ts
export type DiagnosticCategory =
  | "syntax"
  | "lexer"
  | "parser"
  | "realtime-recovery"
  | "declaration"
  | "semantic"
  | "project-symbol"
  | "module-kind"
  | "excel-host"
  | "style";

export interface DiagnosticRuleMetadata {
  code: string;
  title: string;
  defaultSeverity: DiagnosticSeverity;
  category: DiagnosticCategory;
  vbeCompileEquivalent: boolean;
  source: "XLIDE";
  specReference?: string;
  requiresWholeProject?: boolean;
  confidence: "high" | "medium" | "low";
}
```

---

## VBE Equivalence Flag

Each diagnostic should explicitly state whether it is expected to match VBE compile behavior.

Example syntax diagnostic:

```json
{
  "code": "VBA_PARSE_UNTERMINATED_STRING",
  "severity": "error",
  "category": "syntax",
  "vbeCompileEquivalent": true
}
```

Example XLIDE opinion warning:

```json
{
  "code": "VBA_EXCEL_UNQUALIFIED_RANGE",
  "severity": "warning",
  "category": "excel-host",
  "vbeCompileEquivalent": false
}
```

This prevents false failures in the oracle layer.

## Severity Policy

Default severity should reflect whether the diagnostic blocks real VBA
compilation or is XLIDE guidance:

- **Error / red squiggly**: deterministic VBE compile-equivalent failures, or
  constructs XLIDE can prove are invalid from explicit source/metadata.
- **Warning / yellow squiggly**: XLIDE-only guidance, maintainability advice,
  suspicious patterns, or soft runtime risks that may still compile.
- **No diagnostic**: uncertain, incomplete while typing, host/object behavior
  XLIDE cannot prove, or anything that would require heuristic guessing.

In practice, `vbeCompileEquivalent: true` diagnostics generally default to
`error`. Diagnostics with `vbeCompileEquivalent: false` should generally default
to `warning` or lower unless XLIDE has deterministic proof that the construct is
invalid under its own non-VBE rule model.

## Diagnostic Language Policy

Diagnostic wording must be as deterministic as the rule that produced it.

- **Error / red squiggly**: use authoritative language. Say the construct "is
  invalid", "will fail", or "will raise" when XLIDE has compile-oracle,
  runtime-oracle, spec, or deterministic local proof.
- **Warning / yellow squiggly**: use advisory language. Say "may", "can",
  "risk", "consider", or similar wording because the diagnostic represents
  guidance, maintainability advice, or a soft risk.
- **No diagnostic**: do not use message text to hedge around uncertainty. If
  XLIDE cannot make the statement authoritatively enough for red or usefully
  enough for yellow, emit nothing.

Tests for red deterministic runtime diagnostics should pin authoritative
phrasing when the exact behavior matters, for example:

```text
This will raise Run-time error '13': Type mismatch.
```

---

## Fixture Format

Use JSON or JSONC fixtures for executable tests.

Markdown is useful for documentation, but JSON/JSONC is better for automated regression tests.

### Single-File Fixture Example

```json
{
  "id": "line-continuation-trailing-comment",
  "title": "Line continuation cannot have trailing comment",
  "moduleKind": "standard",
  "host": "excel",
  "mode": "strict",
  "settings": {
    "optionExplicit": "off"
  },
  "code": "Public Sub Test()\n    Debug.Print \"hello\" & _ <| ' bad trailing comment |>\nEnd Sub",
  "expected": [
    {
      "code": "VBA_LINE_CONTINUATION_TRAILING_TOKENS",
      "severity": "error",
      "rangeMarker": 0
    }
  ]
}
```

The marker syntax is only for tests:

```text
<| diagnostic target |>
```

The test harness should strip markers before analysis, then use those marker ranges for expected diagnostic locations.

### Fixture Defaults

Fixtures should be explicit about diagnostics that are intentionally enabled.
For example, if a fixture is testing one syntax error and omits `Option
Explicit`, either add `Option Explicit` to the code or set:

```json
{
  "settings": {
    "optionExplicit": "off"
  }
}
```

This avoids noisy fixture failures where a style/configurable warning masks the
rule under test.

If a fixture intentionally allows extra non-error diagnostics, say so directly:

```json
{
  "allowExtraDiagnostics": ["option-explicit-missing"]
}
```

The default should be strict comparison on stable fields: code, severity, range,
category, and mode behavior.

---

## Realtime vs Strict Expectations

Some source should behave differently depending on mode.

Example:

```json
{
  "id": "incomplete-thisworkbook-member-access",
  "title": "Incomplete member access should not hard-error in realtime mode",
  "moduleKind": "standard",
  "host": "excel",
  "code": "Public Sub Test()\n    ThisWorkbook.\nEnd Sub",
  "expectations": {
    "realtime": [],
    "strict": [
      {
        "code": "VBA_INCOMPLETE_MEMBER_ACCESS",
        "severity": "error"
      }
    ]
  }
}
```

This distinction is essential.

Realtime analysis should preserve typing flow.

Strict validation should preserve correctness.

In the current codebase, mode should initially be implemented in the pure facade
by filtering or downgrading diagnostics from `analyzeVbaStructure` and `analyzeModule`.
Only move mode awareness deeper into the parser/analyzer when a rule genuinely
needs different parse recovery behavior.

---

## Project-Level Fixtures

VBA is not purely file-local.

The analyzer should eventually support project-level fixtures.

Example:

```json
{
  "id": "duplicate-public-procedure-across-modules",
  "project": {
    "modules": [
      {
        "name": "Module1",
        "kind": "standard",
        "code": "Public Sub RefreshData()\nEnd Sub"
      },
      {
        "name": "Module2",
        "kind": "standard",
        "code": "Public Sub RefreshData()\nEnd Sub"
      }
    ]
  },
  "expected": [
    {
      "code": "VBA_DECL_AMBIGUOUS_PUBLIC_PROCEDURE",
      "severity": "warning"
    }
  ]
}
```

Project-level tests are required for:

- duplicate public names
- ambiguous references
- module-kind-sensitive events
- class module behavior
- workbook and worksheet module behavior
- project references
- cross-module symbol lookup

The first implementation can be modest: build a project fixture adapter that
feeds module names, module kinds, and a project-wide procedure set into the pure
facade. Deeper reference binding can be added later without changing the fixture
shape.

---

## Normalized Diagnostic Comparison

Do not compare raw VS Code diagnostics directly.

Normalize first.

Compare only stable fields:

```text
code
severity
range
category
mode behavior
```

Avoid strict comparison on full message text at first.

Messages are user-facing and may change.

Diagnostic codes should be stable.

---

## Recommended Test Scripts

```json
{
  "scripts": {
    "test": "vitest run",
    "test:core": "vitest run",
    "test:fixtures": "vitest run tests/vbaAnalysisFixtures.test.ts",
    "test:vscode": "vscode-test",
    "test:oracle:vbe": "node syntax_corpus/oracle/run_excel_vbe_oracle.mjs",
    "check-types": "tsc --noEmit",
    "compile": "npm run check-types && node esbuild.js"
  }
}
```

The normal test command should not require Excel. It also does not need to run
the VS Code integration suite until that suite is stable and fast enough for the
local loop.

---

## Suggested CI Matrix

```text
Pull request CI:
  - TypeScript compile
  - pure analyzer fixture tests
  - selected VS Code integration tests, once stable

Nightly or manual Windows CI:
  - pure analyzer fixture tests
  - VS Code integration tests
  - optional Excel COM / VBE oracle tests

Local developer command:
  - npm test

Local oracle command:
  - npm run test:oracle:vbe
```

---

## Build Order

Recommended implementation order:

```text
1. Add category and VBE-equivalence fields to the diagnostic rule registry.
2. Define the pure `analyzeVbaModule` facade and normalized diagnostic shape.
3. Define fixture format.
4. Build marker stripper.
5. Build pure analyzer fixture runner.
6. Add 20 syntax fixtures.
7. Add realtime-vs-strict expectations.
8. Wire facade output into VS Code diagnostics.
9. Add VS Code integration smoke tests.
10. Add project-level fixture support.
11. Add optional Excel COM / VBE oracle runner.
```

Do not start with Excel automation.

That should come after the analyzer is deterministic.

---

## Confidence Model

A diagnostic has the strongest confidence when all three layers agree.

```text
XLIDE analyzer:     invalid
VS Code surface:  red squiggle in correct range
Excel/VBE oracle: compile failure
Result:           high confidence syntax error
```

A warning can still be valid even when Excel compiles the code.

```text
XLIDE analyzer:     warning
VS Code surface:  yellow squiggle
Excel/VBE oracle: compiles fine
Result:           valid XLIDE opinion warning
```

A realtime recovery case may intentionally differ from strict behavior.

```text
Realtime mode:    no hard error
Strict mode:      error
Result:           correct editor-friendly behavior
```

---

## Do Not Overfit To VBE

VBE is the compatibility oracle for actual VBA compile behavior.

It is not the design oracle for the whole extension.

XLIDE can and should provide diagnostics that VBE does not provide, including:

- unqualified Excel object warnings
- implicit Variant warnings
- risky default member warnings
- style warnings
- dead code informational findings
- incomplete realtime state handling
- project hygiene warnings

The analyzer should distinguish between:

```text
This is invalid VBA.
```

and:

```text
This is valid VBA, but risky Excel code.
```

That distinction is what will make the extension feel accurate instead of noisy.

---

## Definition Of Done

The analysis test strategy is working when:

```text
1. Every diagnostic has a stable code.
2. Every diagnostic declares category and VBE compile equivalence.
3. A pure facade returns normalized diagnostics without VS Code, Excel, or COM.
4. Every fixture has an expected severity and range.
5. Realtime and strict modes can disagree intentionally.
6. VS Code tests prove diagnostics surface correctly.
7. Excel COM oracle tests are optional, not mandatory.
8. Project-level fixtures can model multiple modules.
9. Excel-host warnings are not confused with syntax errors.
10. New parser bugs can be reproduced by adding one fixture.
```

At that point, the analyzer is not just a demo. It is a regression-tested language subsystem.
