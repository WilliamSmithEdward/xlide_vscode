# XLIDE VBA Realtime Linting Test Strategy

## Purpose

This document defines the high-level test strategy for validating realtime Excel VBA syntax linting in a VS Code extension.

The goal is not only to prove that the parser accepts or rejects snippets. The goal is to prove that XLIDE behaves correctly across three different layers:

1. The standalone VBA linting engine.
2. The live VS Code diagnostic surface.
3. The real Excel/VBE compiler behavior, when available.

These layers should not be treated as equal. The pure linting engine should be the primary regression target. VS Code and Excel COM should be integration and oracle layers.

---

## Core Principle

Do not make live VS Code or Excel COM the primary test layer.

Use them as validation layers around a deterministic, fixture-driven linter test suite.

```text
Spec corpus
   ↓
Pure linter engine tests
   ↓
VS Code diagnostic integration tests
   ↓
Optional Excel COM / VBE oracle tests
```

The core linter must be testable without launching VS Code, Excel, or COM.

---

## Layer 1: Pure Linter Engine Tests

### Purpose

Pure linter tests answer this question:

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

```ts
export type LintMode = "realtime" | "strict";

export interface VbaLintOptions {
  mode: LintMode;
  moduleKind: "standard" | "class" | "form" | "worksheet" | "workbook";
  host: "excel";
  optionExplicit?: boolean;
}

export interface VbaDiagnostic {
  code: string;
  severity: "error" | "warning" | "info";
  message: string;
  range: {
    startLine: number;
    startCol: number;
    endLine: number;
    endCol: number;
  };
  source: "xlide-vba";
}

export function lintVbaSource(
  source: string,
  options: VbaLintOptions
): VbaDiagnostic[] {
  return [];
}
```

### Why This Layer Comes First

If the linter cannot be tested outside VS Code, then every small grammar change becomes slow and fragile.

The extension should adapt the linter to VS Code. The VS Code extension should not be the linter itself.

---

## Layer 2: Live VS Code Diagnostic Tests

### Purpose

VS Code tests answer this question:

```text
Does the editor surface the right diagnostic in the right place?
```

This layer verifies extension behavior, not VBA correctness.

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
- command-triggered linting
- behavior while the user is still typing

### Recommended Test Command

Expose an internal command for deterministic tests:

```text
xlide.lintDocument
```

This command should force linting for one document URI.

Avoid making tests wait on debounce timers whenever possible.

### Example Flow

```text
1. Open an in-memory VBA document in VS Code.
2. Run xlide.lintDocument.
3. Read vscode.languages.getDiagnostics(uri).
4. Normalize the diagnostics.
5. Compare against the fixture expectation.
```

### What Not To Test Here

Do not use VS Code tests to prove the full grammar.

That belongs in the pure linter tests.

VS Code tests should prove that the linter result is correctly translated into the editor experience.

---

## Layer 3: Excel COM / VBE Oracle Tests

### Purpose

Excel COM / VBE oracle tests answer this question:

```text
Would real Excel/VBE accept or reject this code?
```

This is a compatibility audit layer.

It should be optional, Windows-only, and not required for normal CI.

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

The test harness should strip markers before linting, then use those marker ranges for expected diagnostic locations.

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

Realtime linting should preserve typing flow.

Strict validation should preserve correctness.

---

## Project-Level Fixtures

VBA is not purely file-local.

The linter should eventually support project-level fixtures.

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
    "test": "npm run test:core && npm run test:vscode",
    "test:core": "mocha \"out/test/**/*.test.js\"",
    "test:vscode": "vscode-test",
    "test:oracle:vbe": "node ./scripts/run-vbe-oracle.js",
    "compile": "tsc -p ./"
  }
}
```

The normal test command should not require Excel.

---

## Suggested CI Matrix

```text
Pull request CI:
  - TypeScript compile
  - pure linter fixture tests
  - selected VS Code integration tests

Nightly or manual Windows CI:
  - pure linter fixture tests
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
1. Define fixture format.
2. Build marker stripper.
3. Build pure linter test runner.
4. Add 20 syntax fixtures.
5. Add realtime-vs-strict expectations.
6. Add diagnostic code registry.
7. Wire linter result into VS Code diagnostics.
8. Add VS Code integration tests.
9. Add project-level fixture support.
10. Add optional Excel COM / VBE oracle runner.
```

Do not start with Excel automation.

That should come after the linter is deterministic.

---

## Confidence Model

A diagnostic has the strongest confidence when all three layers agree.

```text
XLIDE linter:     invalid
VS Code surface:  red squiggle in correct range
Excel/VBE oracle: compile failure
Result:           high confidence syntax error
```

A warning can still be valid even when Excel compiles the code.

```text
XLIDE linter:     warning
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
- dead code hints
- incomplete realtime state handling
- project hygiene warnings

The linter should distinguish between:

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

The linting test strategy is working when:

```text
1. Every diagnostic has a stable code.
2. Every fixture has an expected severity and range.
3. Realtime and strict modes can disagree intentionally.
4. VS Code tests prove diagnostics surface correctly.
5. Excel COM oracle tests are optional, not mandatory.
6. Project-level fixtures can model multiple modules.
7. Excel-host warnings are not confused with syntax errors.
8. New parser bugs can be reproduced by adding one fixture.
```

At that point, the linter is not just a demo. It is a regression-tested language subsystem.
