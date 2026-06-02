# XLIDE VBA Type System Roadmap

## Purpose

XLIDE should grow from high-confidence syntax analysis into a real VBA type
analysis layer: one that understands user code, Excel/VBA native APIs, module
kinds, object members, documentation metadata, and the difference between
compile-time errors, runtime type risks, and intentional VBA coercion.

The goal is ambitious: make VS Code feel like a serious VBA IDE without
pretending VBA is a stricter language than it is.

---

## North Star

The type system should eventually answer these questions:

- What type does this expression produce?
- Which callable overload/signature is being invoked?
- Are supplied arguments compatible with the expected parameter types?
- Does this member exist on the receiver type?
- Is `Set` required, forbidden, or optional here?
- Is a value being assigned to something compatible?
- Is this diagnostic a VBE-equivalent compile error, a deterministic runtime
  error, a runtime risk, or an XLIDE quality hint?

Diagnostics must stay conservative. If XLIDE cannot prove a mismatch, it should
prefer no diagnostic over a noisy one.

---

## Sources Of Type Truth

Use these sources in this order:

1. **Parsed VBA source**
   - Procedure parameters and return types.
   - Local/module variables and constants.
   - Enums, UDTs, class modules, document modules, and userforms.
2. **Project symbol graph**
   - Public procedures and types across modules.
   - Class members and Implements relationships.
   - Workbook/document module code names.
3. **Curated native metadata**
   - VBA runtime functions/statements in `src/analyzer/runtime/vbaRuntime.ts`.
   - Excel object model signatures in `src/analyzer/host/excelObjectModel.ts`.
4. **External XML documentation metadata**
   - Team-authored signatures for missing native/project APIs.
   - Type hints for symbols the source parser cannot see.
5. **Inline `'''` documentation comments**
   - Descriptive metadata and optional type/unit/value hints.
   - Never override a concrete source declaration type.

---

## Documentation Type Hints

Inline and external docs may carry optional type metadata:

```vba
''' <summary>Calculates the invoice total after tax.</summary>
''' <param name="Subtotal" type="Currency" unit="money">The pre-tax amount.</param>
''' <param name="TaxRate" type="Double" unit="decimal">Use 0.0825 for 8.25%.</param>
''' <returns type="Currency" unit="money">The subtotal plus calculated tax.</returns>
Public Function InvoiceTotal(ByVal Subtotal As Currency, ByVal TaxRate As Double) As Currency
End Function
```

Rules:

- Source declarations win over doc-comment `type`.
- Doc `type` fills gaps for external metadata or untyped/Variant-heavy APIs.
- `unit` and `value` are descriptive hints for hover/signature help first; they
  should not become hard errors unless a dedicated rule explicitly says so.

---

## Type Compatibility Model

Build a small compatibility lattice first, then widen it carefully:

- `Variant` accepts everything and disables hard mismatch diagnostics.
- Numeric family: `Byte`, `Integer`, `Long`, `LongLong`, `LongPtr`, `Single`,
  `Double`, `Currency`, `Decimal`.
- String family: `String`, fixed-length strings, string literals.
- Boolean family: `Boolean`, `True`, `False`.
- Date family: `Date`, date literals.
- Object family: `Object`, host object types, class types, `Nothing`.
- Enum family: enum members plus compatible integer values.
- Array family: dynamic/fixed arrays and `ParamArray`.

VBA performs many implicit conversions. XLIDE should start by flagging only
high-confidence mismatches, such as a non-numeric string literal passed to a
numeric parameter.

---

## Diagnostic Policy

Suggested rule families:

```text
argument-type-mismatch
assignment-type-mismatch
set-required
set-forbidden
member-not-found
object-required
array-required
byref-argument-type-mismatch
enum-value-mismatch
```

Each diagnostic should declare:

- stable code
- category
- severity
- confidence
- VBE compile equivalence
- whether it is compile-time, deterministic-runtime-error, runtime-risk, or
  XLIDE guidance

---

## Implementation Phases

Coverage planning lives in `docs/type_analysis_corpus_coverage.md`. Treat that
matrix as the backlog for pending, partial, missing, and verified corpus areas;
do not treat pending Markdown corpus examples as authority for hard diagnostics
until the case has spec, oracle, or deterministic XLIDE-owned evidence.

### Phase 1: Literal and Declared-Type Arguments

Status: started.

- Parse parameter types for same-module procedures.
- Parse selected runtime/native signatures from explicit curated metadata only.
- Infer obvious argument types from literals and declared variables.
- Flag only high-confidence argument mismatches.
- Add tests for user procedures and native VBA runtime functions.

Landed first:

- Same-module procedure parameter/return types are collected from parsed source.
- Literal arguments, declared local/module variables, and nested call returns can
  produce simple inferred types.
- Obvious arithmetic expressions composed only of proven numeric operands infer
  as numeric; expressions with `Variant`, unknown, or string operands remain
  unknown.
- String concatenation expressions using `&` infer as `String` only when every
  operand has a known scalar type; `Variant`, unknown, and object-like operands
  keep the expression unknown.
- Nonnumeric string literals in numeric arithmetic expressions produce an error
  when the expected context is numeric; focused oracle verification shows the
  representative assignment compiles but deterministically raises runtime error
  13 Type mismatch when executed.
- Same-module expression calls are checked for required argument count.
- Empty positional slots are rejected when the corresponding parameter is
  required.
- Call statements that resolve to variables, parameters, constants, types, or
  enum members are rejected as non-callable, including bare variable statements
  such as `testStr`; VBE Compile rejects these as `Expected Sub, Function, or
  Property`.
- Provably nonnumeric string literals passed to numeric parameters are flagged as
  deterministic runtime errors after focused Excel/VBE runtime oracle evidence.
- Numeric literals and numeric strings remain accepted for numeric parameters.
- String formats whose VBA coercion depends on runtime value, locale, or deeper
  conversion semantics remain unknown until modeled explicitly.
- `argument-type-mismatch` is error-severity only when XLIDE has deterministic
  proof that the supplied argument will raise a runtime Type mismatch even
  though VBE Compile accepts the code.
- `assignment-type-mismatch` follows the same red runtime-error policy for
  deterministic scalar assignment coercion failures.
- Object argument mismatches use a separate compile-equivalent error diagnostic
  after focused Excel/VBE oracle verification.
- Runtime functions such as `Int` used as `As` type names are compile-equivalent
  errors. Broad unknown type names remain deferred until the project binder can
  see workbook classes, UDTs, enums, and host object types.
- The project binder now exposes visible project type names from
  `ProjectIndex.visibleTypeNames()`: class/document/UserForm module names,
  current-module `Type`/`Enum` declarations, and non-`Private` cross-module
  `Type`/`Enum` declarations. It preserves duplicates so the shared resolver
  reports ambiguity deterministically instead of picking by module order, and
  the live completion provider now consumes this same project binder.
- `Set` assignment to a known intrinsic scalar variable is a compile-equivalent
  error. Unknown object-like target types remain deferred to the object binder.
- Member access on a declared intrinsic scalar receiver is a red diagnostic:
  named members such as `s.Length` are VBE Compile `Invalid qualifier` errors,
  while trailing dots such as `s.` are VBE Compile `Syntax error`s. Unknown,
  `Variant`, class, object, and UDT receivers stay silent.
- Named arguments map to the named parameter before type validation.
- Curated VBA runtime functions participate when their parameter type is known.
- Runtime parameter types are never inferred from parameter names.

### Phase 2: Expression Return Types

- Infer return types for nested calls, arithmetic, concatenation, comparisons,
  and conversion functions.
- Track function results through assignments.
- Respect `Variant` and unknowns as "do not warn".

### Phase 3: Project-Wide Binder

- Resolve public procedures, classes, enums, UDTs, and module members across the
  workbook project.
- Landed first slice: unique exported standard-module `Sub`/`Function`
  signatures feed cross-module argument count and type diagnostics; ambiguous
  duplicate exported bare names are skipped, while module-qualified
  `ModuleName.ProcedureName` calls resolve through the named standard module.
- Landed visibility slice: bare cross-module procedure names are filtered to
  same-module procedures plus exported standard-module procedures, so private
  procedures in other modules and object-module members do not hide
  `unknown-call`.
- Landed type-name binder groundwork: `ProjectIndex.visibleTypeNames()` exposes
  visible project-defined type names to the shared `As`/`New` type resolver.
- Support document modules, userforms, and class modules.
- Model procedure visibility and shadowing.

### Phase 4: Object and Member Types

- Resolve receiver type chains such as `ThisWorkbook.Worksheets(1).Range("A1")`.
- Resolve workbook-defined class/document/UserForm members for known receiver
  types such as `Dim p As Person: p.`.
- Validate Excel object model method/property calls.
- Add `Set` diagnostics and member-not-found diagnostics where receiver type is
  known.

### Phase 5: Native Metadata Expansion

- Broaden curated VBA runtime signatures.
- Broaden Excel object model signatures.
- Support external `.vbref.xml` metadata as a first-class type-signature and
  object/member completion source for explicitly declared external APIs.
- Keep metadata provenance auditable.
- Ship downstream developer documentation and examples for authoring metadata
  before enabling the workflow by default.

### Phase 6: Realtime Typing Experience

- Suppress hard type errors for incomplete expressions.
- Keep signature help, completions, and diagnostics using the same binder.
- Add code actions where fixes are obvious, such as adding call parentheses.

---

## First Vertical Slice

The first useful slice should catch:

```vba
Public Function InvoiceTotal(ByVal Subtotal As Currency, ByVal TaxRate As Double) As Currency
End Function

Public Sub TestInvoiceTotal()
    total = InvoiceTotal("blah", 0.08)
End Sub
```

Expected diagnostic:

```text
argument-type-mismatch: expected Currency for Subtotal, got String literal that cannot be converted to a number
```

It should not flag:

```vba
total = InvoiceTotal("100", 0.08)
total = InvoiceTotal(100, 0.08)
```

because VBA can coerce numeric strings and numeric literals to `Currency`.

---

## Definition Of Done

The type system is mature when:

1. User procedure calls are type-checked across modules.
2. Native VBA runtime calls are metadata-backed and tested.
3. Excel object model calls are validated when receiver types are known.
4. `Variant` and unknown types suppress noisy false positives.
5. Type hints in comments enrich hover/signature help and metadata gaps.
6. Diagnostics distinguish compile errors, runtime risks, and XLIDE guidance.
7. Every type rule has fixture coverage for valid, invalid, and unknown cases.
