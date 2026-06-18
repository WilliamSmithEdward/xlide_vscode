# XLIDE Roadmap Version 2.6.0

Version 2.6.0 picks up the work v2.5.0 explicitly deferred. v2.5.0 completed the
MS-VBAL §5.6 expression binder and closed the syntax corpus; everything it left
quiet was recorded as a documented deferral (see
`docs/static_analysis_completeness_2.5.0.md` and the v2.5.0 roadmap). v2.6.0 is
where those deferrals get the evidence (oracle / locale / reference metadata) or
the modeling (flow joins) they need to become provable diagnostics.

The same evidence-led, no-false-positive discipline carries over: a hard
diagnostic ships only when provably correct (MS-VBAL, the Excel/VBE oracle, or
deterministic XLIDE-owned metadata with tests), with a positive, a negative, and a
no-diagnostic control plus a named evidence source. Anything unknown, ambiguous,
`Variant`, or runtime-convertible stays quiet.

## Goal 1: Flow phase 2 — definite assignment

v2.5.0 shipped the precision groundwork: the `If`/`ElseIf`/`Else` branch-merge
intersects arm state in the shared dataflow (`walkBranchMergedBody`). Phase 2 adds
a new red and the CFG it needs.

- [ ] **Definite-assignment red (use-before-assignment).** Blocked on:
  - **Control-flow joins beyond structured `If`** — loops (`For`/`Do`/`While`),
    `GoTo`/`GoSub`, `On Error`/`Resume`. Today `procedureHasUnstructuredFlow` makes
    any such procedure fall back to the conservative straight-line walk, so the
    branch-merge never sees them. A definite-assignment red needs these edges
    modeled, which is the main false-positive risk.
  - **Oracle-pinned use-before-def cases.** VBA auto-initializes locals
    (`0` / `""` / `Nothing`), so a benign read of an unassigned scalar is *not* an
    error. The red must fire only where VBE itself flags it or a runtime fault is
    provable — pinned by the Excel/VBE oracle, not assumed.

## Goal 2: Binder-dependent reds that need more evidence

Each shipped in v2.5.0 only as far as provable; the rest deferred because VBA
coerces at runtime. v2.6.0 revisits them with targeted oracle work — and several
may resolve to *confirm-and-close as permanently quiet* rather than a new red.

- [ ] **Comparison / Boolean / string-concat scalar matrix.** Find the narrow
  cells (if any) VBE rejects at COMPILE rather than coercing at runtime; otherwise
  record the family as permanently quiet with the oracle controls that prove it.
- [ ] **Date coercion (broad).** The convert-vs-RTE-13 boundary is value- and
  locale-dependent; needs per-locale oracle pinning before any red. (Date-literal
  inner grammar stays Won't-Implement.)
- [ ] **Default-member-aware diagnostics.** Blocked on a host default-member
  resolver that carries the post-coercion scalar value type. The `VB_UserMemId`
  metadata infra already ships (v2.5.0); this is the consuming diagnostic.
- [ ] **Array-element typing / `ParamArray`-element / object-or-`Variant` ByRef.**
  Narrow follow-ons to `argument-shape-mismatch`: both-arrays-different-element-type
  (without double-reporting on the same span), and object/`Variant` ByRef once the
  oracle tests those shapes.

## Goal 3: Numeric / host boundary overflow

- [ ] **`Single`/`Double`/`Decimal` overflow** — not a clean compile/RTE-6 literal
  boundary; needs an oracle map of which literals VBE rejects.
- [ ] **Hex/octal width overflow** — needs the lexer to expose the declared literal
  width, then an oracle map of rejected width-suffix combinations.
- [ ] **`&`/`^`/`!`/`#`/`@` suffix overflow** — has a *proven* false positive
  (`s = 3000000000&"x"` compiles as concatenation; the lexer greedily glues the
  token), so any future rule must first solve next-token concat-vs-suffix
  disambiguation. `parseVbaIntegerLiteral` must stay out of the literal-typing path
  until then.

## Goal 4: Object-model member binding (carried from v2.5.0)

Product features the binder partly enables but that stand on their own:

- [ ] Declared `Event` / `WithEvents` / `RaiseEvent` binding beyond module-kind
  validation; the `WithEvents As Object` event-source type restriction (needs
  reference metadata).
- [ ] `Implements` interface-member completeness.
- [ ] External `.vbref.xml` object/member metadata as a diagnostic source.

## Also tracked (corpus deferrals, `managed_backlog.md` §C)

Recorded as defer-with-reason in v2.5.0; revisit when oracle-gateable:

- **Host-event wrong-signature binding** — wrong-signature `Worksheet_`/`Workbook_`/
  `App_` handlers compile cleanly and merely fail to bind (silent no-op), so the
  oracle accepts them; only a downgradeable *yellow* host-binding warning would
  ever be defensible, and it stays oracle- and module-kind-gated.
- **`On Error`/`Resume` well-formedness + unreachable-code** (ERROR_FLOW).
- **`Null`/`Empty` `Variant` operand coercion** (COERCE); public-member-exposes-
  Private-UDT (UDT_004); class/UserForm lifecycle + event-signature shape.
- **Continuation-count / line-length limits** — deterministic pure-counting checks,
  but the exact VBE boundary is not oracle-pinned and the value is low.

## Won't implement (unchanged)

Recorded in the MS-VBAL verification-map "Won't Implement" section — no oracle path
and near-zero payoff:

- Date-literal inner-grammar validation (locale-sensitive → false-positive-prone).
- Exact legacy-codepage non-Latin identifier ranges (old-VBE detail).
- UserForm/chart designer-backed (`.frm`/`.frx`) member parsing.

## Relationship to v2.5.0

v2.5.0 built the keystone expression binder and closed the corpus, shipping every
provable binder-dependent red and recording the rest as documented deferrals.
v2.6.0 is where those deferrals — definite assignment, the runtime-coercion
families, the numeric boundaries, and object-model binding — get the evidence or
modeling to become provable, or are confirmed-and-closed as permanently quiet.
