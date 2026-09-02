# What a host must tell the analyzer

The analyzer reads one thing for free: the module's **source text**. Everything
else about a module is either derived from that text or supplied by the host,
and the facts a host has to supply are exactly the ones the text cannot carry.

This is written for a host outside this repo - the VBE add-in, whose
`CodeModule` returns a module's CODE and nothing else. The `VERSION`/`BEGIN`
block, every `Attribute VB_` line, and a UserForm's designer all live in the
EXPORTED file, not in the code pane. A host that can see them should say so; a
host that cannot must say nothing rather than guess.

## The surfaces

`predeclaredId` and the rest appear on every layer, so a host can enter wherever
it already talks to the analyzer:

| surface | file | used by |
| --- | --- | --- |
| `ModuleInput` | `analyzer/symbols/projectIndex.ts` | `ProjectIndex.setModule` |
| `VbaProjectModuleInput` | `vbaProjectAnalysis.ts` | `buildVbaProjectIndex` |
| `WorkerSeedModule` | `analysisWorkerProtocol.ts` | the worker `seed` message |
| `ModuleEntry` | `vba/projectService.ts` | this repo's own file reader |

`ModuleInput` and `ProjectIndexOptions` are exported from `analyzer/index.ts`
so an external host can import the type it writes against.

## The fields

### `source` - required

The module's code. **Every diagnostic offset is measured against exactly this
string.** Do not prepend a synthesized attribute header to smuggle in a fact:
it shifts every span by the length of what you added. That is what the fields
below are for.

### `moduleName` - required

Also the fallback input for `documentType` (below), so send the VBE **code
name**, not a sheet's display name.

### `moduleKind` / `type` - required in practice

`standard` | `class` | `document` | `userform`.

Absent maps to `standard`, and so does **any unrecognised string** -
`moduleKindFromType` has no error path. A typo degrades silently: object-module
rules stop running and `Me` starts reporting `me-outside-object-module` in a
class that is perfectly valid.

### `documentType` - optional, inferred when absent

`workbook` | `worksheet` | `chart` | `document`. Decides which event handlers
belong in the module (`Workbook_Open` in a worksheet module is a real finding -
Excel never wires it there).

Absent falls back to the code name, which is right for every conventional one:

| code name | inferred |
| --- | --- |
| `ThisWorkbook` | workbook |
| `ThisDocument` | document |
| `Chart`, `Chart1`, ... | chart |
| anything else | worksheet |

So a host that cannot answer is fine on conventional projects. The gap is a
**renamed non-worksheet code name** - a chart sheet whose code name is
`Dashboard` infers worksheet, and its `Chart_*` handlers report. Send the field
if you can read it.

### `implicitMembers` - three states

A UserForm's designer-declared controls. **A list**, an **empty array** meaning
"this form genuinely has none", or **absent** meaning "nobody read the
designer".

Absent and empty are different claims and the analyzer treats them differently.
Reading absent as empty is what turned a running form's code-behind red against
source that had just compiled (#48) - the VBE stops handing out a designer once
the form has been shown.

When absent, the analyzer parses the module's own `.frm` header, which only a
standalone VB6-style export carries.

### `predeclaredId` - three states

`Attribute VB_PredeclaredId = True`, which gives a class module a default
instance and makes its own name usable as a value.

**`true`**, **`false`**, or **absent** meaning the attribute header was never
read. Only a vouched-for `false` reports, because the attribute is invisible in
the code pane and a wrong `false` puts `Variable not defined` on every use of a
legitimately predeclared singleton - `fullBuild.xlsm` alone carries 12 of them,
the whole stdVBA library.

If reading the attribute per class module is expensive enough to tempt you into
sending `false` wholesale, **send nothing instead**. Absent is silent; wrong is
loud.

When absent, the analyzer parses the module's own text for the attribute, which
again only a standalone export carries.

### `conditionalCompilation` - optional

```ts
{ compilerConstants?: Record<string, ConditionalValue>,
  projectConstants?:  Record<string, ConditionalValue> }
```

Decides which `#If` branch is live, and therefore which code is analyzed at
all. Absent uses the built-in defaults (VBA7 and friends). A host that knows the
real bitness and the project's `#Const` values should send them.

**The nesting is load-bearing.** A bare `{ VBA7: false }` sets no constant and
is silently ignored - the built-in defaults still decide, so it looks like it
worked. It must be `{ compilerConstants: { VBA7: false } }`.

### `host` - optional, per analyze request

`excel` | `word` | `powerpoint` | `access` | ... . Selects the object model.

Three outcomes, and the middle one matters:

| value | model |
| --- | --- |
| absent, or `excel` | Excel's - the default (#28) |
| a host with a model | that host's |
| a named host with no model yet | the empty model: every lookup misses, so **nothing is asserted** (#24) |

This also decides host-specific syntax. `[A1]` is `Application.Evaluate`
shorthand in Excel and reports as undefined in Word, so naming the host wrongly
moves real findings.

## What a host cannot supply yet

`Attribute VB_UserMemId = 0` marks a class's **default member**, so
`bag("key")` means `bag.Item("key")`. It is the same shape as `predeclaredId` -
hidden in the code pane, present only in the export - and the analyzer reads it
from source text only. There is no field for it, so a VBE host's class members
always answer `defaultMember: undefined`.

Nothing in this repo depends on the flag today; it is carried on the completion
result for consumers rather than driving any rule here. So this costs nothing
yet. If a host starts using it, the field should be added the same way
`predeclaredId` was, with the same three states.

## The rule behind all of it

Where a fact has three states, **absent is not a synonym for the negative one.**
Three separate defects have come from collapsing them - a form's control list
(#48), a class's default instance (#47), and a host's identity (#24) - and each
time the symptom was the same: correct code reported as broken, in the editor
where the developer could see it was fine.

A host that does not know should not answer.
