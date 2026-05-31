# XLIDE VBA Documentation Comments and Metadata

XLIDE gives your VBA code Visual-Studio-style IntelliSense documentation. You
describe a procedure, type, enum, `Declare`, or module-level variable once -
either with an inline `'''` comment directly above it, or in an external
metadata file - and that description appears in **hovers** and **call tips**
(signature help) throughout the project.

Both authoring paths share **one XML vocabulary**, so anything you can write
above a declaration you can also write in a metadata file, and vice versa.

---

## 1. Inline documentation comments

A documentation comment is a run of one or more **contiguous** lines whose first
non-space characters are `'''` (three apostrophes), placed **directly above** a
declaration - exactly like C# `///` comments in Visual Studio.

```vba
 
```

Rules:

- The block must sit on the lines immediately above the declaration. A blank
  line or an ordinary `'` comment between the block and the declaration ends it.
- Ordinary single-apostrophe comments (`'`) are **not** documentation and are
  ignored.
- A leading space after `'''` is stripped, so `''' <summary>` and
  `'''<summary>` are equivalent.
- If the block contains no recognized tags, its entire text becomes the
  `<summary>`. This makes a one-line plain-text note work:
  `''' Adds an item to the cart.`

Documentable declarations: `Sub`, `Function`, `Property Get/Let/Set`, `Type`,
`Enum`, `Declare`, and module-level `Dim`/`Public`/`Private`/`Const` variables.

---

## 2. The tag vocabulary

| Tag | Use | Shown in |
| --- | --- | --- |
| `<summary>` | One-line description of the symbol. | Hover + call tip |
| `<param name="X">` | Description of parameter `X` (matched case-insensitively). | Hover + per-parameter call tip |
| `<returns>` | Description of a function's return value. | Hover + call tip |
| `<remarks>` | Extended notes shown below the summary. | Hover + call tip |
| `<example>` | A usage example, rendered as a VBA code block. | Hover |
| `<signature>` | An explicit call-tip signature. **External files only** - see below. | Call tip |

Notes:

- The five predefined XML entities are decoded: `&lt; &gt; &amp; &quot; &apos;`.
- Whitespace inside a tag is collapsed to single spaces, except `<example>`,
  whose line layout is preserved for the code block.
- Parsing is lenient: a partially written or slightly malformed block still
  yields whatever text it can.

---

## 3. External metadata files

To document symbols you do not own (host object-model members, VBA runtime
functions) or to share documentation across a team, put `<member>` entries in a
metadata file anywhere in your workspace. The default discovery glob is
`**/*.vbref.xml`.

```xml
<xlideDoc>
  <!-- Document one of your own procedures by qualified name. -->
  <member name="Module1.ComputeTax">
    <summary>Returns the tax owed for a pre-tax amount.</summary>
    <param name="Amount">The pre-tax amount, in dollars.</param>
    <returns>The tax owed, in dollars.</returns>
  </member>

  <!-- Re-describe a built-in for your team. -->
  <member name="MsgBox">
    <summary>Team note: prefer the Notify helper over raw MsgBox.</summary>
  </member>

  <!-- Give a call-tip signature to a symbol XLIDE cannot otherwise resolve. -->
  <member name="DoThing">
    <signature>DoThing(Path As String, Optional Retry As Boolean) As Boolean</signature>
    <summary>Performs the thing at the given path.</summary>
    <param name="Path">Filesystem path to operate on.</param>
  </member>
</xlideDoc>
```

The `name` attribute is either:

- **Qualified** - `Module.Symbol` or `ReceiverType.Member` (e.g.
  `Range.Value`). Most specific; matched first.
- **Bare** - `Symbol` (e.g. `MsgBox`). Matches by name anywhere.

The wrapping element name (`<xlideDoc>` above) is not significant; only
`<member>` elements are read, so a `.NET`-style `<doc><members>...` wrapper works
too. A single malformed `<member>` never discards the rest of the file.

`<signature>` is honored **only** when XLIDE has no real signature for the
callee (from your source or the curated host/runtime library). It lets you give
a call tip to an otherwise-unknown procedure; it never overrides a real one.

---

## 4. Precedence: developer-defined overrides the library

When more than one source describes the same symbol, XLIDE resolves the tooltip
in this order (first match wins):

1. **Inline `'''` comment** on the actual declaration - closest to your code,
   most authoritative for symbols you own.
2. **External metadata file** `<member>` entry - overrides the built-in library
   and can document symbols that have no inline comment.
3. **Built-in curated library** - the verified Excel object model and VBA
   runtime signatures shipped with XLIDE (signature only; no prose).

In short: **developer-defined metadata (inline or external) overrides the
curated library.** Among developer sources, an inline comment on the declaration
wins over an external entry for the same symbol.

---

## 5. Discovery and usage paths

- **Inline:** nothing to configure. Type `'''` above any declaration; hovers and
  call tips update live as you edit (no save required).
- **External files:** placed anywhere in any workspace folder, matching the glob
  `xlide.docs.metadataGlob` (default `**/*.vbref.xml`; `node_modules` is
  excluded). They are loaded on startup and **reloaded automatically** when a
  matching file is created, changed, or deleted.

### Settings

| Setting | Default | Effect |
| --- | --- | --- |
| `xlide.docs.enabled` | `true` | Master switch for inline + external documentation in tooltips. |
| `xlide.docs.metadataGlob` | `**/*.vbref.xml` | Glob (per workspace folder) that locates external metadata files. |

---

## 6. Implementation map

| Concern | Location |
| --- | --- |
| Doc model + Markdown rendering | `src/analyzer/docs/docModel.ts` |
| Inline `'''` parser + shared XML body parser | `src/analyzer/docs/docComment.ts` |
| External metadata-file parser | `src/analyzer/docs/externalDoc.ts` |
| Lookup registry (precedence) | `src/analyzer/docs/docRegistry.ts` |
| Inline doc attached to symbols | `src/analyzer/symbols/buildModuleSymbols.ts` |
| Hover integration | `src/analyzer/hover/resolveHover.ts` |
| Call-tip integration | `src/analyzer/signature/signatureHelp.ts` |
| Workspace file loader + watcher | `src/vbaDocMetadata.ts` |
| Provider wiring | `src/vbaMemberCompletion.ts`, `src/vbaLanguageProviders.ts` |
| Tests | `tests/vbaDocComments.test.ts` |
