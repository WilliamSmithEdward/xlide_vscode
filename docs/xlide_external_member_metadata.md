# XLIDE External Object/Member Metadata

Status: planned downstream-developer documentation, not implemented behavior.

Purpose: describe the intended workflow for using explicit metadata files to
power `object.` member completion, hover, signature help, return-type chaining,
and future member diagnostics for APIs XLIDE cannot fully read from workbook
source.

This document should become the user-facing how-to before external object/member
metadata is considered shipped.

## Goals

- Let developers describe referenced libraries, add-ins, host extensions, or
  local APIs in explicit metadata files.
- Use the same explicit metadata for member completion, hovers, call tips, and
  type resolution.
- Keep workbook source authoritative for workbook-owned classes.
- Keep all behavior deterministic: no inferred members from names, comments, or
  partial matches.

## Intended Developer Workflow

1. Create or update a workspace metadata file matching
   `xlide.docs.metadataGlob` (default `**/*.vbref.xml`).
2. Declare external object types and their members explicitly.
3. Open or reload the workbook in XLIDE.
4. Type a variable declaration using the metadata-backed type.
5. Type `object.` and review the member menu.
6. Hover members or invoke signature help to confirm docs and signatures.

Workbook-defined class modules should not require external metadata for their
own members once workbook class-member completion is implemented. Source wins
for those members; external metadata can still enrich documentation where the
final precedence rules allow it.

## Planned Metadata Shape

Exact XML is not final. The object-member extension should remain compatible
with the existing `.vbref.xml` discovery path and should prefer an explicit type
section plus explicit member entries.

```xml
<xlideDoc>
  <type name="Acme.CustomerRepository" kind="class">
    <summary>Repository for customer records exposed by the Acme add-in.</summary>
  </type>

  <member name="Acme.CustomerRepository.FindById"
          kind="function"
          returns="Acme.Customer">
    <signature>FindById(ByVal Id As String) As Acme.Customer</signature>
    <summary>Finds a customer by external id.</summary>
    <param name="Id" type="String">The external customer id.</param>
    <returns type="Acme.Customer">The matching customer object.</returns>
  </member>

  <member name="Acme.CustomerRepository.Count"
          kind="property-get"
          returns="Long">
    <summary>The number of cached customer records.</summary>
  </member>
</xlideDoc>
```

Given:

```vba
Dim repo As Acme.CustomerRepository
repo.
```

XLIDE should suggest `FindById` and `Count`, show the supplied documentation,
and use `FindById`'s return type for later chained completion when possible.

## Required Metadata Fields

The final schema should document:

- Type name and kind.
- Member qualified name.
- Member kind: method, function, property get/let/set, event, field, enum
  member, or constant.
- Return type for functions and property gets.
- Parameter list, parameter types, optional/default metadata, and `ParamArray`
  where applicable.
- Human-readable docs: summary, params, returns, remarks, and examples.
- Provenance: whether the metadata came from vendor docs, generated type
  library output, project docs, or developer-authored declarations.

## Precedence

Planned deterministic precedence:

1. Workbook source symbols for workbook-owned classes/modules.
2. Inline `'''` documentation on the source declaration.
3. External metadata for explicitly declared external or extension members.
4. Curated XLIDE host/runtime metadata.

External metadata must not silently replace a concrete workbook source
signature. If an override mode is ever added, it should be explicit, visible in
the UI, and covered by tests.

## Completion Behavior

XLIDE should show a member menu only when the receiver type is known:

```vba
Dim repo As Acme.CustomerRepository
repo.   ' show metadata-backed members
```

XLIDE should stay silent when the receiver type is unknown or ambiguous:

```vba
Dim repo As Variant
repo.   ' no metadata-backed certainty
```

Return-type chaining should use only explicit return metadata:

```vba
repo.FindById("C-100").   ' customer members only if Acme.Customer is explicit
```

## Troubleshooting Topics for the Shipped Guide

- Confirming the metadata file matches `xlide.docs.metadataGlob`.
- Reloading or watching metadata changes.
- Diagnosing duplicate type/member declarations.
- Understanding source-vs-metadata precedence.
- Explaining why no menu appears for `Variant`, unresolved, or ambiguous types.
- Showing a metadata validation report in the future XLIDE sidebar.
