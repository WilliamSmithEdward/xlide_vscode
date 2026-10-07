# Scripting and RegExp reference metadata

Issue #1240 adds Microsoft Scripting Runtime 1.0 and Microsoft VBScript Regular
Expressions 5.5 to the portable analyzer models. Project reference GUIDs select
them through the existing registry, retaining reference order and the Office
host. This supports desktop, browser, and analysis-worker consumers without
runtime COM extraction. Other third-party libraries are outside this change.

Reference JSON uses the same pyVBAReference scraper as the Office corpus,
pinned to `a6c39d0617085297e88bec7de0ee8afb7aa0746b`. On Windows with the
libraries registered, a clean checkout at that revision, and pywin32 311:

```powershell
python scripts/dump-scripting-typelibs.py path/to/pyVBAReference
node scripts/generate-host-object-model.mjs scripting
node scripts/generate-host-object-model.mjs regexp
```

The source corpus remains under the existing ignored `reference/` directory;
the generated TypeScript snapshots carry provenance and are committed. The
adapter also records COM default-interface identities so `GetFile` and
`OpenAsTextStream` chain through their corresponding VBA classes while retaining
the original declared signatures. These models remain non-exhaustive and do
not assert that an unlisted member is invalid.
