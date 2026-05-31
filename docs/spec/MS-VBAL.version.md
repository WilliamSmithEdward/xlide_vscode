# MS-VBAL Specification Version

Canonical source for all VBA lexer/parser/grammar verification in XLIDE.

| Field | Value |
|---|---|
| Document | [MS-VBAL]: VBA Language Specification |
| Protocol revision | v20250520 |
| Microsoft publication date | May 20, 2025 |
| Local copy | `docs/[MS-VBAL].pdf` (288 pages) |
| Microsoft Learn landing page | <https://learn.microsoft.com/en-us/openspecs/microsoft_general_purpose_programming_languages/ms-vbal/d5418146-0bd2-45eb-9c7a-fd9502722c74> |

## Notes

- The PDF lives at the repository path `docs/[MS-VBAL].pdf` (the roadmap suggested
  `docs/spec/MS-VBAL.pdf`; the existing in-repo location is used to avoid
  duplicating the binary).
- All implemented grammar rules cite a specific MS-VBAL section in code comments
  and in `MS-VBAL.verification-map.md`.
- Casing of VBA keywords is taken from the section 3.3.5.2 reserved-identifier
  grammar, with the documented exception that literal-identifiers (true / false /
  nothing / empty / null) are rendered capitalized to match the VBA editor (VBE).
