# XLIDE User Guides

This folder owns publish-ready, public-facing XLIDE user guides. If you are
using XLIDE rather than changing extension internals, start here.

Use `docs/` for internal roadmap, architecture, implementation policy, and
engineering notes. Use `user_guides/` for guides that should be suitable for
extension users to read directly.

Available guides:

- [getting_started.md](getting_started.md) - install/setup, opening macro-enabled
  Office files (Excel, Word, PowerPoint, Access), editing VBA, analysis, macro
  runs, module sync, documentation comments, and setup troubleshooting.
- [analysis.md](analysis.md) - live/current-module/file analysis,
  diagnostic tracking, severity overrides, untracked rules, and source
  suppression comments.
- [sync.md](sync.md) - exporting, importing, true-up modes, per-file sync
  settings, and safe module synchronization with `.bas`/`.cls` files.
- [vba-doc-comments.md](vba-doc-comments.md) - inline `'''` XML documentation
  comments, module-header docs, and external `.vbref.xml` IntelliSense metadata.
- [testing.md](testing.md) - writing, running, filtering, reviewing, and automating
  `@xlide-test` unit tests in Excel, Word, and PowerPoint files.
- [automation.md](automation.md) - AI-agent file discovery, editing, analysis, test runs,
  and CI artifact flow.
- [support.md](support.md) - safety model, Trust Center notes, file mutation,
  support bundles, and recovery.

Planned guides:

- `files.md` - the file tree, per-container actions, and Office integration
