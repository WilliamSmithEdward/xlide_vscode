# XLIDE User Guides

This folder owns publish-ready, public-facing XLIDE user guides. If you are
using XLIDE rather than changing extension internals, start here.

Use `docs/` for internal roadmap, architecture, implementation policy, and
engineering notes. Use `user_guides/` for guides that should be suitable for
extension users to read directly.

Available guides:

- [getting_started.md](getting_started.md) - install/setup, opening workbooks, editing VBA,
  analysis, macro runs, module sync, documentation comments, and setup
  troubleshooting.
- [analysis.md](analysis.md) - live/current-module/workbook analysis,
  diagnostic tracking, severity overrides, untracked rules, and source
  suppression comments.
- [sync.md](sync.md) - exporting, importing, true-up modes, workbook sync
  settings, and safe module synchronization with `.bas`/`.cls` files.
- [testing.md](testing.md) - writing, running, filtering, reviewing, and automating
  `@xlide-test` workbook tests.
- [automation.md](automation.md) - AI-agent workbook discovery, editing, analysis, test runs,
  and CI artifact flow.
- [support.md](support.md) - safety model, Excel Trust Center notes, workbook mutation,
  support bundles, and recovery.

Planned guides:

- `workbooks.md` - workbook tree, workbook actions, and Excel integration
