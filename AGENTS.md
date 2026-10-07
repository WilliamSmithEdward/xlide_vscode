# Notes for agents

<!-- repo-standards:begin. Copied from WilliamSmithEdward/repo-standards, templates/agents/AGENTS-block.md. Change it there; the weekly rescan fails a copy that differs. -->
## Releases, CI and security

These rules are the same in every WilliamSmithEdward repository.

- **How a release happens here:** pushing a `vX.Y.Z` tag runs Publish, which builds the release files in CI and creates the GitHub release with them, their signed provenance and the security reports. Any other step, such as a marketplace upload, is described elsewhere in this file.
- **Starting a workflow by hand never releases anything.** Publish and every
  release report are dry runs when started with `gh workflow run` or the Run
  workflow button. They build, scan and assemble the release files exactly
  as a release would, and upload them as the `release-preview` artifact
  instead. Run one after changing anything on the release path:
  `gh workflow run <file> --ref main`, then
  `gh run download <run-id> -n release-preview`.
- **Do not create, publish, edit or delete a release or a `v*` tag** unless
  the owner asks for it. A `v*` tag cannot be moved or deleted once pushed.
- **Every change to `main` goes through a pull request** that passes CI
  passed, Security passed and Malware scan passed. No one can push to `main`
  directly or skip the checks, admins included. Push a branch, open a pull
  request, and let it merge itself: `gh pr merge --auto --squash <number>`.
- **Pins.** Actions by full commit SHA with the version as a comment. Images
  by digest, in `.github/security/<tool>/Dockerfile`. Python tools from the
  hash-locked `.github/requirements/<purpose>.txt`, compiled from the `.in`
  beside it with
  `uv pip compile <purpose>.in --universal --generate-hashes --python-version 3.12 -o <purpose>.txt`.
  Runners are named releases, never `-latest`.
- **Updates merge themselves.** Dependabot and the Update YARA rules workflow
  open pull requests that merge once the three checks pass, except a
  third-party major version, which waits for the owner. Leave them alone
  unless asked.
- **A scanner finding is fixed or accepted with a written reason** in the
  repository's accepted list. Never silence a scanner without one.
<!-- repo-standards:end -->

## Releasing

The version lives in `package.json`, and CHANGELOG.md needs a
`## [X.Y.Z] - YYYY-MM-DD` section for it: the release's notes are taken from
there, and Publish fails without one. With both merged to `main`, the owner
tags that commit:

```bash
git tag vX.Y.Z && git push origin vX.Y.Z
```

Publish builds the vsix from the tag, scans it, signs its build provenance
and creates the GitHub release. The Marketplace upload stays with the owner,
and takes the vsix from that release rather than a local build, so the
Marketplace serves the signed file:

```bash
gh release download vX.Y.Z --pattern '*.vsix'
npx @vscode/vsce publish --packagePath xlide-X.Y.Z.vsix
```

## Editor responsiveness

The owner wants completion menus to update nearly instantaneously while typing
(including `ThisWorkbook.Sheets(1).`) and symbol hover to be as snappy as
possible. Treat typing latency, menu updates and mouse hover as product
priorities, including slow outliers and cold-start behavior.

- Return available cached/current-module results immediately. Do not wait for
  full-project loading before a completion update, or before returning a hover
  that can already be resolved.
- Do not introduce fixed sleeps or debounce delays on the response path to
  trade responsiveness for a more complete first result. Load additional facts
  in the background and keep incomplete completion lists refreshable.
- Keep background indexing from synchronously running ahead of an available
  response. Avoid module-sized scans/projections per keystroke or mouse move
  when a token, logical line or unchanged source snapshot suffices.
- Verify changes with meaningful work-count/scheduling regressions and the
  integration harness. Report both typical latency and slow samples, separate
  warm behavior from startup, and distinguish provider response from menu or
  tooltip painting. Fast analyzer averages alone do not establish a snappy UI.
- Continue performance hunting proactively within the requested editor
  surfaces. The owner authorizes using the integration harness and creating
  issues and pull requests; the release rules above still apply.
