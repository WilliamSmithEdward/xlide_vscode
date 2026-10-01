# Security policy

## Reporting a vulnerability

Report a vulnerability privately, not in a public issue or pull request:
[open a private report](https://github.com/WilliamSmithEdward/xlide_vscode/security/advisories/new).
Only the maintainer sees it. Include the XLIDE version and the host (Excel,
Word, PowerPoint, Access or VB6), and the smallest file or steps that show
it, with credentials and private data removed.

A confirmed vulnerability is fixed in a release on the Visual Studio
Marketplace, and the advisory is published with it,
crediting you unless you ask otherwise.

## Supported versions

Only the latest release on the Visual Studio Marketplace receives security
fixes. Older releases are not maintained separately; update when a fix ships.
VS Code updates extensions automatically, so that is the version nearly
every user runs.

## Scope

XLIDE opens Office and VB6 files that other people may have written, serves
a local API to the processes of the signed-in user, and registers tools a
chat can call. A file, a request or a tool call that makes XLIDE run code,
write outside what it was asked to change, or reach another user's data is a
vulnerability here.

### What XLIDE touches

- **Office files you open.** XLIDE reads and writes the VBA project inside
  the file itself. VBA runs only through the Office application, when you run
  it or approve a chat's request to run tests.
- **A local API server.** Each VS Code window listens on 127.0.0.1 at a port
  the system picks. Every request needs a random token, compared in constant
  time, from a record only your user account can read.
- **Language model tools.** The tools XLIDE registers act on the Office
  files and modules a chat asks them to. Every tool that changes a file, and
  the one that runs VBA tests, asks before it acts; the read-only tools run
  when called.
- **No telemetry.** XLIDE makes no network requests of its own.

## How the code is checked

Three workflows check every pull request and every push to `main`, and
their gates decide whether a change can merge: **CI passed**,
**Security passed** and **Malware scan passed**. A gate passes only when
every job before it did, and any unexpected finding fails it, whatever its
severity. Security also runs weekly and Malware scan daily, and both run on
a published release.

- **Code:** CodeQL with the `security-extended` queries, for the extension's
  TypeScript and the GitHub Actions workflows, and Semgrep with `p/default`,
  `p/typescript` and `p/github-actions` over `src`, `scripts` and `.github`.
  CodeQL leaves out tests, the syntax corpus and build output. Results go to
  the repository's code scanning.
- **Workflows:** zizmor audits the GitHub Actions workflows; a finding fails
  Security.
- **Dependencies:** `npm audit` over every dependency, production and
  development; any vulnerability fails it.
- **Malware:** ClamAV, with signatures freshclam fetches and verifies on
  every run, and YARA-X, with the YARA Forge rules pinned to a release and
  its SHA-256, scan every tracked file, every dependency npm installs from the
  lockfile with no install script run, and the extension package, as the
  vsix and unpacked, built from that commit. In a release that vsix is the
  file the release carries. YARA-X uses the YARA Forge core rules. Each scan must also detect
  the EICAR test file written for the run, so a scan that read nothing
  cannot pass. The signatures are cached a day at a time; if freshclam fails,
  the scan uses cached signatures up to two days old, with a warning, and
  fails without them.
- **Fuzzing:** fast-check properties in [`tests/properties`](tests/properties)
  check the readers of files XLIDE did not write. Office files with bytes
  changed, and streams of random bytes, go through the ZIP, deflate, compound
  file, MS-OVBA and VBA project readers, and through opening a workbook,
  document, presentation or Access database end to end: each must read the
  file or refuse it with its own error, within a time limit. The codecs must
  read back what they write, and inflate must agree with zlib. Generated VBA
  must lex back to its exact text, parse without throwing, and format to the
  same text a second time; VB6 form headers and project manifests must print
  back as they were read. `npm test` runs every property a hundred times. The
  Fuzz workflow runs on every change to those readers (`src/vba`, the
  analyzer's lexer, parser and formatter, and `tests/properties`) and daily,
  twenty thousand times per property on a change and two hundred thousand
  daily. It is not a gate: a finding becomes a regression test with its fix.
- **OpenSSF Scorecard** rates the repository's security practices on every
  change to `main` and weekly, and the README badge shows the result.
  One of its checks does not fit this project: a single maintainer cannot
  have a second person approve every change. Signed-Releases rises as
  releases carry the provenance bundle; it counts the last five.

## Accepted findings

A finding is fixed, or accepted with a written reason in
[`.github/codeql/reviewed.json`](.github/codeql/reviewed.json) for CodeQL
and [`.github/scans/reviewed.json`](.github/scans/reviewed.json) for ClamAV
and YARA-X. Each entry names the scan it applies to. A CodeQL entry
accepts one result: it matches the rule, the exact path and the text of the
flagged line in the checkout, trimmed. It follows the line when code above it
moves, and the result comes back for review when the line itself changes or
the rule flags another line of the file. A ClamAV or YARA-X detection is of a
whole file, so its entry matches the signature or rule name and a file path,
where `*` matches within one path segment. An entry that no longer matches
any result fails its scan, so remove it in the change that makes it stale.
On a pull request CodeQL reports only the changed code, so a CodeQL entry
outside it is listed there but not failed; the run on `main` fails on it.
zizmor keeps its exceptions in
`.github/zizmor.yml` or inline beside the line they excuse, each with its
reason.

The current entries:

- CodeQL, eleven entries: `js/insufficient-password-hash` on SHA-256
  content tokens in `src/moduleContentToken.ts` and
  `src/projectModuleOperations.ts`, where no password is involved;
  `js/incomplete-multi-character-sanitization` on two lines of
  `src/vba/xlsxShapes.ts`, which strip tags to make a plain-text shape
  caption; `js/incomplete-url-substring-sanitization` in `src/vba/xlsx.ts`,
  which tests for an XML namespace and fetches nothing; and
  `js/missing-origin-check` on the message listener of six pages in
  `assets/webview/`, which allow no frames, so only VS Code's webview host
  can post to them.
- Semgrep, two rules left out of the scan in
  [the Security workflow](.github/workflows/security.yml), which records the
  reasons beside them: regular expressions built from a variable are built
  from VBA identifiers the code has already parsed, or are the search tool's
  documented regex mode; and paths joined from a variable stay inside your
  workspace, your Office files and XLIDE's own state directory.
- ClamAV and YARA-X: there are none.
- zizmor: the `self-repository` rule is turned off in `.github/zizmor.yml`
  until GitHub documents that syntax for called reusable workflows.

## Pinning and updates

Everything the workflows run is pinned: actions to full commit SHAs,
runners to named OS releases, scanner images to digests, Python tools to
hash-locked lock files, the project's own dependencies to exact versions in
`package.json` and the npm lockfile alike, Node to an exact version, and the
YARA-X engine and YARA Forge rules to a release and its SHA-256. ClamAV's
signatures change too often to pin, so freshclam fetches and verifies them
on every run, and the malware scan report records their versions. Semgrep's
registry rules are also fetched on every run.

Dependabot proposes updates to npm, GitHub Actions, the ClamAV and Semgrep
images, and the hash-locked zizmor requirements once a version is a week
old, and at once for a security advisory. The Update YARA rules workflow
proposes new YARA pins each week. A minor or patch update, and the YARA
pull request, merges itself once CI, Security and Malware scan pass; a
third-party major version waits for review.

## Releases

Pushing a `vX.Y.Z` tag runs the Publish workflow. It builds the vsix from
the tagged commit in CI, refusing a tag that is not `package.json`'s
version, runs Security and Malware scan on that commit and that file, and
only when both pass signs the vsix's build provenance and creates the GitHub
release with:

- `xlide-<version>.vsix`: the extension, the file uploaded to the
  Marketplace.
- `xlide-<version>.sigstore.json`: the signed build provenance.
- `security-report.md` and `security-sarif.zip`: the code checks' verdicts,
  findings and raw results.
- `malware-scan-report.md` and `malware-scan-results.zip`: the scans'
  verdicts, detections and signature versions, with the vsix's SHA-256.

To check that a vsix was built by this repository's Publish workflow from a
tagged commit:

```bash
gh attestation verify xlide-<version>.vsix --repo WilliamSmithEdward/xlide_vscode
```

The Marketplace upload is by hand, of the vsix on the release. Releases up
to 10.14.6 were built locally and carry no provenance.

## Repository settings

<!-- repo-standards:begin security-settings. Copied from WilliamSmithEdward/repo-standards, templates/security/settings-block.md. Change it there; the weekly rescan fails a copy that differs. -->
- `main` accepts changes only through a pull request that passes
  **CI passed**, **Security passed** and **Malware scan passed**. The
  ruleset has no bypass, for the owner either, and refuses force-pushes and
  deleting the branch.
- A `v*` release tag cannot be moved or deleted once pushed, except by a
  repository admin.
- A workflow that uses an action not pinned to a full commit SHA fails to
  run. Workflow tokens are read-only unless a job is granted more for
  itself.
- Secret scanning with push protection, Dependabot alerts and security
  updates, and private vulnerability reporting are on.
<!-- repo-standards:end -->
