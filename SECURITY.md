# Security

## Reporting a vulnerability

Please report a vulnerability privately, not in a public issue:
[open a private report](https://github.com/WilliamSmithEdward/xlide_vscode/security/advisories/new).
Only the maintainer sees it. Include the XLIDE version, the host (Excel,
Word, PowerPoint, Access or VB6), and the smallest file or steps that show
it.

You will get an answer within a week. A confirmed issue is fixed in a
release on the Visual Studio Marketplace, and the advisory is published with
it, crediting you unless you ask otherwise.

## Supported versions

Only the latest release on the Marketplace receives fixes. VS Code updates
extensions automatically, so that is the version nearly every user runs.

## What XLIDE touches

- **Office files you open.** XLIDE reads and writes the VBA project inside
  the file itself. Code runs only when you run it, through the Office
  application.
- **A local API server.** Each VS Code window listens on 127.0.0.1 at a port
  the system picks. Every request needs a random token, compared in constant
  time, from a record only your user account can read.
- **Language model tools.** The tools XLIDE registers act on the Office
  files and modules a chat asks them to. The module write, edit and rename
  tools ask before they change anything; the others run when called.
- **No telemetry.** XLIDE makes no network requests of its own.

## How the code is checked

Every push and pull request to `main`, a weekly schedule, and every release
run five checks in [the Security workflow](.github/workflows/security.yml).
Each fails on any finding.

| Check | Covers |
| --- | --- |
| CodeQL, `security-extended` | The extension's TypeScript and the GitHub Actions workflows |
| Semgrep, `p/default`, `p/typescript` and `p/github-actions` | The same, with a second engine |
| `npm audit` | Every dependency, production and development |
| ClamAV, with the current signatures | Every file in the repository, every installed dependency, and the extension package |
| YARA-X, with the current [YARA Forge](https://yarahq.github.io/) core rules | The same files |

Two Semgrep rules are excluded, each reviewed finding by finding. Regular
expressions built from a variable are built from VBA identifiers the code
has already parsed, or are the search tool's documented regex mode. Paths
joined from a variable stay inside your workspace, your Office files and
XLIDE's own state directory. The workflow file records the reasons beside
the exclusions.

ClamAV and YARA-X see the dependencies as npm installs them with no install
script run, and the package both as the vsix and unpacked. On a release they
scan the vsix attached to it, the file published to the Marketplace; on any
other run, a package built from that commit. Each scan must also detect the
EICAR test file written for the run, so a scan that read nothing cannot
pass. A detection found harmless is recorded with its reason in
[`.github/scans/reviewed.json`](.github/scans/reviewed.json).

Every GitHub release carries the vsix published to the Marketplace,
`security-report.md`, the verdicts and every finding for that commit with
the vsix's SHA-256, and `security-sarif.zip`, the raw results.

Dependabot opens a pull request as soon as an advisory affects a
dependency, and groups routine updates weekly. Workflow actions are pinned
to commit SHAs, the ClamAV engine to an image digest, and YARA-X to a
version and its SHA-256.
