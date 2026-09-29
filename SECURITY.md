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

Two workflows check every push and pull request to `main` and every
release, and each check fails on any finding.

[The Security workflow](.github/workflows/security.yml) also runs weekly:

| Check | Covers |
| --- | --- |
| CodeQL, `security-extended` | The extension's TypeScript and the GitHub Actions workflows |
| Semgrep, `p/default`, `p/typescript` and `p/github-actions` | The same, with a second engine |
| `npm audit` | Every dependency, production and development |

Two Semgrep rules are excluded, each reviewed finding by finding. Regular
expressions built from a variable are built from VBA identifiers the code
has already parsed, or are the search tool's documented regex mode. Paths
joined from a variable stay inside your workspace, your Office files and
XLIDE's own state directory. The workflow file records the reasons beside
the exclusions.

[The Malware scan workflow](.github/workflows/malware-scan.yml) also runs
daily:

| Scan | Covers |
| --- | --- |
| ClamAV, with signatures freshclam brings up to date on every run | Every file in the repository, every installed dependency, and the extension package |
| YARA-X, with the pinned [YARA Forge](https://yarahq.github.io/) core rules | The same files |

ClamAV and YARA-X see the dependencies as npm installs them with no install
script run, and the package both as the vsix and unpacked. On a release they
scan the vsix attached to it, the file published to the Marketplace; on any
other run, a package built from that commit. Each scan must also detect the
EICAR test file written for the run, so a scan that read nothing cannot
pass. A detection found harmless is recorded with its reason in
[`.github/scans/reviewed.json`](.github/scans/reviewed.json).

The YARA Forge release is pinned with the SHA-256 of its rules in
[`.github/scans/yara-forge.json`](.github/scans/yara-forge.json). Each week
[an update workflow](.github/workflows/update-yara-rules.yml) proposes the
newest release in a pull request, and both workflows scan the pull request
before it is merged.

Every GitHub release carries the vsix published to the Marketplace;
`security-report.md` and `security-sarif.zip`, the code checks' verdicts,
findings and raw results; and `malware-scan-report.md` and
`malware-scan-results.zip`, the scans' verdicts, detections and signature
versions, with the vsix's SHA-256.

Each workflow ends in one gate, `CI passed`, `Security passed` and
`Malware scan passed`, which passes only when every check before it did.
The ruleset in [`.github/rulesets/main.json`](.github/rulesets/main.json)
requires the three gates before a pull request merges into `main`, and
refuses force-pushes and deleting the branch. The repository also requires
every action a workflow uses to be pinned to a full commit SHA.

Everything the checks and the build run on is pinned: workflow actions to
commit SHAs, the Semgrep and ClamAV images to digests, YARA-X to a version
and its SHA-256, the YARA Forge rules to a release and its SHA-256, runners
to an OS release, Node to a version, and npm packages to exact versions in
`package.json` and the lockfile alike. Two things change too often to pin:
the ClamAV signatures, which freshclam verifies on every run and the malware
scan report records, and Semgrep's registry rules, fetched on every run.
Dependabot opens a pull request as soon as an advisory affects a
dependency, and proposes routine updates weekly, once a version is a week
old.
