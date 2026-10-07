"""Propose the newest YARA pins the malware scan may take.

This file is the same in every repository; the canonical copy lives in
WilliamSmithEdward/repo-standards (templates/yara/yara_update.py) and the
rescan checks each copy against it.

The pins live in .github/security/yara.json:

    {
      "yara_forge": {"release": "20260927", "asset": "yara-forge-rules-full.zip",
                     "url": "https://github.com/...", "sha256": "..."},
      "yara_x": {"release": "v1.20.0",
                 "asset": "yara-x-{release}-x86_64-unknown-linux-gnu.tar.gz",
                 "url": "https://github.com/...", "sha256": "..."}
    }

`yara_x` is present only where the scan downloads the YARA-X binary; where
YARA-X comes from a Python lock, Dependabot moves it. A scanner downloads each
`url` and checks it against `sha256` before use.

Rules, from the standard:
- YARA Forge: the newest release is taken at once. It is detection content,
  and the pull request is scanned before anyone merges it.
- YARA-X: the newest full release at least seven days old.
- Every SHA-256 is the digest GitHub records for the release asset, never
  one worked out from a download. A pinned release whose digest changed is an
  error, not an update. A release older than the pin is never proposed; a pin
  moved ahead of the cooldown by hand is kept.

Usage, from the two jobs of update-yara-rules.yml:

    python .github/security/yara_update.py prepare --out DIR
    python .github/security/yara_update.py propose --in DIR

`prepare` needs only read access. It writes the proposal to DIR and sets
`changed` in $GITHUB_OUTPUT. `propose` runs with write access and trusts
nothing in DIR it cannot check: the proposal may change only the pin file,
only in the fields a move allows, from the commit the workflow ran on. It
commits through the GitHub API (no git credentials on disk), opens the pull
request with a GitHub App token, which triggers normal pull-request checks
without manual workflow approval. It turns
on auto-merge, so the pull request merges itself once those pass and stays
open if any fails. It never reopens a pull request someone closed, and never
overwrites a branch it did not just create.
"""
from __future__ import annotations

import argparse
import copy
import datetime as dt
import json
import os
import re
import subprocess
import sys
from collections.abc import Callable
from pathlib import Path
from typing import Any

PIN_PATH = ".github/security/yara.json"
SOURCES = {
    "yara_forge": {"repo": "YARAHQ/yara-forge", "name": "YARA Forge", "cooldown_days": 0},
    "yara_x": {"repo": "VirusTotal/yara-x", "name": "YARA-X", "cooldown_days": 7},
}
SHA256 = re.compile(r"[0-9a-f]{64}")

Api = Callable[..., Any]


class ApiError(RuntimeError):
    pass


def gh_api(endpoint: str, method: str = "GET", payload: Any = None) -> Any:
    args = ["gh", "api", "--method", method, endpoint]
    if payload is not None:
        args += ["--input", "-"]
    result = subprocess.run(args, input=json.dumps(payload) if payload is not None else None,
                            capture_output=True, text=True, encoding="utf-8", timeout=120)
    if result.returncode != 0:
        raise ApiError(f"{method} {endpoint}: {(result.stdout + result.stderr).strip()}")
    return json.loads(result.stdout) if result.stdout.strip() else None


def asset_pin(release: dict[str, Any], asset_template: str, repo: str) -> dict[str, str]:
    """The pin for one asset of `release`, with GitHub's recorded digest."""
    tag = release["tag_name"]
    name = asset_template.format(release=tag)
    for asset in release.get("assets", []):
        if asset["name"] != name:
            continue
        recorded = asset.get("digest") or ""
        if not recorded.startswith("sha256:") or not SHA256.fullmatch(recorded[7:]):
            raise ValueError(f"GitHub records no SHA-256 for {name} in {repo} {tag}")
        url = f"https://github.com/{repo}/releases/download/{tag}/{name}"
        if asset.get("browser_download_url") != url:
            raise ValueError(f"{name} in {repo} {tag} is not served from {url}")
        return {"release": tag, "asset": asset_template, "url": url, "sha256": recorded[7:]}
    raise ValueError(f"{repo} {tag} has no asset {name}")


def choose(releases: list[dict[str, Any]], cooldown_days: int, now: dt.datetime) -> dict[str, Any]:
    """The newest full release published at least `cooldown_days` before `now`."""
    for release in releases:
        if release.get("draft") or release.get("prerelease"):
            continue
        published = dt.datetime.fromisoformat(release["published_at"].replace("Z", "+00:00"))
        if now - published >= dt.timedelta(days=cooldown_days):
            return release
    raise ValueError("no release is old enough yet")


def release_key(tag: str) -> tuple:
    """Order release tags: dates (20260927) and versions (v1.20.0) alike."""
    return tuple(int(part) for part in re.findall(r"\d+", tag))


def new_pins(pins: dict[str, Any], api: Api, now: dt.datetime) -> tuple[dict[str, Any], list[str]]:
    """The pins after every allowed move, and one line per move."""
    updated = copy.deepcopy(pins)
    moved = []
    for key, pin in pins.items():
        if key not in SOURCES:
            raise ValueError(f"{PIN_PATH} has an unknown entry {key!r}")
        source = SOURCES[key]
        releases = api(f"repos/{source['repo']}/releases?per_page=30")
        newest = choose(releases, source["cooldown_days"], now)
        candidate = asset_pin(newest, pin["asset"], source["repo"])
        if candidate["release"] == pin["release"]:
            if candidate["sha256"] != pin["sha256"]:
                raise ValueError(
                    f"{source['name']} {pin['release']} changed its SHA-256 since it was pinned")
            continue
        if release_key(candidate["release"]) < release_key(pin["release"]):
            # The pin was moved ahead by hand (for a fix the cooldown would
            # delay). Keep it; never propose going back.
            print(f"{source['name']} {pin['release']} is newer than the newest release "
                  f"the cooldown allows ({candidate['release']}); keeping it.")
            continue
        updated[key] = candidate
        moved.append(f"{source['name']} {pin['release']} -> {candidate['release']}")
    return updated, moved


def check_move(old: dict[str, Any], new: dict[str, Any]) -> None:
    """Refuse a proposal that changes anything a move does not allow."""
    if set(old) != set(new):
        raise ValueError("the proposal adds or removes a pin")
    for key in old:
        source = SOURCES[key]
        pin = new[key]
        if set(pin) != {"release", "asset", "url", "sha256"} or pin["asset"] != old[key]["asset"]:
            raise ValueError(f"the proposal changes the shape or asset of {key}")
        name = pin["asset"].format(release=pin["release"])
        if pin["url"] != f"https://github.com/{source['repo']}/releases/download/{pin['release']}/{name}":
            raise ValueError(f"the proposal points {key} somewhere else")
        if not SHA256.fullmatch(pin["sha256"]):
            raise ValueError(f"the proposal's {key} SHA-256 is malformed")
        if pin != old[key] and release_key(pin["release"]) <= release_key(old[key]["release"]):
            raise ValueError(f"the proposal does not move {key} forward")


def title_and_branch(old: dict[str, Any], new: dict[str, Any]) -> tuple[str, str]:
    parts, names = [], []
    for key in SOURCES:
        if key in new and new[key] != old[key]:
            parts.append(f"{SOURCES[key]['name']} {new[key]['release']}")
            names.append(f"{key.replace('_', '-')}-{new[key]['release']}")
    return "Update YARA pins: " + ", ".join(parts), "update-yara-rules/" + "-".join(names)


def body(moved: list[str], new: dict[str, Any]) -> str:
    lines = ["The weekly YARA update, proposed by `update-yara-rules.yml`.", ""]
    lines += [f"- {m}" for m in moved]
    lines += ["", "| Pin | Release | SHA-256 (GitHub's record of the asset) |", "|---|---|---|"]
    lines += [f"| {SOURCES[k]['name']} | [{p['release']}]({p['url']}) | `{p['sha256']}` |"
              for k, p in new.items()]
    lines += ["",
              "CI, Security and Malware scan run on this branch. If a new rule matches a file, the "
              "Malware scan fails until the match is fixed or accepted with a reason. Auto-merge "
              "is enabled only subject to the repository's required checks and review rules."]
    return "\n".join(lines) + "\n"


def prepare(root: Path, out: Path, base_sha: str, api: Api = gh_api,
            now: dt.datetime | None = None) -> bool:
    pins = json.loads((root / PIN_PATH).read_text(encoding="utf-8"))
    new, moved = new_pins(pins, api, now or dt.datetime.now(dt.timezone.utc))
    print("\n".join(moved) or "The YARA pins are current.")
    if not moved:
        return False
    title, branch = title_and_branch(pins, new)
    out.mkdir(parents=True, exist_ok=True)
    (out / "proposal.json").write_text(json.dumps(
        {"base_sha": base_sha, "title": title, "branch": branch, "pins": new}, indent=2) + "\n",
        encoding="utf-8")
    (out / "body.md").write_text(body(moved, new), encoding="utf-8")
    return True


def gh_auto_merge(url: str) -> None:
    result = subprocess.run(["gh", "pr", "merge", "--auto", "--squash", url],
                            capture_output=True, text=True, encoding="utf-8", timeout=120)
    if result.returncode != 0:
        raise ApiError(f"auto-merge for {url}: {(result.stdout + result.stderr).strip()}")


def propose(root: Path, inp: Path, repository: str, base_sha: str, api: Api = gh_api,
            auto_merge: Callable[[str], None] = gh_auto_merge) -> str | None:
    proposal = json.loads((inp / "proposal.json").read_text(encoding="utf-8"))
    if proposal["base_sha"] != base_sha or not re.fullmatch(r"[0-9a-f]{40}", base_sha):
        raise ValueError("the proposal was prepared from a different commit")
    if not re.fullmatch(r"[\w.-]+/[\w.-]+", repository):
        raise ValueError("unexpected repository name")
    old = json.loads((root / PIN_PATH).read_text(encoding="utf-8"))
    new = proposal["pins"]
    check_move(old, new)
    expected_title, branch = title_and_branch(old, new)
    if proposal["title"] != expected_title or proposal["branch"] != branch:
        raise ValueError("the proposal's title or branch does not match its pins")
    base = f"repos/{repository}"
    owner = repository.split("/")[0]
    pulls = api(f"{base}/pulls?state=all&head={owner}:{branch}&per_page=10")
    if pulls:
        pull = pulls[0]
        if pull["state"] != "open":
            print(f"{pull['html_url']} was closed; leaving that decision alone.")
            return None
    else:
        refs = api(f"{base}/git/matching-refs/heads/{branch}")
        if any(r["ref"] == f"refs/heads/{branch}" for r in refs):
            raise ValueError(f"branch {branch} exists without a pull request; look at it first")
        commit = api(f"{base}/git/commits/{base_sha}")
        tree = api(f"{base}/git/trees", "POST", {
            "base_tree": commit["tree"]["sha"],
            "tree": [{"path": PIN_PATH, "mode": "100644", "type": "blob",
                      "content": json.dumps(new, indent=2) + "\n"}]})
        created = api(f"{base}/git/commits", "POST", {
            "message": expected_title, "tree": tree["sha"], "parents": [base_sha]})
        api(f"{base}/git/refs", "POST", {"ref": f"refs/heads/{branch}", "sha": created["sha"]})
        pull = api(f"{base}/pulls", "POST", {
            "title": expected_title, "head": branch, "base": "main",
            "body": (inp / "body.md").read_text(encoding="utf-8")})
    # The App-created PR triggers checks normally. Do not dispatch duplicate
    # branch runs or restart the checks on a PR already awaiting review.
    # It merges itself once CI, Security and Malware scan pass; a failing scan
    # leaves it open. Without the "Allow auto-merge" setting it simply stays
    # open for review, which is not a failure of this run.
    try:
        auto_merge(pull["html_url"])
    except ApiError as exc:
        print(f"Auto-merge not turned on, so the pull request waits for review: {exc}")
    return pull["html_url"]


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("mode", choices=("prepare", "propose"))
    ap.add_argument("--out", type=Path)
    ap.add_argument("--in", dest="inp", type=Path)
    args = ap.parse_args(argv)
    root = Path.cwd()
    if args.mode == "prepare":
        changed = prepare(root, args.out, os.environ["GITHUB_SHA"])
        if os.environ.get("GITHUB_OUTPUT"):
            with open(os.environ["GITHUB_OUTPUT"], "a", encoding="utf-8") as fh:
                fh.write(f"changed={'true' if changed else 'false'}\n")
    else:
        url = propose(root, args.inp, os.environ["GITHUB_REPOSITORY"], os.environ["GITHUB_SHA"])
        if url:
            print(f"Proposed: {url}")
            if os.environ.get("GITHUB_STEP_SUMMARY"):
                with open(os.environ["GITHUB_STEP_SUMMARY"], "a", encoding="utf-8") as fh:
                    fh.write(f"Proposed the YARA update: {url}\n")
    return 0


if __name__ == "__main__":
    sys.exit(main())
