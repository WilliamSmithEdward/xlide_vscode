"""Propose the newest Semgrep image the Security workflow may take.

This file is the same in every repository; the canonical copy lives in
WilliamSmithEdward/repo-standards (templates/semgrep/semgrep_update.py) and
the rescan checks each copy against it.

Dependabot cannot move this pin: GitHub has refused every one of its Semgrep
image updates since 2026-10-20 ("invalid or unauthorized changes"). The pin is
the one line of .github/security/semgrep/Dockerfile that starts with FROM:

    FROM semgrep/semgrep:1.182.1@sha256:<multi-platform index digest>

Rules, from the standard:
- The highest X.Y.Z tag Docker Hub has held for at least seven days. Other
  tags (latest, canary, -nonroot) are never taken. A version older than the
  pin is never proposed; a pin moved ahead of the cooldown by hand is kept.
- The digest is the multi-platform index digest. Docker Hub's API and the
  registry must report the same one. A pinned tag whose digest changed is an
  error, not an update.

Usage, from the two jobs of update-semgrep-image.yml:

    python .github/security/semgrep_update.py prepare --out DIR
    python .github/security/semgrep_update.py propose --in DIR

`prepare` needs no write access. It writes the proposal to DIR and sets
`changed` in $GITHUB_OUTPUT. `propose` runs with write access and trusts
nothing in DIR it cannot check: the proposal may change only the FROM line,
to a later version of the same image, from the commit the workflow ran on.
It commits through the GitHub API (no git credentials on disk) and opens the
pull request with a GitHub App token, which triggers the normal checks. It
turns on auto-merge, so the pull request merges itself once those pass and
stays open if any fails. It never reopens a pull request someone closed,
and never overwrites a branch it did not just create.
"""
from __future__ import annotations

import argparse
import datetime as dt
import json
import os
import re
import subprocess
import sys
from collections.abc import Callable
from pathlib import Path
from typing import Any

PIN_PATH = ".github/security/semgrep/Dockerfile"
IMAGE = "semgrep/semgrep"
COOLDOWN_DAYS = 7
TAGS_URL = (f"https://hub.docker.com/v2/namespaces/{IMAGE.split('/')[0]}/repositories/"
            f"{IMAGE.split('/')[1]}/tags?page_size=100&ordering=last_updated")
TOKEN_URL = f"https://auth.docker.io/token?service=registry.docker.io&scope=repository:{IMAGE}:pull"
MANIFEST_URL = "https://registry-1.docker.io/v2/" + IMAGE + "/manifests/{tag}"
INDEX_TYPES = "application/vnd.oci.image.index.v1+json, application/vnd.docker.distribution.manifest.list.v2+json"
VERSION = re.compile(r"\d+\.\d+\.\d+")
FROM = re.compile(r"^FROM " + re.escape(IMAGE) + r":(\d+\.\d+\.\d+)@sha256:([0-9a-f]{64})$", re.M)
SHA256 = re.compile(r"[0-9a-f]{64}")
TIMEOUT = 60

Api = Callable[..., Any]
Fetch = Callable[[str, dict[str, str]], tuple[bytes, dict[str, str]]]


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


def http_get(url: str, headers: dict[str, str]) -> tuple[bytes, dict[str, str]]:
    """GET `url` with curl, HTTPS only; the body and the response headers."""
    args = ["curl", "--fail", "--silent", "--show-error", "--proto", "=https",
            "--max-time", str(TIMEOUT), "--dump-header", "-", url]
    for name, value in headers.items():
        args[1:1] = ["--header", f"{name}: {value}"]
    result = subprocess.run(args, capture_output=True, timeout=TIMEOUT + 10)
    if result.returncode != 0:
        raise ApiError(f"GET {url}: {result.stderr.decode('utf-8', 'replace').strip()}")
    head, _, payload = result.stdout.partition(b"\r\n\r\n")
    fields = {}
    for line in head.decode("iso-8859-1").splitlines()[1:]:
        name, _, value = line.partition(":")
        fields[name.strip().lower()] = value.strip()
    return payload, fields


def version_key(tag: str) -> tuple[int, ...]:
    return tuple(int(part) for part in tag.split("."))


def read_pin(text: str) -> tuple[str, str]:
    found = FROM.findall(text)
    if len(found) != 1:
        raise ValueError(f"{PIN_PATH} needs exactly one 'FROM {IMAGE}:X.Y.Z@sha256:...' line")
    return found[0]


def registry_digest(tag: str, fetch: Fetch) -> str:
    """The digest the registry serves for `tag`, as `docker pull` would get it."""
    body, _ = fetch(TOKEN_URL, {})
    token = json.loads(body)["token"]
    _, headers = fetch(MANIFEST_URL.format(tag=tag),
                       {"Authorization": f"Bearer {token}", "Accept": INDEX_TYPES})
    digest = headers.get("docker-content-digest", "")
    if not digest.startswith("sha256:") or not SHA256.fullmatch(digest[7:]):
        raise ValueError(f"the registry reports no SHA-256 for {IMAGE}:{tag}")
    return digest[7:]


def choose(tags: list[dict[str, Any]], now: dt.datetime) -> tuple[str, str]:
    """The highest X.Y.Z tag pushed at least COOLDOWN_DAYS before `now`, with its digest."""
    best = None
    for tag in tags:
        name = tag.get("name", "")
        if not VERSION.fullmatch(name):
            continue
        pushed = dt.datetime.fromisoformat(tag["tag_last_pushed"].replace("Z", "+00:00"))
        if now - pushed < dt.timedelta(days=COOLDOWN_DAYS):
            continue
        if best is None or version_key(name) > version_key(best["name"]):
            best = tag
    if best is None:
        raise ValueError("no Semgrep version is old enough yet")
    digest = best.get("digest") or ""
    if not digest.startswith("sha256:") or not SHA256.fullmatch(digest[7:]):
        raise ValueError(f"Docker Hub records no SHA-256 for {IMAGE}:{best['name']}")
    return best["name"], digest[7:]


def new_pin(text: str, fetch: Fetch, now: dt.datetime) -> tuple[str, str, str | None]:
    """The version and digest to pin, and a line describing the move (None for none)."""
    version, digest = read_pin(text)
    body, _ = fetch(TAGS_URL, {})
    candidate, candidate_digest = choose(json.loads(body)["results"], now)
    if candidate == version:
        if registry_digest(version, fetch) != digest:
            raise ValueError(f"Semgrep {version} changed its digest since it was pinned")
        return version, digest, None
    if version_key(candidate) < version_key(version):
        # The pin was moved ahead by hand. Keep it; never propose going back.
        print(f"Semgrep {version} is newer than the newest version the cooldown allows "
              f"({candidate}); keeping it.")
        return version, digest, None
    if registry_digest(candidate, fetch) != candidate_digest:
        raise ValueError(f"Docker Hub and the registry disagree on the digest of {IMAGE}:{candidate}")
    return candidate, candidate_digest, f"Semgrep {version} -> {candidate}"


def replace_pin(text: str, version: str, digest: str) -> str:
    return FROM.sub(f"FROM {IMAGE}:{version}@sha256:{digest}", text)


def check_move(old: str, new: str) -> None:
    """Refuse a proposal that changes anything but a forward move of the FROM line."""
    old_version, _ = read_pin(old)
    new_version, new_digest = read_pin(new)
    if replace_pin(old, new_version, new_digest) != new:
        raise ValueError("the proposal changes more than the FROM line")
    if version_key(new_version) <= version_key(old_version):
        raise ValueError("the proposal does not move Semgrep forward")


def title_and_branch(version: str) -> tuple[str, str]:
    return f"Update the Semgrep image to {version}", f"update-semgrep-image/{version}"


def body(moved: str, version: str, digest: str) -> str:
    return "\n".join([
        "The weekly Semgrep image update, proposed by `update-semgrep-image.yml`.",
        "",
        f"- {moved}",
        f"- `{IMAGE}:{version}@sha256:{digest}`, the multi-platform index digest Docker Hub "
        "and the registry both report.",
        "",
        "CI, Security and Malware scan run on this branch. If the new Semgrep finds something, "
        "Security fails until the finding is fixed or accepted with a reason. Auto-merge is "
        "enabled only subject to the repository's required checks and review rules.",
    ]) + "\n"


def prepare(root: Path, out: Path, base_sha: str, fetch: Fetch = http_get,
            now: dt.datetime | None = None) -> bool:
    text = (root / PIN_PATH).read_text(encoding="utf-8")
    version, digest, moved = new_pin(text, fetch, now or dt.datetime.now(dt.timezone.utc))
    print(moved or "The Semgrep image is current.")
    if not moved:
        return False
    title, branch = title_and_branch(version)
    out.mkdir(parents=True, exist_ok=True)
    (out / "proposal.json").write_text(json.dumps(
        {"base_sha": base_sha, "title": title, "branch": branch,
         "dockerfile": replace_pin(text, version, digest)}, indent=2) + "\n", encoding="utf-8")
    (out / "body.md").write_text(body(moved, version, digest), encoding="utf-8")
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
    old = (root / PIN_PATH).read_text(encoding="utf-8")
    new = proposal["dockerfile"]
    check_move(old, new)
    expected_title, branch = title_and_branch(read_pin(new)[0])
    if proposal["title"] != expected_title or proposal["branch"] != branch:
        raise ValueError("the proposal's title or branch does not match its pin")
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
            "tree": [{"path": PIN_PATH, "mode": "100644", "type": "blob", "content": new}]})
        created = api(f"{base}/git/commits", "POST", {
            "message": expected_title, "tree": tree["sha"], "parents": [base_sha]})
        api(f"{base}/git/refs", "POST", {"ref": f"refs/heads/{branch}", "sha": created["sha"]})
        pull = api(f"{base}/pulls", "POST", {
            "title": expected_title, "head": branch, "base": "main",
            "body": (inp / "body.md").read_text(encoding="utf-8")})
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
                    fh.write(f"Proposed the Semgrep image update: {url}\n")
    return 0


if __name__ == "__main__":
    sys.exit(main())
