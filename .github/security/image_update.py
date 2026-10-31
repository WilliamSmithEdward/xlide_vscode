"""Propose the newest Docker Hub images the workflows may take.

This file is the same in every repository; the canonical copy lives in
WilliamSmithEdward/repo-standards (templates/images/image_update.py) and the
rescan checks each copy against it.

Dependabot cannot move these pins: GitHub has refused every one of its image
updates in these repositories ("invalid or unauthorized changes"). A pin is a
Dockerfile at .github/security/<tool>/Dockerfile or
.github/test/<service>/Dockerfile with one FROM line naming a Docker Hub image
by version tag and digest:

    FROM clamav/clamav:1.5.4@sha256:<multi-platform index digest>
    FROM golang:1.27.2-trixie@sha256:<multi-platform index digest>

A Dockerfile without such a line (another registry, or a digest with no tag)
is not followed here, and Dependabot keeps it.

Rules, from the standard:
- The highest X.Y.Z tag, with the pin's own suffix (-trixie above, or none),
  whose current digest Docker Hub has held for at least seven days. Other
  tags (latest, 1.5, -nonroot, _base) are never taken. A version older than
  the pin is never proposed; a pin moved ahead by hand is kept.
- Some publishers rebuild a version in place (ClamAV, the official images).
  A new digest for the pinned version is proposed under the same seven days.
  When nothing is old enough yet, the pin is kept.
- The digest is the multi-platform index digest. Docker Hub's API and the
  registry must report the same one.

Usage, from the two jobs of update-images.yml:

    python .github/security/image_update.py prepare --out DIR
    python .github/security/image_update.py propose --in DIR

`prepare` needs no write access. It writes one proposal per image to DIR and
sets `changed` in $GITHUB_OUTPUT. `propose` runs with write access and trusts
nothing in DIR it cannot check: a proposal may change only the FROM line of a
pin on the commit the workflow ran on, to a later version of the same image or
to the digest the registry now serves for the pinned version. It commits
through the GitHub API (no git credentials on disk) and opens each pull
request with a GitHub App token, which triggers the normal checks. It turns on
auto-merge, so a pull request merges itself once those pass and stays open if
any fails. It never reopens a pull request someone closed, and never
overwrites a branch it did not just create.
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

PIN_GLOBS = (".github/security/*/Dockerfile", ".github/test/*/Dockerfile")
COOLDOWN_DAYS = 7
MAX_PAGES = 10
HUB = "https://hub.docker.com/v2/"
TAGS_URL = HUB + "namespaces/{namespace}/repositories/{name}/tags?page_size=100&ordering=last_updated"
TOKEN_URL = "https://auth.docker.io/token?service=registry.docker.io&scope=repository:{path}:pull"
MANIFEST_URL = "https://registry-1.docker.io/v2/{path}/manifests/{tag}"
INDEX_TYPES = "application/vnd.oci.image.index.v1+json, application/vnd.docker.distribution.manifest.list.v2+json"
# A Docker Hub image (no registry host), a version tag with an optional
# suffix, and a digest; an optional build-stage name.
FROM = re.compile(
    r"^FROM (?P<image>[a-z0-9]+(?:[_-][a-z0-9]+)*(?:/[a-z0-9]+(?:[._-][a-z0-9]+)*)?)"
    r":(?P<version>\d+\.\d+\.\d+)(?P<suffix>-[a-z0-9][a-z0-9.-]*)?"
    r"@sha256:(?P<digest>[0-9a-f]{64})(?P<stage> +AS +[A-Za-z0-9_.-]+)?$", re.M)
SHA256 = re.compile(r"[0-9a-f]{64}")
TIMEOUT = 60

Api = Callable[..., Any]
Fetch = Callable[..., tuple[bytes, dict[str, str]]]


class ApiError(RuntimeError):
    pass


class Pin:
    """One followed FROM line."""

    def __init__(self, match: re.Match[str]):
        self.image, self.version = match["image"], match["version"]
        self.suffix, self.digest = match["suffix"] or "", match["digest"]
        self.stage = match["stage"] or ""

    @property
    def path(self) -> str:
        """The image's repository path on Docker Hub."""
        return self.image if "/" in self.image else f"library/{self.image}"

    @property
    def tag(self) -> str:
        return self.version + self.suffix

    def line(self, version: str, digest: str) -> str:
        return f"FROM {self.image}:{version}{self.suffix}@sha256:{digest}{self.stage}"


def gh_api(endpoint: str, method: str = "GET", payload: Any = None) -> Any:
    args = ["gh", "api", "--method", method, endpoint]
    if payload is not None:
        args += ["--input", "-"]
    result = subprocess.run(args, input=json.dumps(payload) if payload is not None else None,
                            capture_output=True, text=True, encoding="utf-8", timeout=120)
    if result.returncode != 0:
        raise ApiError(f"{method} {endpoint}: {(result.stdout + result.stderr).strip()}")
    return json.loads(result.stdout) if result.stdout.strip() else None


def http_get(url: str, headers: dict[str, str], head: bool = False) -> tuple[bytes, dict[str, str]]:
    """GET (or HEAD) `url` with curl, HTTPS only; the body and the response headers."""
    args = ["curl", "--fail", "--silent", "--show-error", "--proto", "=https",
            "--max-time", str(TIMEOUT), "--head" if head else "--dump-header", *([] if head else ["-"]), url]
    for name, value in headers.items():
        args[1:1] = ["--header", f"{name}: {value}"]
    result = subprocess.run(args, capture_output=True, timeout=TIMEOUT + 10)
    if result.returncode != 0:
        raise ApiError(f"GET {url}: {result.stderr.decode('utf-8', 'replace').strip()}")
    head_part, _, payload = result.stdout.partition(b"\r\n\r\n")
    fields = {}
    for line in head_part.decode("iso-8859-1").splitlines()[1:]:
        name, _, value = line.partition(":")
        fields[name.strip().lower()] = value.strip()
    return payload, fields


def version_key(version: str) -> tuple[int, ...]:
    return tuple(int(part) for part in version.split("."))


def read_pin(text: str) -> Pin | None:
    """The file's followed FROM line, or None when it has none."""
    found = list(FROM.finditer(text))
    if len(found) > 1:
        raise ValueError("more than one followed FROM line; split the Dockerfile")
    return Pin(found[0]) if found else None


def pin_files(root: Path) -> list[str]:
    return sorted(str(p.relative_to(root)).replace("\\", "/")
                  for pattern in PIN_GLOBS for p in root.glob(pattern) if p.is_file())


def hub_tags(pin: Pin, fetch: Fetch) -> list[dict[str, Any]]:
    namespace, name = pin.path.split("/")
    url = TAGS_URL.format(namespace=namespace, name=name)
    if pin.suffix:
        url += f"&name={pin.suffix}"
    tags: list[dict[str, Any]] = []
    for _ in range(MAX_PAGES):
        page = json.loads(fetch(url, {})[0])
        tags += page["results"]
        url = page.get("next")
        if not url:
            break
        if not url.startswith(HUB):
            raise ValueError(f"Docker Hub's next page is not on {HUB}")
    return tags


def registry_digest(pin: Pin, tag: str, fetch: Fetch) -> str:
    """The digest the registry serves for `tag`, as `docker pull` would get it.
    A HEAD request, which Docker Hub does not count against the pull limit."""
    body, _ = fetch(TOKEN_URL.format(path=pin.path), {})
    token = json.loads(body)["token"]
    _, headers = fetch(MANIFEST_URL.format(path=pin.path, tag=tag),
                       {"Authorization": f"Bearer {token}", "Accept": INDEX_TYPES}, head=True)
    digest = headers.get("docker-content-digest", "")
    if not digest.startswith("sha256:") or not SHA256.fullmatch(digest[7:]):
        raise ValueError(f"the registry reports no SHA-256 for {pin.image}:{tag}")
    return digest[7:]


def choose(pin: Pin, tags: list[dict[str, Any]], now: dt.datetime) -> tuple[str, str] | None:
    """The highest version, with the pin's suffix, whose digest is a week old."""
    wanted = re.compile(r"(\d+\.\d+\.\d+)" + re.escape(pin.suffix))
    best = None
    for tag in tags:
        match = wanted.fullmatch(tag.get("name", ""))
        if not match or not tag.get("tag_last_pushed"):
            continue
        pushed = dt.datetime.fromisoformat(tag["tag_last_pushed"].replace("Z", "+00:00"))
        if now - pushed < dt.timedelta(days=COOLDOWN_DAYS):
            continue
        if best is None or version_key(match[1]) > version_key(best[0]):
            best = (match[1], tag)
    if best is None:
        return None
    digest = best[1].get("digest") or ""
    if not digest.startswith("sha256:") or not SHA256.fullmatch(digest[7:]):
        raise ValueError(f"Docker Hub records no SHA-256 for {pin.image}:{best[1]['name']}")
    return best[0], digest[7:]


def new_pin(pin: Pin, fetch: Fetch, now: dt.datetime) -> tuple[str, str, str | None]:
    """The version and digest to pin, and a line describing the move (None for none)."""
    chosen = choose(pin, hub_tags(pin, fetch), now)
    if chosen is None:
        print(f"{pin.image}: no {pin.suffix or 'plain'} version is a week old yet; keeping {pin.tag}.")
        return pin.version, pin.digest, None
    version, digest = chosen
    if version_key(version) < version_key(pin.version):
        # The pin was moved ahead by hand. Keep it; never propose going back.
        print(f"{pin.image}:{pin.tag} is newer than the newest version the cooldown allows "
              f"({version}{pin.suffix}); keeping it.")
        return pin.version, pin.digest, None
    if version == pin.version and digest == pin.digest:
        return pin.version, pin.digest, None
    if registry_digest(pin, version + pin.suffix, fetch) != digest:
        raise ValueError(f"Docker Hub and the registry disagree on the digest of "
                         f"{pin.image}:{version}{pin.suffix}")
    if version == pin.version:
        return version, digest, f"{pin.image}:{pin.tag} was rebuilt; {pin.digest[:12]} -> {digest[:12]}"
    return version, digest, f"{pin.image} {pin.tag} -> {version}{pin.suffix}"


def replace_pin(text: str, pin: Pin, version: str, digest: str) -> str:
    return FROM.sub(lambda _: pin.line(version, digest), text)


def check_move(old: str, new: str) -> Pin:
    """Refuse a proposal that changes anything but a forward move of the FROM line."""
    before, after = read_pin(old), read_pin(new)
    if before is None or after is None:
        raise ValueError("the proposal has no followed FROM line")
    if (after.image, after.suffix, after.stage) != (before.image, before.suffix, before.stage):
        raise ValueError("the proposal names another image")
    if replace_pin(old, before, after.version, after.digest) != new:
        raise ValueError("the proposal changes more than the FROM line")
    if version_key(after.version) < version_key(before.version) or (
            after.version == before.version and after.digest == before.digest):
        raise ValueError("the proposal does not move the image forward")
    return after


def title_and_branch(path: str, pin: Pin) -> tuple[str, str]:
    service = path.split("/")[2]
    return (f"Update the {pin.image} image to {pin.tag} ({pin.digest[:12]})",
            f"update-image/{service}/{pin.tag}-{pin.digest[:12]}")


def body(path: str, moved: str, pin: Pin) -> str:
    return "\n".join([
        "The weekly image update, proposed by `update-images.yml`.",
        "",
        f"- {moved}",
        f"- `{path}`: `{pin.image}:{pin.tag}@sha256:{pin.digest}`, the multi-platform index "
        "digest Docker Hub and the registry both report.",
        "",
        "CI, Security and Malware scan run on this branch. Auto-merge is enabled only subject "
        "to the repository's required checks and review rules.",
    ]) + "\n"


def prepare(root: Path, out: Path, base_sha: str, fetch: Fetch = http_get,
            now: dt.datetime | None = None) -> bool:
    now = now or dt.datetime.now(dt.timezone.utc)
    proposals = []
    for path in pin_files(root):
        text = (root / path).read_text(encoding="utf-8")
        pin = read_pin(text)
        if pin is None:
            print(f"{path}: no Docker Hub version pin; Dependabot keeps it.")
            continue
        version, digest, moved = new_pin(pin, fetch, now)
        print(f"{path}: {moved or f'{pin.image}:{pin.tag} is current.'}")
        if not moved:
            continue
        dockerfile = replace_pin(text, pin, version, digest)
        after = read_pin(dockerfile)
        title, branch = title_and_branch(path, after)
        proposals.append({"path": path, "title": title, "branch": branch,
                          "dockerfile": dockerfile, "body": body(path, moved, after)})
    if not proposals:
        print("The images are current.")
        return False
    out.mkdir(parents=True, exist_ok=True)
    (out / "proposals.json").write_text(json.dumps(
        {"base_sha": base_sha, "proposals": proposals}, indent=2) + "\n", encoding="utf-8")
    return True


def gh_auto_merge(url: str) -> None:
    result = subprocess.run(["gh", "pr", "merge", "--auto", "--squash", url],
                            capture_output=True, text=True, encoding="utf-8", timeout=120)
    if result.returncode != 0:
        raise ApiError(f"auto-merge for {url}: {(result.stdout + result.stderr).strip()}")


def propose_one(root: Path, proposal: dict[str, Any], repository: str, base_sha: str,
                api: Api, auto_merge: Callable[[str], None], fetch: Fetch) -> str | None:
    path = proposal["path"]
    if path not in pin_files(root):
        raise ValueError(f"{path} is not an image pin")
    new = proposal["dockerfile"]
    after = check_move((root / path).read_text(encoding="utf-8"), new)
    if registry_digest(after, after.tag, fetch) != after.digest:
        raise ValueError(f"the registry does not serve {after.digest[:12]} for {after.image}:{after.tag}")
    title, branch = title_and_branch(path, after)
    if proposal["title"] != title or proposal["branch"] != branch:
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
            "tree": [{"path": path, "mode": "100644", "type": "blob", "content": new}]})
        created = api(f"{base}/git/commits", "POST", {
            "message": title, "tree": tree["sha"], "parents": [base_sha]})
        api(f"{base}/git/refs", "POST", {"ref": f"refs/heads/{branch}", "sha": created["sha"]})
        pull = api(f"{base}/pulls", "POST", {
            "title": title, "head": branch, "base": "main", "body": proposal["body"]})
    # It merges itself once CI, Security and Malware scan pass; a failing scan
    # leaves it open. Without the "Allow auto-merge" setting it simply stays
    # open for review, which is not a failure of this run.
    try:
        auto_merge(pull["html_url"])
    except ApiError as exc:
        print(f"Auto-merge not turned on, so the pull request waits for review: {exc}")
    return pull["html_url"]


def propose(root: Path, inp: Path, repository: str, base_sha: str, api: Api = gh_api,
            auto_merge: Callable[[str], None] = gh_auto_merge, fetch: Fetch = http_get) -> list[str]:
    """Opens a pull request for each proposal; returns their URLs. Every
    proposal is tried, and any that failed are raised together at the end."""
    data = json.loads((inp / "proposals.json").read_text(encoding="utf-8"))
    if data["base_sha"] != base_sha or not re.fullmatch(r"[0-9a-f]{40}", base_sha):
        raise ValueError("the proposals were prepared from a different commit")
    if not re.fullmatch(r"[\w.-]+/[\w.-]+", repository):
        raise ValueError("unexpected repository name")
    urls, errors = [], []
    for proposal in data["proposals"]:
        try:
            url = propose_one(root, proposal, repository, base_sha, api, auto_merge, fetch)
        except (ApiError, ValueError, KeyError) as exc:
            errors.append(f"{proposal.get('path')}: {exc}")
            continue
        if url:
            urls.append(url)
    if errors:
        raise ValueError("; ".join(errors))
    return urls


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
        for url in propose(root, args.inp, os.environ["GITHUB_REPOSITORY"], os.environ["GITHUB_SHA"]):
            print(f"Proposed: {url}")
            if os.environ.get("GITHUB_STEP_SUMMARY"):
                with open(os.environ["GITHUB_STEP_SUMMARY"], "a", encoding="utf-8") as fh:
                    fh.write(f"Proposed an image update: {url}\n")
    return 0


if __name__ == "__main__":
    sys.exit(main())
