"""Relock what a Dependabot pull request changed, the way the locks are built.

The same in every repository; the canonical copy lives in
WilliamSmithEdward/repo-standards (templates/relock). Dependabot recompiles
uv's hash-locked requirement files with its own pip-compile, which drops
setuptools, pip and the platform markers, and rewrites a multi-framework
packages.lock.json for one target framework only. Either way CI fails and
the update never merges. relock-dependabot.yml runs this in two jobs:

  relock <base-ref> <out-dir>
      On the pull request's branch, with no secret. For each changed lock
      file that records a `uv pip compile` command on its second line, the
      file starts again from the base branch's copy and that command runs
      with `--upgrade-package name==version` for every package whose pinned
      version the pull request moved, so the moves Dependabot meant are kept
      and nothing else changes. If any packages.lock.json, project file or
      Directory.Packages.props changed, `dotnet restore --force-evaluate`
      runs on the repository's solution. Every file that then differs is
      copied into <out-dir>, with manifest.json listing them.

  push <in-dir> <head-ref> <head-sha>
      From the base branch's copy of this script, with the Repo Bot App's
      token in GH_TOKEN. Accepts only files that already exist at <head-sha>
      and are lock files (a packages.lock.json, or a requirement file whose
      second line is a `uv pip compile` command), and commits them onto
      <head-sha> through the API, moving <head-ref> only if it has not moved
      since. It never runs the pull request's code.
"""
import base64
import json
import re
import shlex
import shutil
import subprocess
import sys
from pathlib import Path, PurePosixPath

PIN = re.compile(r"^([A-Za-z0-9][A-Za-z0-9._-]*)==([^\s;\\]+)", re.M)
DOTNET_INPUTS = ("packages.lock.json", ".csproj", "Directory.Packages.props")
MESSAGE = ("Relock with the locks' own commands\n\n"
           "Dependabot's rebuild dropped what uv or dotnet restore keeps; relock.py\n"
           "rebuilt each changed lock from the base branch, keeping the versions this\n"
           "pull request moves.")


def git(*args):
    return subprocess.run(["git", *args], check=True, capture_output=True, text=True).stdout


def pins(text):
    return {m.group(1).lower().replace("_", "-"): m.group(2) for m in PIN.finditer(text)}


def uv_command(text):
    lines = text.splitlines()
    if len(lines) < 2:
        return None
    command = lines[1].lstrip("#").strip()
    return shlex.split(command) if command.startswith("uv pip compile ") else None


def is_lock(path, text):
    return path.endswith("packages.lock.json") or (path.endswith(".txt") and uv_command(text) is not None)


def relock_requirements(base, path):
    head = Path(path).read_text(encoding="utf-8")
    command = uv_command(head)
    if command is None:
        return
    try:
        old = git("show", f"{base}:{path}")
    except subprocess.CalledProcessError:
        return  # a new lock file: nothing to start again from
    moved = [f"{name}=={version}" for name, version in pins(head).items()
             if pins(old).get(name) not in (None, version)]
    Path(path).write_text(old, encoding="utf-8", newline="\n")
    upgrades = [arg for spec in moved for arg in ("--upgrade-package", spec)]
    subprocess.run(command + upgrades + ["--quiet"], cwd=Path(path).parent, check=True)
    print(f"relocked {path}" + (f" ({', '.join(moved)})" if moved else ""))


def relock(base, out):
    changed = git("diff", "--name-only", f"{base}...HEAD").split()
    for path in changed:
        if path.endswith(".txt") and Path(path).is_file():
            relock_requirements(base, path)
    if any(path.endswith(DOTNET_INPUTS) for path in changed):
        found = sorted(Path(".").glob("*.sln")) + sorted(Path(".").glob("*.slnx"))
        if found:
            subprocess.run(["dotnet", "restore", str(found[0]), "--force-evaluate"], check=True)
            print(f"restored {found[0]} with --force-evaluate")
    differ = git("diff", "--name-only").split()
    out = Path(out)
    out.mkdir(parents=True, exist_ok=True)
    for path in differ:
        (out / "files" / path).parent.mkdir(parents=True, exist_ok=True)
        shutil.copyfile(path, out / "files" / path)
    (out / "manifest.json").write_text(json.dumps(differ), encoding="utf-8")
    print(f"{len(differ)} file(s) to push")


def api(method, url, body=None):
    """One call to the GitHub REST API for this repository, through the gh
    CLI the runner provides. gh reads the App's token from GH_TOKEN and
    fills {owner}/{repo} from the checkout's remote."""
    command = ["gh", "api", "--method", method, "repos/{owner}/{repo}/" + url]
    if body is not None:
        command += ["--input", "-"]
    result = subprocess.run(command, input=None if body is None else json.dumps(body),
                            capture_output=True, text=True, timeout=120)
    if result.returncode != 0:
        sys.exit(f"{method} {url} failed: {result.stderr.strip()[:300]}")
    return json.loads(result.stdout or "null")


def push(src, head_ref, head_sha):
    if not re.fullmatch(r"[0-9a-f]{40}", head_sha) or not re.fullmatch(r"[\w./-]+", head_ref) \
            or ".." in head_ref:
        sys.exit("refused: the head branch or commit is not well formed")
    src = Path(src)
    paths = json.loads((src / "manifest.json").read_text(encoding="utf-8"))
    tree = []
    for path in paths:
        posix = PurePosixPath(path)
        if posix.is_absolute() or ".." in posix.parts:
            sys.exit(f"refused {path}: not a plain path inside the repository")
        current = api("GET", f"contents/{path}?ref={head_sha}")
        text = base64.b64decode(current["content"]).decode("utf-8")
        if not is_lock(path, text):
            sys.exit(f"refused {path}: not a lock file at {head_sha}")
        content = (src / "files" / path).read_text(encoding="utf-8")
        blob = api("POST", "git/blobs", {"content": content, "encoding": "utf-8"})
        tree.append({"path": path, "mode": "100644", "type": "blob", "sha": blob["sha"]})
    if not tree:
        print("nothing to push")
        return
    base_tree = api("GET", f"git/commits/{head_sha}")["tree"]["sha"]
    new_tree = api("POST", "git/trees", {"base_tree": base_tree, "tree": tree})
    commit = api("POST", "git/commits", {"message": MESSAGE, "tree": new_tree["sha"], "parents": [head_sha]})
    api("PATCH", f"git/refs/heads/{head_ref}", {"sha": commit["sha"], "force": False})
    print(f"pushed {commit['sha']} to {head_ref}: {', '.join(paths)}")


if __name__ == "__main__":
    if sys.argv[1:2] == ["relock"] and len(sys.argv) == 4:
        relock(sys.argv[2], sys.argv[3])
    elif sys.argv[1:2] == ["push"] and len(sys.argv) == 5:
        push(sys.argv[2], sys.argv[3], sys.argv[4])
    else:
        sys.exit(__doc__)
