// `--changed`: the fixtures a branch adds or changes (issue #1498).
//
// The base is the cases file at the commit where the branch left
// `upstream/main`, else `origin/main`, or the ref `--changed-from` names.
// The branch side is the file on disk, so uncommitted edits count. A fixture
// is changed when its id is new or any field differs from the base fixture of
// that id; removed fixtures do not run. When the base cannot be worked out,
// the selection refuses rather than falling back to every fixture.

import { spawnSync } from 'node:child_process';
import fs from 'node:fs';
import path from 'node:path';

const DEFAULT_BASES = ['upstream/main', 'origin/main'];
const BOM = new RegExp(`^${String.fromCharCode(0xfeff)}`);

/** Runs git with fixed arguments and no shell; `{ ok, out, why }`. */
function git(cwd, args) {
    const run = spawnSync('git', ['-C', cwd, ...args], { encoding: 'buffer', maxBuffer: 1 << 30, windowsHide: true });
    if (run.error) {
        return { ok: false, why: run.error.code === 'ENOENT' ? 'git is not installed or not on PATH' : run.error.message };
    }
    const err = run.stderr.toString('utf8').trim();
    return run.status === 0
        ? { ok: true, out: run.stdout }
        : { ok: false, why: err || `git ${args[0]} exited with code ${run.status}` };
}

/** Equal as JSON values, whatever order their keys are in. */
function sameValue(a, b) {
    if (a === b) { return true; }
    if (typeof a !== 'object' || typeof b !== 'object' || a === null || b === null || Array.isArray(a) !== Array.isArray(b)) {
        return false;
    }
    if (Array.isArray(a)) {
        return a.length === b.length && a.every((item, i) => sameValue(item, b[i]));
    }
    const keys = Object.keys(a);
    return keys.length === Object.keys(b).length && keys.every((key) => Object.hasOwn(b, key) && sameValue(a[key], b[key]));
}

/**
 * The fixtures of `cases` (the document on disk at `casesFile`) that are new
 * or changed against the base. Returns `{ cases, base }`, or `{ refusal }`
 * saying why the base cannot be established.
 */
export function changedCases(casesFile, cases, from) {
    if (from !== undefined && from.startsWith('-')) {
        return { refusal: `--changed-from expects a git ref, not '${from}'` };
    }
    const absolute = path.resolve(casesFile);
    const dir = path.dirname(absolute);
    const top = git(dir, ['rev-parse', '--show-toplevel']);
    if (!top.ok) {
        return { refusal: `${casesFile} is not in a git work tree (${top.why})` };
    }
    const root = top.out.toString('utf8').trim();
    const shallow = git(dir, ['rev-parse', '--is-shallow-repository']);
    if (!shallow.ok || shallow.out.toString('utf8').trim() !== 'false') {
        return { refusal: `the clone at ${root} is shallow, so the base commit may be missing; fetch its full history (git fetch --unshallow)` };
    }
    // The path as git names it, relative to the work tree and with forward
    // slashes; realpath, because the work tree root git prints is resolved.
    const relative = path.relative(fs.realpathSync.native(root), fs.realpathSync.native(absolute)).split(path.sep).join('/');
    const tracked = git(root, ['ls-files', '--error-unmatch', '--', relative]);
    if (!tracked.ok) {
        return { refusal: `${casesFile} is not tracked by git, so it has no base version` };
    }
    let ref;
    if (from !== undefined) {
        ref = from;
    } else {
        ref = DEFAULT_BASES.find((candidate) => git(root, ['rev-parse', '--verify', '--quiet', '--end-of-options', `${candidate}^{commit}`]).ok);
        if (!ref) {
            return { refusal: `neither ${DEFAULT_BASES.join(' nor ')} exists; name the base with --changed-from <ref>` };
        }
    }
    const resolved = git(root, ['rev-parse', '--verify', '--quiet', '--end-of-options', `${ref}^{commit}`]);
    if (!resolved.ok) {
        return { refusal: `the base ref '${ref}' does not name a commit` };
    }
    const fork = git(root, ['merge-base', 'HEAD', resolved.out.toString('utf8').trim()]);
    if (!fork.ok) {
        return { refusal: `HEAD and '${ref}' share no history, so there is no base` };
    }
    const base = fork.out.toString('utf8').trim();
    const shown = git(root, ['show', `${base}:${relative}`]);
    if (!shown.ok) {
        return { refusal: `${relative} does not exist at the base ${base.slice(0, 12)} (where HEAD left '${ref}')` };
    }
    let baseCases;
    try {
        // utf-8-sig, as loadCasesDocument reads the file on disk.
        baseCases = JSON.parse(shown.out.toString('utf8').replace(BOM, '')).cases;
    } catch (err) {
        return { refusal: `${relative} at the base ${base.slice(0, 12)} cannot be read: ${err.message}` };
    }
    if (!Array.isArray(baseCases)) {
        return { refusal: `${relative} at the base ${base.slice(0, 12)} has no cases array` };
    }
    const byId = new Map(baseCases.map((kase) => [kase.id, kase]));
    return {
        base: { ref, commit: base },
        cases: cases.filter((kase) => !byId.has(kase.id) || !sameValue(byId.get(kase.id), kase)),
    };
}
