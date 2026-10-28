// `--changed` measures only the fixtures a branch adds or changes (issue
// #1498). Every case runs against real git in a temporary repository.

import { spawnSync } from 'node:child_process';
import * as fs from 'node:fs';
import * as os from 'node:os';
import * as path from 'node:path';
import { afterEach, describe, expect, it } from 'vitest';
import { changedCases } from '../../syntax_corpus/oracle/oracleChangedCases.mjs';

interface Selection { cases?: { id: string }[]; base?: { ref: string; commit: string }; refusal?: string }
const select = changedCases as (file: string, cases: object[], from?: string) => Selection;

const RUNNER = path.join(process.cwd(), 'syntax_corpus/oracle/run_vbe_oracle.mjs');
const made: string[] = [];

afterEach(() => {
	for (const dir of made.splice(0)) {
		fs.rmSync(dir, { recursive: true, force: true });
	}
});

function git(cwd: string, ...args: string[]): string {
	const run = spawnSync('git', ['-c', 'user.name=Oracle Test', '-c', 'user.email=oracle@example.invalid', '-c', 'commit.gpgsign=false', ...args], { cwd, encoding: 'utf8' });
	if (run.status !== 0) {
		throw new Error(`git ${args.join(' ')}: ${run.stderr}`);
	}
	return run.stdout.trim();
}

const kase = (id: string, source = 'Sub Main()\r\nEnd Sub\r\n'): object => ({ id, expected: 'accepted', mode: 'compile', modules: [{ name: 'Module1', type: 'standard', source }] });

/** A repository whose `origin/main` holds `cases` in `folder/file`, checked out on a branch. */
function repository(cases: object[], folder = 'oracle', file = 'cases.json'): { root: string; cases: string } {
	const root = fs.mkdtempSync(path.join(os.tmpdir(), 'xlide-changed-'));
	made.push(root);
	git(root, 'init', '-q', '-b', 'main');
	const target = path.join(root, folder, file);
	fs.mkdirSync(path.dirname(target), { recursive: true });
	write(target, cases);
	git(root, 'add', '-A');
	git(root, 'commit', '-q', '-m', 'base');
	git(root, 'update-ref', 'refs/remotes/origin/main', 'HEAD');
	git(root, 'checkout', '-q', '-b', 'feature');
	return { root, cases: target };
}

function write(file: string, cases: object[]): void {
	fs.writeFileSync(file, JSON.stringify({ version: 1, cases }, null, 2));
}

const ids = (selection: Selection): string[] => (selection.cases ?? []).map((c) => c.id);

describe('the fixtures --changed selects', () => {
	it('takes a new id and a changed field, uncommitted, and leaves the rest', () => {
		const repo = repository([kase('kept'), kase('edited'), kase('removed')]);
		const now = [kase('kept'), kase('edited', 'Sub Main()\r\n    x = 1\r\nEnd Sub\r\n'), kase('added')];
		write(repo.cases, now);
		const selection = select(repo.cases, now);
		expect(ids(selection)).toEqual(['edited', 'added']);
		expect(selection.base?.ref).toBe('origin/main');
		expect(selection.base?.commit).toBe(git(repo.root, 'rev-parse', 'origin/main'));
	});

	it('does not count a field order change', () => {
		const repo = repository([kase('kept')]);
		const reordered = [{ modules: [{ source: 'Sub Main()\r\nEnd Sub\r\n', type: 'standard', name: 'Module1' }], mode: 'compile', expected: 'accepted', id: 'kept' }];
		expect(ids(select(repo.cases, reordered))).toEqual([]);
	});

	it('compares with where the branch left main, not with main\'s later commits', () => {
		const repo = repository([kase('old')]);
		const branch = [kase('old'), kase('mine')];
		write(repo.cases, branch);
		git(repo.root, 'commit', '-qam', 'branch fixture');
		// Main moves on: it adds a fixture and edits another after the fork.
		git(repo.root, 'checkout', '-q', 'main');
		write(repo.cases, [kase('old', 'Sub Main()\r\n    y = 2\r\nEnd Sub\r\n'), kase('later')]);
		git(repo.root, 'commit', '-qam', 'main moves on');
		git(repo.root, 'update-ref', 'refs/remotes/origin/main', 'HEAD');
		git(repo.root, 'checkout', '-q', 'feature');
		expect(ids(select(repo.cases, branch))).toEqual(['mine']);
	});

	it('prefers upstream/main to origin/main, and takes --changed-from over both', () => {
		const repo = repository([kase('a')]);
		write(repo.cases, [kase('a'), kase('b')]);
		git(repo.root, 'commit', '-qam', 'b');
		const withB = git(repo.root, 'rev-parse', 'HEAD');
		git(repo.root, 'update-ref', 'refs/remotes/upstream/main', withB);
		const now = [kase('a'), kase('b'), kase('c')];
		write(repo.cases, now);
		expect(ids(select(repo.cases, now))).toEqual(['c']);
		expect(ids(select(repo.cases, now, 'origin/main'))).toEqual(['b', 'c']);
	});

	it('works in a folder whose name has spaces and non-ASCII letters', () => {
		const repo = repository([kase('a')], `fix tures ${String.fromCharCode(0xe9, 0x4e2d)}`, `cases ${String.fromCharCode(0xfc)}.json`);
		const now = [kase('a'), kase('b')];
		write(repo.cases, now);
		expect(ids(select(repo.cases, now))).toEqual(['b']);
	});
});

describe('when --changed cannot establish the base', () => {
	it('refuses outside a git work tree', () => {
		const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'xlide-changed-nogit-'));
		made.push(dir);
		const file = path.join(dir, 'cases.json');
		write(file, [kase('a')]);
		expect(select(file, [kase('a')]).refusal).toMatch(/not in a git work tree/);
	});

	it('refuses an untracked cases file', () => {
		const repo = repository([kase('a')]);
		const other = path.join(path.dirname(repo.cases), 'other.json');
		write(other, [kase('a')]);
		expect(select(other, [kase('a')]).refusal).toMatch(/not tracked by git/);
	});

	it('refuses a cases file that did not exist at the base', () => {
		const repo = repository([kase('a')]);
		const other = path.join(path.dirname(repo.cases), 'other.json');
		write(other, [kase('a')]);
		git(repo.root, 'add', '-A');
		git(repo.root, 'commit', '-qm', 'other');
		expect(select(other, [kase('a')]).refusal).toMatch(/does not exist at the base/);
	});

	it('refuses an unreadable base file', () => {
		const repo = repository([kase('a')]);
		git(repo.root, 'checkout', '-q', 'main');
		fs.writeFileSync(repo.cases, '{ not json');
		git(repo.root, 'commit', '-qam', 'broken');
		git(repo.root, 'update-ref', 'refs/remotes/origin/main', 'HEAD');
		git(repo.root, 'checkout', '-q', '-b', 'later');
		write(repo.cases, [kase('a')]);
		expect(select(repo.cases, [kase('a')]).refusal).toMatch(/cannot be read/);
	});

	it('refuses a shallow clone', () => {
		const repo = repository([kase('a')]);
		git(repo.root, 'commit', '-q', '--allow-empty', '-m', 'second');
		const clone = fs.mkdtempSync(path.join(os.tmpdir(), 'xlide-changed-shallow-'));
		made.push(clone);
		git(clone, 'clone', '-q', '--depth', '1', '--branch', 'feature', `file://${repo.root.split(path.sep).join('/')}`, 'copy');
		const file = path.join(clone, 'copy', 'oracle', 'cases.json');
		expect(select(file, [kase('a')]).refusal).toMatch(/shallow/);
	});

	it('refuses when neither upstream/main nor origin/main exists', () => {
		const repo = repository([kase('a')]);
		git(repo.root, 'update-ref', '-d', 'refs/remotes/origin/main');
		expect(select(repo.cases, [kase('a')]).refusal).toMatch(/neither upstream\/main nor origin\/main exists/);
	});

	it('refuses a --changed-from that names no commit, or reads as an option', () => {
		const repo = repository([kase('a')]);
		expect(select(repo.cases, [kase('a')], 'no-such-ref').refusal).toMatch(/does not name a commit/);
		expect(select(repo.cases, [kase('a')], '--output=x').refusal).toMatch(/expects a git ref/);
	});

	it('refuses a base with no history in common', () => {
		const repo = repository([kase('a')]);
		git(repo.root, 'checkout', '-q', '--orphan', 'unrelated');
		git(repo.root, 'commit', '-q', '-m', 'unrelated');
		git(repo.root, 'checkout', '-q', 'feature');
		expect(select(repo.cases, [kase('a')], 'unrelated').refusal).toMatch(/share no history/);
	});
});

describe('the runner with --changed', () => {
	const run = (args: string[], env: NodeJS.ProcessEnv = process.env) => spawnSync(process.execPath, [RUNNER, ...args], { encoding: 'utf8', env });

	it('exits 0 before anything starts when nothing changed', () => {
		const repo = repository([kase('a')]);
		const result = run(['--cases', repo.cases, '--changed']);
		expect(result.stderr).toContain('--changed: 0 new or changed case(s) against origin/main');
		expect(result.status).toBe(0);
	});

	it('refuses with exit code 2, never falling back to every fixture', () => {
		const repo = repository([kase('a')]);
		git(repo.root, 'update-ref', '-d', 'refs/remotes/origin/main');
		const result = run(['--cases', repo.cases, '--changed']);
		expect(result.status).toBe(2);
		expect(result.stderr).toMatch(/^--changed: neither/);
	});

	it('refuses when git cannot be run', () => {
		const repo = repository([kase('a')]);
		const result = run(['--cases', repo.cases, '--changed'], { ...process.env, PATH: '', Path: '' });
		expect(result.status).toBe(2);
		expect(result.stderr).toMatch(/git is not installed or not on PATH/);
	});

	it('cannot be combined with --case', () => {
		const result = run(['--changed', '--case', 'x']);
		expect(result.status).toBe(2);
		expect(result.stderr).toContain('cannot be combined with --case');
	});
});
