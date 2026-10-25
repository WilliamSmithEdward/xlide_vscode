// --changed: the VBE oracle runs only the fixtures that are new or changed
// since the branch left main, compared through git.

import { spawnSync } from 'node:child_process';
import * as fs from 'node:fs';
import * as path from 'node:path';
import { runInNewContext } from 'node:vm';
import { afterAll, describe, expect, it } from 'vitest';
import { OFFICE_HOST_APPS } from '../../src/officeHostApps';
import { makeTempDir, removeTempDirs } from '../helpers/tempDirs';

const coordinator = path.join(process.cwd(), 'syntax_corpus/oracle/run_vbe_oracle.mjs');
// A Windows checkout has CRLF lines.
const source = fs.readFileSync(coordinator, 'utf8').replaceAll('\r\n', '\n');
const lift = (name: string): string => {
	const start = source.indexOf(`function ${name}(`);
	return source.slice(start, source.indexOf('\n}', start) + 2);
};
const constant = (name: string): string => {
	const start = source.indexOf(`const ${name} =`);
	return source.slice(start, source.indexOf(';\n', start) + 2);
};

interface Kase { id: string | number; host?: string; source?: string; expected?: string }
type Base = { ref: string; commit: string; cases: Kase[] } | { error: string };

const api = runInNewContext(`
	const DEFAULT_HOST = 'excel';
	const DEFAULT_CASES = 'vbe_oracle_cases.json';
	const OFFICE_HOSTS = ${JSON.stringify(Object.fromEntries(Object.keys(OFFICE_HOST_APPS).map((host) => [host, {}])))};
	const DEFAULT_TIMEOUT_RETRIES = 2;
	const DEFAULT_CASES_PER_INSTANCE = 40;
	${constant('CHANGED_BASE_REFS')}
	${['parseArgs', 'changedBase', 'changedCases', 'describeChanged', 'caseHost'].map(lift).join('\n')}
	({ parseArgs, changedBase, changedCases, describeChanged })`, { path, spawnSync }) as {
	parseArgs: (argv: string[]) => { changed: boolean; changedFrom: string | null; caseIds: string[] };
	changedBase: (casesFile: string, ref: string | null) => Base;
	changedCases: (cases: Kase[], baseCases: Kase[]) => Kase[];
	describeChanged: (cases: Kase[], total: number, base: { ref: string; commit: string }) => string;
};

afterAll(removeTempDirs);

// Starting git is slow on Windows, more so while other test files run: a test
// there took up to 12.5 s when each made its repository from scratch. So each
// starting repository is made once and copied, and the limit is 20 s.
const GIT_TESTS = { timeout: 20_000 };

function git(dir: string, ...args: string[]): string {
	const run = spawnSync('git', ['-C', dir, '-c', 'user.name=t', '-c', 'user.email=t@example.com', '-c', 'commit.gpgsign=false', ...args], { encoding: 'utf8' });
	if (run.status !== 0) { throw new Error(`git ${args.join(' ')}: ${run.stderr}`); }
	return run.stdout.trim();
}

const doc = (...cases: Kase[]): string => `${JSON.stringify({ cases }, null, 2)}\n`;
const a: Kase = { id: 'a', source: 'Sub A()\nEnd Sub', expected: 'accepted' };
const b: Kase = { id: 'b', host: 'word', source: 'Sub B()\nEnd Sub', expected: 'rejected' };
const c: Kase = { id: 'c', host: 'access', source: 'Sub C()\nEnd Sub', expected: 'accepted' };

const made = new Map<string, { dir: string; main: string }>();

/** A repository of its own whose main holds `cases` at `file`, with an upstream/main ref at that commit and a branch checked out from it. */
function repo(cases: string, file = 'vbe_oracle_cases.json'): { dir: string; file: string; main: string } {
	const key = `${file}\0${cases}`;
	let first = made.get(key);
	if (!first) {
		const dir = makeTempDir('xlide-oracle-changed-');
		git(dir, 'init', '-q', '-b', 'main');
		fs.mkdirSync(path.dirname(path.join(dir, file)), { recursive: true });
		fs.writeFileSync(path.join(dir, file), cases);
		git(dir, 'add', '-A');
		git(dir, 'commit', '-q', '-m', 'main');
		const main = git(dir, 'rev-parse', 'HEAD');
		git(dir, 'update-ref', 'refs/remotes/upstream/main', main);
		git(dir, 'checkout', '-q', '-b', 'work');
		first = { dir, main };
		made.set(key, first);
	}
	const dir = makeTempDir('xlide-oracle-changed-');
	fs.cpSync(first.dir, dir, { recursive: true });
	return { dir, file: path.join(dir, file), main: first.main };
}

/** The base, and the ids of the cases on disk that changed since it. */
function compare(file: string, ref: string | null = null): { base: Base; ids: string[] } {
	const base = api.changedBase(file, ref);
	if ('error' in base) { throw new Error(base.error); }
	const current = JSON.parse(fs.readFileSync(file, 'utf8').replace(/^\uFEFF/, '')).cases as Kase[];
	return { base, ids: api.changedCases(current, base.cases).map((kase) => String(kase.id)) };
}

const selected = (file: string, ref: string | null = null): string[] => compare(file, ref).ids;

describe('--changed options', () => {
	it('--changed compares with the default branch, --changed-from with the one named', () => {
		expect(api.parseArgs(['--changed'])).toMatchObject({ changed: true, changedFrom: null });
		expect(api.parseArgs(['--changed-from', 'origin/release'])).toMatchObject({ changed: true, changedFrom: 'origin/release' });
		expect(api.parseArgs([])).toMatchObject({ changed: false, changedFrom: null });
	});

	it('refuses --changed-from without a branch', () => {
		expect(() => api.parseArgs(['--changed-from'])).toThrow('--changed-from expects a value');
		expect(() => api.parseArgs(['--changed-from', '--strict'])).toThrow("--changed-from expects a value, not the option '--strict'");
	});

	it('refuses --changed with --case, in either order', () => {
		for (const argv of [['--changed', '--case', 'a'], ['--case', 'a', '--changed'], ['--changed-from', 'x', '--case', 'a']]) {
			expect(() => api.parseArgs(argv), argv.join(' ')).toThrow('--changed and --case each choose the cases to run; give one of them');
		}
	});
});

describe('which cases count as changed', () => {
	it('takes new cases and cases that differ in any field, and leaves the rest', () => {
		const changedB = { ...b, expected: 'accepted' };
		expect(api.changedCases([a, changedB, c], [a, b]).map((k) => k.id)).toEqual(['b', 'c']);
		expect(api.changedCases([a, b], [a, b])).toEqual([]);
	});

	it('counts a case whose fields only moved, and a renamed case, as changed', () => {
		const reordered = { expected: a.expected, source: a.source, id: a.id };
		expect(api.changedCases([reordered], [a]).map((k) => k.id)).toEqual(['a']);
		expect(api.changedCases([{ ...a, id: 'a2' }], [a]).map((k) => k.id)).toEqual(['a2']);
	});

	it('matches ids as the rest of the oracle does, as strings', () => {
		expect(api.changedCases([{ ...a, id: 7 }], [{ ...a, id: '7' }]).map((k) => k.id)).toEqual([7]);
		expect(api.changedCases([{ ...a, id: 7 }], [{ ...a, id: 7 }])).toEqual([]);
	});

	it('ignores cases removed since the base', () => {
		expect(api.changedCases([a], [a, b, c])).toEqual([]);
	});

	it('says how many it will run, of how many, against which commit, by host', () => {
		const base = { ref: 'upstream/main', commit: '0123456789abcdef0123' };
		expect(api.describeChanged([a, b, c, { ...b, id: 'b2' }], 14868, base))
			.toBe('--changed: 4 of 14868 case(s) are new or changed since this branch left upstream/main at 0123456789ab (1 excel, 2 word, 1 access).');
		expect(api.describeChanged([], 3, base)).toBe('--changed: 0 of 3 case(s) are new or changed since this branch left upstream/main at 0123456789ab.');
	});
});

describe('the base it compares with', GIT_TESTS, () => {
	it('is the commit where the branch left upstream/main, so later commits on main are left out', () => {
		const r = repo(doc(a, b));
		// main moves on after the branch left it: it changes a and adds c.
		git(r.dir, 'checkout', '-q', 'main');
		fs.writeFileSync(r.file, doc({ ...a, expected: 'rejected' }, b, c));
		git(r.dir, 'commit', '-q', '-am', 'later on main');
		git(r.dir, 'update-ref', 'refs/remotes/upstream/main', 'HEAD');
		git(r.dir, 'checkout', '-q', 'work');
		const { base, ids } = compare(r.file);
		expect(base).toMatchObject({ ref: 'upstream/main', commit: r.main });
		expect(ids).toEqual([]);
	});

	it('counts the branch\'s commits and its uncommitted edits', () => {
		const r = repo(doc(a, b));
		fs.writeFileSync(r.file, doc(a, b, c));
		git(r.dir, 'commit', '-q', '-am', 'add c');
		fs.writeFileSync(r.file, doc({ ...a, source: 'Sub A()\n  x\nEnd Sub' }, b, c));
		expect(selected(r.file)).toEqual(['a', 'c']);
	});

	it('takes upstream/main before origin/main, origin/main without an upstream, and the branch named', () => {
		const r = repo(doc(a));
		fs.writeFileSync(r.file, doc(a, b));
		git(r.dir, 'commit', '-q', '-am', 'add b');
		git(r.dir, 'update-ref', 'refs/remotes/origin/main', 'HEAD');
		expect(compare(r.file)).toMatchObject({ base: { ref: 'upstream/main', commit: r.main }, ids: ['b'] });
		git(r.dir, 'update-ref', '-d', 'refs/remotes/upstream/main');
		expect(compare(r.file)).toMatchObject({ base: { ref: 'origin/main' }, ids: [] });
		expect(selected(r.file, 'main')).toEqual(['b']);
	});

	it('names --changed-from when there is neither upstream/main nor origin/main', () => {
		const r = repo(doc(a));
		git(r.dir, 'update-ref', '-d', 'refs/remotes/upstream/main');
		expect(api.changedBase(r.file, null)).toEqual({
			error: '--changed: neither upstream/main nor origin/main exists here; name the branch to compare with, --changed-from <ref>',
		});
	});

	it('refuses a branch with no commit in common with the base', () => {
		const r = repo(doc(a));
		git(r.dir, 'checkout', '-q', '--orphan', 'unrelated');
		git(r.dir, 'commit', '-q', '-m', 'unrelated');
		expect(api.changedBase(r.file, null)).toEqual({ error: '--changed: cannot find where this branch left upstream/main (no commit in common)' });
	});

	it('refuses a branch that does not exist', () => {
		const r = repo(doc(a));
		const base = api.changedBase(r.file, 'nowhere/main') as { error: string };
		expect(base.error).toMatch(/^--changed: cannot find where this branch left nowhere\/main \(/);
	});

	it('gives every case when the file did not exist at the base', () => {
		const r = repo(doc(a));
		const extra = path.join(r.dir, 'more_cases.json');
		fs.writeFileSync(extra, doc(b, c));
		git(r.dir, 'add', 'more_cases.json');
		expect(compare(extra)).toEqual({ base: { ref: 'upstream/main', commit: r.main, cases: [] }, ids: ['b', 'c'] });
	});

	it('finds a cases file in a subfolder whose name has spaces and letters git quotes', () => {
		const r = repo(doc(a), path.join('fixtures é', 'my cases.json'));
		fs.writeFileSync(r.file, doc(a, b));
		expect(selected(r.file)).toEqual(['b']);
	});

	it('reads a base file written with a byte order mark', () => {
		const r = repo(`\uFEFF${doc(a, b)}`);
		fs.writeFileSync(r.file, `\uFEFF${doc(a, { ...b, expected: 'accepted' })}`);
		expect(selected(r.file)).toEqual(['b']);
	});

	it('refuses a base file that is not JSON, or has no cases array, rather than measure everything', () => {
		for (const text of ['{ not json', '{}', '{ "cases": null }', 'null', '[]']) {
			const r = repo(text);
			fs.writeFileSync(r.file, doc(a));
			const base = api.changedBase(r.file, null) as { error: string };
			expect(base.error, text).toMatch(/^--changed: cannot read vbe_oracle_cases\.json as it was at [0-9a-f]{12} \(/);
		}
	});
});

describe('where --changed refuses to guess', GIT_TESTS, () => {
	it('says so when git cannot be found', () => {
		const r = repo(doc(a));
		const noGit = runInNewContext(`${constant('CHANGED_BASE_REFS')}\n${lift('changedBase')}\nchangedBase`, {
			path, spawnSync: (_command: string, args: string[], options: object) => spawnSync('xlide-no-such-git', args, options),
		}) as typeof api.changedBase;
		expect(noGit(r.file, null)).toEqual({ error: '--changed needs git, and git is not on the PATH' });
	});

	it('refuses a base git cannot read rather than measure everything', () => {
		// Removing one loose object makes the base's tree, then its file, unreadable.
		const unreadable = (object: string): { error: string } => {
			const r = repo(doc(a));
			fs.writeFileSync(r.file, doc(a, b));
			const sha = git(r.dir, 'rev-parse', `${r.main}${object}`);
			fs.rmSync(path.join(r.dir, '.git', 'objects', sha.slice(0, 2), sha.slice(2)), { force: true });
			return api.changedBase(r.file, null) as { error: string };
		};
		expect(unreadable('^{tree}').error).toMatch(/^--changed: cannot list vbe_oracle_cases\.json at [0-9a-f]{12} \(/);
		expect(unreadable(':vbe_oracle_cases.json').error).toMatch(/^--changed: cannot read vbe_oracle_cases\.json as it was at [0-9a-f]{12} \(/);
	});

	it('refuses a cases file git does not track, and one outside any repository', () => {
		const r = repo(doc(a));
		const loose = path.join(r.dir, 'loose.json');
		fs.writeFileSync(loose, doc(a));
		expect((api.changedBase(loose, null) as { error: string }).error).toMatch(/^--changed: git does not track .*loose\.json, so there is nothing to compare it with \(/);
		const outside = path.join(makeTempDir('xlide-oracle-nogit-'), 'cases.json');
		fs.writeFileSync(outside, doc(a));
		expect((api.changedBase(outside, null) as { error: string }).error).toMatch(/^--changed: git does not track .*cases\.json/);
	});

	it('refuses a shallow clone, which may not hold the base', () => {
		const r = repo(doc(a));
		fs.writeFileSync(r.file, doc(a, b));
		git(r.dir, 'commit', '-q', '-am', 'add b');
		const clone = path.join(makeTempDir('xlide-oracle-shallow-'), 'clone');
		const run = spawnSync('git', ['clone', '-q', '--depth', '1', '--branch', 'work', `file://${r.dir.replace(/\\/g, '/')}`, clone], { encoding: 'utf8' });
		expect(run.status, run.stderr).toBe(0);
		git(clone, 'update-ref', 'refs/remotes/upstream/main', 'HEAD');
		expect(api.changedBase(path.join(clone, 'vbe_oracle_cases.json'), null)).toEqual({
			error: '--changed: this is a shallow clone, which may not hold the commit where this branch left main; run git fetch --unshallow first',
		});
	});
});

describe.runIf(process.platform === 'win32')('--changed on Windows', GIT_TESTS, () => {
	// Nothing changed, so the run ends before it takes the lock or starts Office.
	it('says so, and ends with nothing to measure', () => {
		const r = repo(doc(a, b, c));
		const run = spawnSync(process.execPath, [coordinator, '--cases', r.file, '--changed', '--strict'], { encoding: 'utf8', timeout: 60_000 });
		expect(run.status, run.stderr).toBe(0);
		expect(run.stderr).toContain(`--changed: 0 of 3 case(s) are new or changed since this branch left upstream/main at ${r.main.slice(0, 12)}.`);
		const json = spawnSync(process.execPath, [coordinator, '--cases', r.file, '--changed', '--json'], { encoding: 'utf8', timeout: 60_000 });
		expect(json.status, json.stderr).toBe(0);
		expect(JSON.parse(json.stdout)).toEqual({ results: [], failureCount: 0, oracleFailureCount: 0, skippedCount: 0 });
	});

	it('refuses a base it cannot read before it starts anything', () => {
		const r = repo(doc(a));
		git(r.dir, 'update-ref', '-d', 'refs/remotes/upstream/main');
		const run = spawnSync(process.execPath, [coordinator, '--cases', r.file, '--changed'], { encoding: 'utf8', timeout: 60_000 });
		expect(run.status).toBe(2);
		expect(run.stderr).toContain('--changed: neither upstream/main nor origin/main exists here');
	});
});
