// The marks a project's rows carry for modules that differ from HEAD: what
// is computed, when it is recomputed, and when nothing is said at all.

import { beforeEach, describe, expect, it, vi } from 'vitest';

vi.mock('vscode', async () => (await import('./helpers/vscodeMock')).vscodeMock());

import type { GitRunResult } from '../src/gitFileHistory';
import { GitChangeMarks, type GitChangeMarksDeps } from '../src/gitChangeMarks';
import { changedCountBadge } from '../src/agentReviewDecorations';
import type { ModuleSnapshot } from '../src/gitModuleCompare';

const PROJECT = process.platform === 'win32' ? 'C:\\work\\Book.xlsm' : '/work/Book.xlsm';
const ROOT = process.platform === 'win32' ? 'C:\\work' : '/work';
const HEAD = 'a'.repeat(40);

function result(code: number, stdout = ''): GitRunResult {
	return { code, stdout: Buffer.from(stdout), stderr: '' };
}

interface Fixture {
	committed: ModuleSnapshot[] | undefined;
	current: ModuleSnapshot[];
	stamp: { mtimeMs: number; size: number } | undefined;
	tracked: boolean;
	inRepo: boolean;
	head: string;
}

function fixture(overrides: Partial<Fixture> = {}): Fixture {
	return {
		committed: [{ name: 'Same', source: 'x' }, { name: 'Changed', source: 'a' }, { name: 'Gone', source: 'g' }],
		current: [{ name: 'Same', source: 'x' }, { name: 'Changed', source: 'b' }, { name: 'Fresh', source: 'f' }],
		stamp: { mtimeMs: 1, size: 10 },
		tracked: true,
		inRepo: true,
		head: HEAD,
		...overrides,
	};
}

function deps(f: Fixture): GitChangeMarksDeps & { runs: string[][] } {
	const runs: string[][] = [];
	return {
		runs,
		git: {
			run: vi.fn(async (args: readonly string[]) => {
				runs.push([...args]);
				if (args[0] === 'rev-parse' && args[1] === '--show-toplevel') {
					return f.inRepo ? result(0, `${ROOT.replace(/\\/g, '/')}\n`) : result(128);
				}
				if (args[0] === 'rev-parse') {
					return result(0, `${f.head}\n`);
				}
				if (args[0] === 'ls-files') {
					return result(f.tracked ? 0 : 1);
				}
				return result(0);
			}),
		},
		stat: () => f.stamp,
		currentModules: vi.fn(async () => f.current),
		modulesAtRevision: vi.fn(async () => f.committed),
	};
}

/** Lets the scheduled computation land. */
async function settle(): Promise<void> {
	for (let i = 0; i < 10; i++) {
		await new Promise((resolve) => setTimeout(resolve, 0));
	}
}

let changes: string[];

beforeEach(() => {
	changes = [];
});

describe('GitChangeMarks', () => {
	it('answers nothing on the first ask, computes, and then answers the marks', async () => {
		const f = fixture();
		const marks = new GitChangeMarks(deps(f));
		marks.onDidChange((projectPath) => changes.push(projectPath));

		expect(marks.marksFor(PROJECT)).toBeUndefined();
		await settle();

		expect(changes).toEqual([PROJECT]);
		const answer = marks.marksFor(PROJECT)!;
		expect([...answer.byModule.entries()]).toEqual([['changed', 'modified'], ['fresh', 'added']]);
		expect(answer.removed).toBe(1);
		expect(answer.head).toBe(HEAD);
	});

	it('compares against HEAD by its hash, so a later commit is a different answer', async () => {
		const f = fixture();
		const d = deps(f);
		const marks = new GitChangeMarks(d);
		marks.marksFor(PROJECT);
		await settle();
		expect(d.modulesAtRevision).toHaveBeenCalledWith(PROJECT, expect.objectContaining({ relativePath: 'Book.xlsm' }), HEAD);
	});

	it('serves the cache while the file is unchanged and recomputes when it changes', async () => {
		const f = fixture();
		const d = deps(f);
		const marks = new GitChangeMarks(d);
		marks.onDidChange((projectPath) => changes.push(projectPath));
		marks.marksFor(PROJECT);
		await settle();
		marks.marksFor(PROJECT);
		marks.marksFor(PROJECT);
		await settle();
		expect(d.modulesAtRevision).toHaveBeenCalledTimes(1);

		f.stamp = { mtimeMs: 2, size: 11 };
		f.current = [{ name: 'Same', source: 'x' }, { name: 'Changed', source: 'a' }, { name: 'Gone', source: 'g' }];
		// The stale answer is served while the fresh one is computed.
		expect(marks.marksFor(PROJECT)?.byModule.size).toBe(2);
		await settle();
		expect(d.modulesAtRevision).toHaveBeenCalledTimes(2);
		expect(marks.marksFor(PROJECT)?.byModule.size).toBe(0);
		expect(changes).toEqual([PROJECT, PROJECT]);
	});

	it('does not fire when a recomputation lands on the same marks', async () => {
		const f = fixture();
		const marks = new GitChangeMarks(deps(f));
		marks.onDidChange((projectPath) => changes.push(projectPath));
		marks.marksFor(PROJECT);
		await settle();
		marks.invalidateAll();
		await settle();
		expect(changes).toEqual([PROJECT]);
	});

	it('recomputes every asked-for project when the repository changes', async () => {
		const f = fixture();
		const d = deps(f);
		const marks = new GitChangeMarks(d);
		marks.marksFor(PROJECT);
		await settle();
		f.head = 'b'.repeat(40);
		marks.invalidateAll();
		await settle();
		expect(d.modulesAtRevision).toHaveBeenCalledTimes(2);
		expect(marks.marksFor(PROJECT)?.head).toBe('b'.repeat(40));
	});

	it('answers nothing for a file outside a repository, an untracked file, or one git cannot read', async () => {
		for (const f of [fixture({ inRepo: false }), fixture({ tracked: false }), fixture({ committed: undefined })]) {
			const d = deps(f);
			const marks = new GitChangeMarks(d);
			marks.marksFor(PROJECT);
			await settle();
			expect(marks.marksFor(PROJECT)).toBeUndefined();
		}
	});

	it('asks git once per project while a computation is in flight', async () => {
		const f = fixture();
		const d = deps(f);
		const marks = new GitChangeMarks(d);
		marks.marksFor(PROJECT);
		marks.marksFor(PROJECT);
		marks.marksFor(PROJECT);
		await settle();
		expect(d.runs.filter((args) => args[0] === 'rev-parse' && args[1] === '--show-toplevel')).toHaveLength(1);
	});

	it('logs a git failure and keeps answering nothing for that project', async () => {
		const f = fixture();
		const d = deps(f);
		d.modulesAtRevision = vi.fn(async () => { throw new Error('boom'); });
		const log = vi.fn();
		const marks = new GitChangeMarks(d, log);
		marks.marksFor(PROJECT);
		await settle();
		expect(log.mock.calls[0][0]).toContain('boom');
		expect(marks.marksFor(PROJECT)).toBeUndefined();
	});
});

describe('changedCountBadge', () => {
	it('fits the count into two characters', () => {
		expect(changedCountBadge(1)).toBe('1');
		expect(changedCountBadge(42)).toBe('42');
		expect(changedCountBadge(150)).toBe('99');
	});
});
