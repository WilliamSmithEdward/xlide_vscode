// The git file-history helpers against a real repository: a workbook is
// committed, edited on disk, and read back at HEAD as the modules it had.

import { execFileSync } from 'node:child_process';
import * as fs from 'node:fs';
import * as os from 'node:os';
import * as path from 'node:path';
import { afterAll, beforeAll, describe, expect, it } from 'vitest';
import {
	GitUnavailableError,
	gitFileAtRevision,
	gitFileHistory,
	gitFileIsTracked,
	gitFileRef,
	gitRevisionLabel,
	gitRunner,
} from '../src/gitFileHistory';
import { readModules, readModulesFromBuffer, writeModule } from '../src/vba/projectService';

const FIXTURE = path.join(__dirname, 'fixtures', 'binaries', 'FormFixture.xlsm');

function gitAvailable(): boolean {
	try {
		execFileSync('git', ['--version'], { stdio: 'ignore' });
		return true;
	} catch {
		return false;
	}
}

describe.skipIf(!gitAvailable())('gitFileHistory against a real repository', () => {
	let root: string;
	let workbook: string;
	const git = gitRunner();
	const run = (args: string[]): void => {
		execFileSync('git', args, { cwd: root, stdio: 'ignore' });
	};

	beforeAll(() => {
		root = fs.realpathSync.native(fs.mkdtempSync(path.join(os.tmpdir(), 'xlide-git-')));
		fs.mkdirSync(path.join(root, 'books'));
		workbook = path.join(root, 'books', 'Fixture.xlsm');
		fs.copyFileSync(FIXTURE, workbook);
		run(['init', '-q']);
		run(['config', 'user.email', 'test@example.com']);
		run(['config', 'user.name', 'Test']);
		run(['add', '.']);
		run(['commit', '-q', '-m', 'add workbook']);
	});

	afterAll(() => {
		fs.rmSync(root, { recursive: true, force: true });
	});

	it('finds the repository and the path git knows the file by', async () => {
		const ref = await gitFileRef(workbook, git);
		expect(ref).toBeDefined();
		expect(fs.realpathSync.native(ref!.root)).toBe(root);
		expect(ref!.relativePath).toBe('books/Fixture.xlsm');
	});

	it('answers undefined for a file outside any repository', async () => {
		const outside = fs.mkdtempSync(path.join(os.tmpdir(), 'xlide-nogit-'));
		try {
			expect(await gitFileRef(path.join(outside, 'x.xlsm'), git)).toBeUndefined();
		} finally {
			fs.rmSync(outside, { recursive: true, force: true });
		}
	});

	it('tells a tracked file from an untracked one', async () => {
		const ref = (await gitFileRef(workbook, git))!;
		expect(await gitFileIsTracked(ref, git)).toBe(true);
		const loose = path.join(root, 'books', 'Loose.xlsm');
		fs.copyFileSync(FIXTURE, loose);
		expect(await gitFileIsTracked((await gitFileRef(loose, git))!, git)).toBe(false);
	});

	it('lists the commits that touched the file, newest first', async () => {
		const ref = (await gitFileRef(workbook, git))!;
		const history = await gitFileHistory(ref, git);
		expect(history).toHaveLength(1);
		expect(history[0].subject).toBe('add workbook');
		expect(history[0].author).toBe('Test');
		expect(history[0].shortHash).toBe(history[0].hash.slice(0, history[0].shortHash.length));
		expect(history[0].date).toMatch(/^\d{4}-\d{2}-\d{2}$/);
	});

	it('reads the committed bytes at HEAD, and the modules they hold, after the file changed on disk', async () => {
		const ref = (await gitFileRef(workbook, git))!;
		const before = (await gitFileAtRevision(ref, 'HEAD', git))!;
		expect(before.equals(fs.readFileSync(workbook))).toBe(true);

		const target = readModules(workbook).find((module) => module.type === 'standard' && module.source !== undefined)!;
		const original = target.source!;
		writeModule(workbook, target.name, `${original}\n' changed on disk\n`, 'standard');

		const committed = readModulesFromBuffer((await gitFileAtRevision(ref, 'HEAD', git))!)
			.find((module) => module.name === target.name)!;
		expect(committed.source).toBe(original);
		const now = readModules(workbook).find((module) => module.name === target.name)!;
		expect(now.source).toContain("' changed on disk");
	});

	it('answers undefined for a revision without the file', async () => {
		const ref = (await gitFileRef(workbook, git))!;
		expect(await gitFileAtRevision({ ...ref, relativePath: 'books/Missing.xlsm' }, 'HEAD', git)).toBeUndefined();
	});

	it('reports a git executable that cannot run', async () => {
		await expect(gitFileRef(workbook, gitRunner('xlide-no-such-git-executable')))
			.rejects.toBeInstanceOf(GitUnavailableError);
	});
});

describe('gitRevisionLabel', () => {
	it('shortens a full hash and keeps a symbolic revision', () => {
		expect(gitRevisionLabel('HEAD')).toBe('HEAD');
		expect(gitRevisionLabel('0123456789abcdef0123456789abcdef01234567')).toBe('0123456');
		expect(gitRevisionLabel('0123456789abcdef0123456789abcdef01234567', {
			hash: '0123456789abcdef0123456789abcdef01234567', shortHash: '0123456', author: 'a', date: 'd', subject: 's',
		})).toBe('0123456');
	});
});
