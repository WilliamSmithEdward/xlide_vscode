import { describe, expect, it } from 'vitest';
import * as fs from 'fs';
import * as path from 'path';
import * as nodePath from 'path';
import * as webPath from '../src/util/webPath';

// node:path's posix implementation is the oracle. The browser build aliases
// `path` to src/util/webPath.ts, so any disagreement here is a path the web
// build would compute differently from the desktop one - which, for a module
// whose answers become file locations, is how a user loses work.

/** Path shapes that exercise roots, dots, repeats, trailing slashes. */
const SHAPES = [
	'',
	'.',
	'..',
	'...',
	'/',
	'//',
	'///',
	'a',
	'a/b',
	'a/b/c',
	'/a',
	'/a/b',
	'/a/b/',
	'a/',
	'a//b',
	'a/./b',
	'a/../b',
	'a/b/..',
	'a/b/../..',
	'a/b/../../..',
	'../a',
	'../../a',
	'./a',
	'/..',
	'/../a',
	'/a/../..',
	'.hidden',
	'.hidden.txt',
	'Book.xlsm',
	'Book.xlsm/',
	'/repo/Book.xlsm',
	'/repo/sub/Book.with.dots.xlsm',
	'repo/Module1.bas',
	'a.',
	'a..',
	'/a.b/c',
	'/a/b.txt',
	'noext',
	'/trailing/dot.',
	'vscode-vfs/github/owner/repo/Book.xlsm',
];

describe('the shim covers what the codebase actually calls', () => {
	// Same lesson the Buffer shim learned the hard way: tsc type-checks
	// against @types/node, not against this file, so a path function the
	// shim lacks compiles fine and fails only in a browser.
	it('implements every path member src/ uses', () => {
		const srcDir = path.join(__dirname, '..', 'src');
		const used = new Set<string>();

		const walk = (dir: string): void => {
			for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
				const full = path.join(dir, entry.name);
				if (entry.isDirectory()) {
					walk(full);
				} else if (entry.name.endsWith('.ts')) {
					for (const match of fs.readFileSync(full, 'utf8').matchAll(/\bpath\.([a-z][A-Za-z0-9]*)/g)) {
						// The codebase has plenty of string variables named
						// `path`; only names node:path actually exports count.
						if (match[1] in nodePath) {
							used.add(match[1]);
						}
					}
				}
			}
		};
		walk(srcDir);

		expect(used.size, 'found no path members - has the scan broken?').toBeGreaterThan(5);
		const shim = webPath as unknown as Record<string, unknown>;
		expect([...used].filter((name) => shim[name] === undefined),
			'src/ calls these, the browser build has none of them').toEqual([]);
	});
});

describe('webPath matches node:path posix', () => {
	it.each(SHAPES)('normalize(%j)', (p) => {
		expect(webPath.normalize(p)).toBe(nodePath.posix.normalize(p));
	});

	it.each(SHAPES)('dirname(%j)', (p) => {
		expect(webPath.dirname(p)).toBe(nodePath.posix.dirname(p));
	});

	it.each(SHAPES)('basename(%j)', (p) => {
		expect(webPath.basename(p)).toBe(nodePath.posix.basename(p));
	});

	it.each(SHAPES)('extname(%j)', (p) => {
		expect(webPath.extname(p)).toBe(nodePath.posix.extname(p));
	});

	it.each(SHAPES)('isAbsolute(%j)', (p) => {
		expect(webPath.isAbsolute(p)).toBe(nodePath.posix.isAbsolute(p));
	});

	it.each(SHAPES)('resolve(%j)', (p) => {
		// resolve() falls back to a cwd; a browser has none, so the root
		// stands in. Compare against node with the same base.
		expect(webPath.resolve(p)).toBe(nodePath.posix.resolve('/', p));
	});

	it('basename with an extension argument', () => {
		const cases: [string, string][] = [
			['Book.xlsm', '.xlsm'],
			['Module1.bas', '.bas'],
			['.txt', '.txt'],
			['a.txt', '.txt'],
			['a.txt', '.md'],
			['/repo/Form1.frm', '.frm'],
			['noext', '.bas'],
			['.bas', '.bas'],
		];
		for (const [p, ext] of cases) {
			expect(webPath.basename(p, ext), `basename(${p}, ${ext})`).toBe(
				nodePath.posix.basename(p, ext),
			);
		}
	});

	it('join over every pair of shapes', () => {
		for (const a of SHAPES) {
			for (const b of SHAPES) {
				expect(webPath.join(a, b), `join(${JSON.stringify(a)}, ${JSON.stringify(b)})`).toBe(
					nodePath.posix.join(a, b),
				);
			}
		}
	});

	it('join over every triple of a smaller set', () => {
		const few = ['', '.', '..', '/', 'a', 'a/', '/a', 'b/c'];
		for (const a of few) {
			for (const b of few) {
				for (const c of few) {
					expect(webPath.join(a, b, c), `join(${a}, ${b}, ${c})`).toBe(
						nodePath.posix.join(a, b, c),
					);
				}
			}
		}
	});

	it('resolve over every pair of shapes', () => {
		for (const a of SHAPES) {
			for (const b of SHAPES) {
				expect(webPath.resolve(a, b), `resolve(${JSON.stringify(a)}, ${JSON.stringify(b)})`).toBe(
					nodePath.posix.resolve('/', a, b),
				);
			}
		}
	});

	it('relative over every pair of absolute shapes', () => {
		const absolute = SHAPES.filter((p) => p.startsWith('/'));
		for (const from of absolute) {
			for (const to of absolute) {
				expect(webPath.relative(from, to), `relative(${from}, ${to})`).toBe(
					nodePath.posix.relative(from, to),
				);
			}
		}
	});

	it('agrees on the separator', () => {
		expect(webPath.sep).toBe(nodePath.posix.sep);
		expect(webPath.delimiter).toBe(nodePath.posix.delimiter);
	});
});

describe('the win32 namespace', () => {
	it('refuses rather than answering wrongly', () => {
		// A web workspace has no Windows paths. Callers reach these only
		// through a platform test that is false in a browser.
		expect(() => webPath.win32.normalize('C:\\a\\b')).toThrow(/POSIX/);
		expect(() => webPath.win32.basename('C:\\a\\b')).toThrow(/POSIX/);
	});

	it('still reports the Windows separator, which is a constant', () => {
		expect(webPath.win32.sep).toBe(nodePath.win32.sep);
		expect(webPath.win32.delimiter).toBe(nodePath.win32.delimiter);
	});
});
