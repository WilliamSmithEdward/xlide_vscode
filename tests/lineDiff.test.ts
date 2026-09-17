// The line diff: a correct edit script, and the unified text git would print.

import { describe, expect, it } from 'vitest';
import { diffLines, splitLines, unifiedDiff } from '../src/util/lineDiff';

/** Applies an edit script to `before` and hands back the lines it yields. */
function apply(before: readonly string[], ops: ReturnType<typeof diffLines>): string[] {
	const out: string[] = [];
	let cursor = 0;
	for (const op of ops) {
		if (op.kind === 'delete') {
			expect(before[cursor]).toBe(op.line);
			cursor++;
		} else if (op.kind === 'equal') {
			expect(before[cursor]).toBe(op.line);
			out.push(op.line);
			cursor++;
		} else {
			out.push(op.line);
		}
	}
	expect(cursor).toBe(before.length);
	return out;
}

describe('diffLines', () => {
	it('turns before into after for inserts, deletes and replacements', () => {
		const cases: Array<[string[], string[]]> = [
			[[], []],
			[[], ['a']],
			[['a'], []],
			[['a', 'b', 'c'], ['a', 'b', 'c']],
			[['a', 'b', 'c'], ['a', 'x', 'c']],
			[['a', 'b', 'c'], ['a', 'c']],
			[['a', 'c'], ['a', 'b', 'c']],
			[['a', 'b', 'c', 'd'], ['d', 'c', 'b', 'a']],
			[['x', 'a', 'b', 'y'], ['a', 'b']],
			[['a', 'b'], ['x', 'a', 'b', 'y']],
			[['same', 'same', 'same'], ['same', 'same']],
		];
		for (const [before, after] of cases) {
			expect(apply(before, diffLines(before, after)), JSON.stringify([before, after])).toEqual(after);
		}
	});

	it('keeps a shared head and tail as equal lines and edits only the middle', () => {
		const ops = diffLines(['h', 'a', 'b', 't'], ['h', 'a', 'c', 't']);
		expect(ops).toEqual([
			{ kind: 'equal', line: 'h' },
			{ kind: 'equal', line: 'a' },
			{ kind: 'delete', line: 'b' },
			{ kind: 'insert', line: 'c' },
			{ kind: 'equal', line: 't' },
		]);
	});

	it('finds a shortest script', () => {
		// One change, not a delete-everything-and-insert.
		const before = Array.from({ length: 200 }, (_, i) => `line ${i}`);
		const after = [...before.slice(0, 100), 'inserted', ...before.slice(100)];
		const ops = diffLines(before, after);
		expect(ops.filter((op) => op.kind !== 'equal')).toEqual([{ kind: 'insert', line: 'inserted' }]);
	});
});

describe('splitLines', () => {
	it('drops line endings and the empty line a trailing terminator would add', () => {
		expect(splitLines('')).toEqual([]);
		expect(splitLines('a\r\nb\r\n')).toEqual(['a', 'b']);
		expect(splitLines('a\nb')).toEqual(['a', 'b']);
		expect(splitLines('a\n\n')).toEqual(['a', '']);
	});
});

describe('unifiedDiff', () => {
	const labels = { beforeLabel: 'Module1 (HEAD)', afterLabel: 'Module1 (current)' };

	it('is empty for the same lines, whatever their endings', () => {
		expect(unifiedDiff('Sub A()\r\nEnd Sub\r\n', 'Sub A()\nEnd Sub\n', labels)).toBe('');
	});

	it('prints one hunk with three lines of context, in git\'s form', () => {
		const before = ['Option Explicit', '', 'Sub A()', '    x = 1', '    y = 2', '    z = 3', 'End Sub', '', 'Sub B()', 'End Sub'].join('\n');
		const after = ['Option Explicit', '', 'Sub A()', '    x = 1', '    y = 20', '    z = 3', 'End Sub', '', 'Sub B()', 'End Sub'].join('\n');
		expect(unifiedDiff(before, after, labels)).toBe([
			'--- Module1 (HEAD)',
			'+++ Module1 (current)',
			'@@ -2,7 +2,7 @@',
			' ',
			' Sub A()',
			'     x = 1',
			'-    y = 2',
			'+    y = 20',
			'     z = 3',
			' End Sub',
			' ',
		].join('\n'));
	});

	it('splits far-apart changes into separate hunks and merges near ones', () => {
		const before = Array.from({ length: 30 }, (_, i) => `l${i}`).join('\n');
		const far = before.replace('l2', 'L2').replace('l27', 'L27');
		expect(unifiedDiff(before, far, labels).match(/^@@/gm)).toHaveLength(2);
		const near = before.replace('l10', 'L10').replace('l14', 'L14');
		expect(unifiedDiff(before, near, labels).match(/^@@/gm)).toHaveLength(1);
	});

	it('handles an added and a removed module', () => {
		expect(unifiedDiff('', 'Sub A()\nEnd Sub\n', labels)).toBe([
			'--- Module1 (HEAD)',
			'+++ Module1 (current)',
			'@@ -0,0 +1,2 @@',
			'+Sub A()',
			'+End Sub',
		].join('\n'));
		expect(unifiedDiff('Sub A()\nEnd Sub\n', '', labels)).toBe([
			'--- Module1 (HEAD)',
			'+++ Module1 (current)',
			'@@ -1,2 +0,0 @@',
			'-Sub A()',
			'-End Sub',
		].join('\n'));
	});
});
