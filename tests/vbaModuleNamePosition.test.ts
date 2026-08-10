import { describe, expect, it } from 'vitest';
import { moduleNamePositionKind } from '../src/vbaModuleNamePosition';

// Issue #9 rule 3. The caller supplies "nothing else resolved here"; this
// answers "and it stands where only a module can".
function kindOf(source: string, word: string) {
	const start = source.indexOf(word);
	return moduleNamePositionKind(source, start, start + word.length);
}

describe('positions only a module name can occupy', () => {
	it('before a member dot', () => {
		expect(kindOf('    Helpers.Recalculate\r\n', 'Helpers')).toBe('qualifier');
	});

	it('after As', () => {
		expect(kindOf('    Dim s As IShape\r\n', 'IShape')).toBe('as');
	});

	it('after New', () => {
		expect(kindOf('    Set s = New RoundShape\r\n', 'RoundShape')).toBe('new');
	});

	it('after Implements', () => {
		expect(kindOf('Implements IShape\r\n', 'IShape')).toBe('implements');
	});

	it('across a line continuation before the dot', () => {
		// Rule 6: a continuation is whitespace, so this is one qualified call.
		expect(kindOf('    Helpers _\r\n        .Recalculate\r\n', 'Helpers')).toBe('qualifier');
	});

	it('across a line continuation after As', () => {
		expect(kindOf('    Dim s As _\r\n        IShape\r\n', 'IShape')).toBe('as');
	});
});

describe('positions that are not a module name', () => {
	it.each([
		['a bare call', '    Recalculate\r\n', 'Recalculate'],
		['an assignment target', '    Config = 1\r\n', 'Config'],
		['a declaration name', '    Dim Log As Long\r\n', 'Log'],
		['the member after the dot', '    Helpers.Recalculate\r\n', 'Recalculate'],
		['an argument', '    Debug.Print Data\r\n', 'Data'],
	])('%s', (_label, source, word) => {
		expect(kindOf(source, word)).toBeUndefined();
	});

	it('does not treat a decimal point as a member dot', () => {
		expect(kindOf('    x = Rate\r\n', 'Rate')).toBeUndefined();
	});

	it('is undefined for an out-of-range span', () => {
		expect(moduleNamePositionKind('abc', 5, 9)).toBeUndefined();
		expect(moduleNamePositionKind('abc', 2, 1)).toBeUndefined();
	});
});
