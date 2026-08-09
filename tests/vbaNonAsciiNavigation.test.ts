import { describe, expect, it } from 'vitest';
import { VBA_IDENTIFIER_RE, VBA_IDENTIFIER_NAME_RE, VBA_IDENTIFIER_PATTERN } from '../src/vbaSourceScan';
import { detectSmartBlockOpener } from '../src/vbaSmartEnter';

// The same bug class as issues #6 and #8, one layer out from the analyzer.
// VS Code picks the word under the cursor with VBA_IDENTIFIER_RE, and it
// matched none of a Cyrillic, Greek, Thai or Japanese name - so go to
// definition, find references, prepare rename and rename all did nothing at
// all on those identifiers, and smart Enter did not close their blocks.
const NAMES: Array<[string, string]> = [
	['Cyrillic', 'Проверка'],
	['Greek', 'Δοκιμή'],
	['Thai (combining mark)', 'ค่า'],
	['Devanagari (matra)', 'नाम'],
	['Japanese', 'モジュール'],
	['Chinese', '测试模块'],
	['Latin control', 'Recalculate'],
];

describe('the word under the cursor on a non-ASCII identifier', () => {
	it.each(NAMES)('selects the whole %s name', (_label, name) => {
		expect(VBA_IDENTIFIER_RE.exec(name)?.[0]).toBe(name);
	});

	it.each(NAMES)('accepts %s as a declarable name', (_label, name) => {
		expect(VBA_IDENTIFIER_NAME_RE.test(name)).toBe(true);
	});

	it('still rejects what is not an identifier', () => {
		expect(VBA_IDENTIFIER_NAME_RE.test('1Bad')).toBe(false);
		expect(VBA_IDENTIFIER_NAME_RE.test('has space')).toBe(false);
		// A bare combining mark cannot start a name.
		expect(VBA_IDENTIFIER_NAME_RE.test('่')).toBe(false);
	});
});

describe('smart Enter closes a block whose name is not ASCII', () => {
	it.each(NAMES)('detects a Sub opener named in %s', (_label, name) => {
		expect(detectSmartBlockOpener(`Public Sub ${name}()`)?.endKeyword).toBe('End Sub');
	});

	it.each(NAMES)('detects a Function opener named in %s', (_label, name) => {
		expect(detectSmartBlockOpener(`Public Function ${name}() As Long`)?.endKeyword)
			.toBe('End Function');
	});
});

describe('the identifier pattern is only safe with the u flag', () => {
	// \p{L} in a regex built without `u` is read as a literal 'p{L}' and
	// silently matches nothing - the failure mode is an empty result, not an
	// error, so it has to be pinned rather than trusted.
	it('matches a non-ASCII name when built with u', () => {
		const withU = new RegExp(`^${VBA_IDENTIFIER_PATTERN}$`, 'u');
		expect(withU.test('Проверка')).toBe(true);
	});

	it('is documented as requiring u, and fails silently without it', () => {
		const withoutU = new RegExp(`^${VBA_IDENTIFIER_PATTERN}$`);
		expect(withoutU.test('Проверка')).toBe(false);
		expect(withoutU.test('Recalculate')).toBe(false);
	});
});
