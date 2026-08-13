import { describe, expect, it } from 'vitest';
import { VBA_IDENTIFIER_RE, VBA_IDENTIFIER_NAME_RE, VBA_IDENTIFIER_PATTERN } from '../src/vbaSourceScan';
import { detectSmartBlockOpener } from '../src/vbaSmartEnter';
import { ProjectIndex, resolveMemberCompletions } from '../src/analyzer';
import { identifierSpanEndingAt } from '../src/analyzer/completion/cursorContext';

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

describe('project structure resolves through non-ASCII type names', () => {
	function classPair(interfaceName: string, className: string, memberName: string): ProjectIndex {
		const index = new ProjectIndex();
		index.setModule({
			moduleName: interfaceName, moduleKind: 'class',
			source: `Option Explicit\n\nPublic Sub ${memberName}()\nEnd Sub\n`,
		});
		index.setModule({
			moduleName: className, moduleKind: 'class',
			source: `Option Explicit\n\nImplements ${interfaceName}\n\n`
				+ `Private Sub ${interfaceName}_${memberName}()\nEnd Sub\n`,
		});
		index.setModule({ moduleName: 'Uses', moduleKind: 'standard', source: '' });
		return index;
	}

	it.each([
		['Cyrillic', 'Фигура', 'Круг', 'Рисовать'],
		['Greek', 'Σχήμα', 'Κύκλος', 'Σχεδίαση'],
		['Latin control', 'IShape', 'RoundShape', 'Draw'],
	])('sees Implements with a %s interface name', (_label, iface, cls, member) => {
		const surfaces = classPair(iface, cls, member).projectMemberSurfaces('Uses');
		expect(surfaces.find((s) => s.name === cls)?.implements).toContain(iface);
	});

	it.each([
		['Cyrillic', 'Прибор', 'Проверить'],
		['Thai', 'อุปกรณ์', 'ตรวจสอบ'],
		['Latin control', 'Gadget', 'Describe'],
	])('offers members of a %s-named class receiver', (_label, className, memberName) => {
		const index = new ProjectIndex();
		index.setModule({
			moduleName: className, moduleKind: 'class',
			source: `Option Explicit\n\nPublic Sub ${memberName}()\nEnd Sub\n\nPublic Sub Describe()\nEnd Sub\n`,
		});
		index.setModule({ moduleName: 'Uses', moduleKind: 'standard', source: '' });
		const source = `Public Sub T()\n    Dim g As ${className}\n    g.\nEnd Sub\n`;
		const got = resolveMemberCompletions(source, source.indexOf('g.') + 2, {
			projectClassMembers: index.projectMemberSurfaces('Uses'),
		}).map((m) => m.name);
		// The whole receiver failed to resolve before, so even Describe vanished.
		expect(got).toContain(memberName);
		expect(got).toContain('Describe');
	});
});

describe('For loops whose iterator is not ASCII', () => {
	// detectSmartBlockOpener returned undefined for these, so pressing Enter
	// after the header closed nothing at all - the loop-completeness test was
	// ASCII-only, so a Cyrillic or Thai iterator made a valid header look
	// unfinished.
	it.each([
		['Cyrillic', 'товар'],
		['Thai', 'ค่า'],
		['Latin control', 'item'],
	])('closes For Each with a %s iterator', (_label, name) => {
		expect(detectSmartBlockOpener(`For Each ${name} In Items`)?.endKeyword)
			.toBe(`Next ${name}`);
	});

	it.each([
		['Cyrillic', 'товар'],
		['Thai', 'ค่า'],
		['Latin control', 'item'],
	])('closes a counted For with a %s iterator', (_label, name) => {
		expect(detectSmartBlockOpener(`For ${name} = 1 To 10`)?.endKeyword)
			.toBe(`Next ${name}`);
	});

	it('still treats an unfinished For header as unfinished', () => {
		expect(detectSmartBlockOpener('For Each')).toBeUndefined();
		expect(detectSmartBlockOpener('For x =')).toBeUndefined();
	});
});

describe('the identifier span under the cursor', () => {
	// Feeds completion's word range. It walked back over [A-Za-z0-9_] only, so
	// a partially typed non-Latin name had no span and completion had nothing
	// to filter on.
	it.each([
		['Cyrillic', 'Проверка'],
		['Thai (combining mark)', 'ค่า'],
		['Devanagari (matra)', 'नाम'],
		['Japanese', 'モジュール'],
		['Latin control', 'Recalculate'],
	])('covers the whole %s name', (_label, name) => {
		const source = `Public Sub T()\r\n    ${name}`;
		const span = identifierSpanEndingAt(source, source.length);
		expect(span).toBeDefined();
		expect(source.slice(span!.start, span!.end)).toBe(name);
	});

	it('does not start a span on a combining mark', () => {
		expect(identifierSpanEndingAt('    ่', 5)).toBeUndefined();
	});

	it('is undefined where there is no identifier', () => {
		expect(identifierSpanEndingAt('    = ', 6)).toBeUndefined();
	});
});
