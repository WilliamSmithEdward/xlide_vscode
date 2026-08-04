import { describe, expect, it } from 'vitest';
import { tokenize } from '../src/analyzer/lexer/tokenize';
import { analyzeVbaModuleSource } from '../src/vbaModuleAnalysis';
import { matchOpener } from '../src/vbaStructuralDiagnostics';

// Thai and Devanagari write a single letter as a base plus a tone mark, vowel
// sign or matra. Those marks are Unicode Mn/Mc, not L, so identifier patterns
// built on \p{L} alone stop dead at the mark and split the name in half. The
// VBE compiles and runs these identifiers (oracle-verified against Excel 16),
// so every split here is a false positive on valid VBA.
//
// A mark is only ever a continuation character: at position 0 it has nothing
// to combine with, and VBA does not accept it as an identifier start.
const THAI = 'ค่า';            // kho khwai + mai ek (Mn) + sara aa
const THAI_PLAIN = 'ราคา';     // mark-free Thai control
const DEVANAGARI = 'नाम';      // na + aa matra (Mc) + ma

function identifiersIn(source: string): string[] {
	return tokenize(source).filter((token) => token.kind === 'identifier').map((token) => token.rawText);
}

function moduleUsing(name: string): string {
	return [
		'Option Explicit',
		'',
		'Public Sub Probe()',
		`    Dim ${name} As String`,
		`    ${name} = "ok"`,
		'End Sub',
		'',
	].join('\r\n');
}

describe('identifiers containing combining marks', () => {
	it.each([
		['Thai (Mn tone mark)', THAI],
		['Devanagari (Mc matra)', DEVANAGARI],
		['Thai without marks', THAI_PLAIN],
	])('lexes %s as one identifier', (_label, name) => {
		expect(identifiersIn(moduleUsing(name))).toEqual(['Probe', name, name]);
	});

	it.each([
		['Thai (Mn tone mark)', THAI],
		['Devanagari (Mc matra)', DEVANAGARI],
	])('reports no diagnostics for a declared %s variable', (_label, name) => {
		const result = analyzeVbaModuleSource({
			source: moduleUsing(name),
			moduleName: 'Module1',
			moduleType: 'standard',
			moduleKind: 'standard',
			// The undeclared-variable rule only runs with project context; the
			// split name is what made it fire on a variable that IS declared.
			knownIdentifiers: new Set<string>(),
		});
		expect(result.diagnostics).toEqual([]);
	});

	it('still refuses a bare combining mark as an identifier start', () => {
		// U+0E48 alone is not a name: it has nothing to combine with.
		const identifiers = identifiersIn('Option Explicit\r\nPublic Sub P()\r\n    Dim ่ As String\r\nEnd Sub\r\n');
		expect(identifiers).toEqual(['P']);
	});

	it.each([
		['Thai', THAI],
		['Devanagari', DEVANAGARI],
	])('captures the whole %s name as a block opener', (_label, name) => {
		// The opener still matched on the base letter, so no false "no matching
		// Sub" - but the name it captured was only the first half.
		const opener = matchOpener(`Public Sub ${name}()`);
		expect(opener?.kind).toBe('Sub');
		expect(opener?.label).toContain(name);
	});

	it('reports no diagnostics for a mark-carrying procedure', () => {
		const source = ['Option Explicit', '', `Public Sub ${THAI}()`,
			'    Debug.Print "ok"', 'End Sub', ''].join('\r\n');
		expect(analyzeVbaModuleSource({
			source, moduleName: 'Module1', moduleType: 'standard', moduleKind: 'standard',
		}).diagnostics).toEqual([]);
	});
});
