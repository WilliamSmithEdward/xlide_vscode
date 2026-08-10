import { describe, expect, it } from 'vitest';
import { interfacePrefixHits, renameInterfacePrefixes } from '../src/vbaInterfacePrefix';

// Issue #9 rule 4: `Private Sub IShape_Draw()` binds the implementation to the
// interface named by `Implements IShape`. The prefix is a contract - rename the
// interface and leave it behind and the class stops implementing anything.
const IMPL = [
	'Option Explicit',
	'',
	'Implements IShape',
	'',
	'Private Sub IShape_Draw()',
	'    Debug.Print "drawing"',
	'End Sub',
	'',
	'Private Function IShape_Area() As Double',
	'    IShape_Area = 1',
	'End Function',
	'',
	'Private Property Get IShape_Name() As String',
	'    IShape_Name = "circle"',
	'End Property',
	'',
].join('\r\n');

describe('the interface prefix on a declaration', () => {
	it('finds Sub, Function and Property declarations', () => {
		const hits = interfacePrefixHits(IMPL, 'IShape');
		expect(hits).toHaveLength(3);
		for (const hit of hits) {
			// The span covers the interface name only, not the whole procedure name.
			const line = IMPL.split('\r\n')[hit.line];
			expect(line.slice(hit.column, hit.column + hit.length)).toBe('IShape');
		}
	});

	it('renames the prefix and nothing else on the line', () => {
		const renamed = renameInterfacePrefixes(IMPL, 'IShape', 'IDrawable');
		expect(renamed).toContain('Private Sub IDrawable_Draw()');
		expect(renamed).toContain('Private Function IDrawable_Area() As Double');
		expect(renamed).toContain('Private Property Get IDrawable_Name() As String');
		// `Implements` is a type reference and is renamed by the type pass, not here.
		expect(renamed).toContain('Implements IShape');
	});

	it.each([
		['a variable that merely starts with the name', '    Dim IShapeLookalike As Long'],
		['a string', '    Debug.Print "IShape_Draw"'],
		['a comment', "    ' IShape_Draw is the contract"],
		['a call rather than a declaration', '    IShape_Draw'],
		['a procedure named exactly after the interface', 'Private Sub IShape()'],
	])('leaves %s alone', (_label, line) => {
		expect(interfacePrefixHits(line, 'IShape')).toEqual([]);
	});

	it('matches the prefix case-insensitively, as VBA does', () => {
		expect(interfacePrefixHits('Private Sub ishape_Draw()', 'IShape')).toHaveLength(1);
	});

	it('handles an interface name that is not Latin', () => {
		const source = 'Private Sub Фигура_Рисовать()\r\nEnd Sub\r\n';
		expect(renameInterfacePrefixes(source, 'Фигура', 'Контур'))
			.toContain('Private Sub Контур_Рисовать()');
	});
});
