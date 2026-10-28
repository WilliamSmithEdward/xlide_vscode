// Block shapes Format Document, Smart Enter and the analyzer misread (issue
// #1496). Every module here compiles in Excel, Word, PowerPoint and Access
// 16.0 (build 20527): the VBE oracle fixtures i1496_<host>_<family>.

import { describe, expect, it } from 'vitest';
import { formatVbaModule } from '../../src/analyzer/format/formatModule';
import { analyzeVbaModuleSource } from '../../src/vbaModuleAnalysis';
import { isSmartBlockClosedAhead, procedureHeaderParensEdit } from '../../src/vbaSmartEnter';
import { lexerStrippedLine } from '../../src/analyzer/lexer/strippedLines';
import { diagnosticMetadataForCode } from '../../src/analyzer/diagnostics/ruleMetadata';

const SPACES = { tabSize: 4, insertSpaces: true };

function format(...source: string[]): string[] {
	const result = formatVbaModule(source.join('\n'), SPACES);
	expect(result.refusal).toBeUndefined();
	return (result.text as string).split('\n');
}

// Compile-time findings only: the oracle compiled these modules, it did not run them.
const errors = (...source: string[]): string[] => analyzeVbaModuleSource({
	source: ['Option Explicit', ...source, ''].join('\n'), moduleName: 'Module1', moduleType: 'standard',
}).diagnostics.filter((d) => d.severity === 'error' && diagnosticMetadataForCode(d.code)?.diagnosticKind !== 'deterministic-runtime-error').map((d) => `${d.code}: ${d.message}`);

describe('a procedure header in each #If arm (issue #1496)', () => {
	it('keeps the procedure open after #End If when every arm opens it', () => {
		expect(format('#If VBA7 Then', 'Public Sub Foo(ByVal p As LongPtr)', '#Else', 'Public Sub Foo(ByVal p As Long)', '#End If', 'Dim x As Long', 'x = 1', 'End Sub', 'Sub Main()', 'End Sub')).toEqual([
			'#If VBA7 Then', '    Public Sub Foo(ByVal p As LongPtr)', '#Else', '    Public Sub Foo(ByVal p As Long)', '#End If',
			'    Dim x As Long', '    x = 1', 'End Sub', 'Sub Main()', 'End Sub',
		]);
	});

	it('keeps it open when only the first arm opens it', () => {
		expect(format('#If VBA7 Then', 'Public Sub Foo(ByVal p As LongPtr)', '#Else', '#End If', 'x = 1', 'End Sub')).toEqual([
			'#If VBA7 Then', '    Public Sub Foo(ByVal p As LongPtr)', '#Else', '#End If', '    x = 1', 'End Sub',
		]);
	});

	it('still closes what a one-arm #If opened', () => {
		expect(format('#If FLAG Then', 'Sub First()', '#End If', 'x = 1')).toEqual(['#If FLAG Then', '    Sub First()', '#End If', 'x = 1']);
	});

	it('reports no error on either shape', () => {
		expect(errors('#If VBA7 Then', 'Public Sub Foo(ByVal p As LongPtr)', '#Else', 'Public Sub Foo(ByVal p As Long)', '#End If', '    Dim x As Long', '    x = 1', 'End Sub')).toEqual([]);
		expect(errors('#If VBA7 Then', 'Public Sub Foo(ByVal p As LongPtr)', '#Else', '#End If', '    Dim x As Long', '    x = 1', 'End Sub')).toEqual([]);
	});
});

describe('Type and Enum members named like keywords (issue #1496)', () => {
	it('indents For, Sub and Next members as members', () => {
		expect(format('Private Type Rec', 'For As Long', 'Sub As Long', 'Next As Long', 'End Type', 'Dim x As Long')).toEqual([
			'Private Type Rec', '    For As Long', '    Sub As Long', '    Next As Long', 'End Type', 'Dim x As Long',
		]);
	});

	it('treats an array member and an Enum value the same way', () => {
		expect(format('Type Rec', 'Do(3) As Long', 'End Type', 'Enum E', 'With = 1', 'End Enum', 'Dim x As Long')).toEqual([
			'Type Rec', '    Do(3) As Long', 'End Type', 'Enum E', '    With = 1', 'End Enum', 'Dim x As Long',
		]);
	});
});

describe('one-line If tails and Else after a colon (issue #1496)', () => {
	it('opens and closes a With in the tail without closing the outer With', () => {
		expect(format('Sub A()', 'With c', 'If a Then With d: .Add 1: End With', '.Add 2', 'End With', 'End Sub')).toEqual([
			'Sub A()', '    With c', '        If a Then With d: .Add 1: End With', '        .Add 2', '    End With', 'End Sub',
		]);
	});

	it('reads an Else after a colon as the one-line If\'s own', () => {
		expect(format('Sub A()', 'If outer Then', 'With c', 'If a Then b = 1: Else b = 2', '.Add 1', 'End With', 'End If', 'End Sub')).toEqual([
			'Sub A()', '    If outer Then', '        With c', '            If a Then b = 1: Else b = 2', '            .Add 1', '        End With', '    End If', 'End Sub',
		]);
	});

	it('opens a block in that Else', () => {
		expect(format('Sub A()', 'With c', 'If a Then b = 1: Else With e: .Add 3: End With', '.Add 1', 'End With', 'End Sub')).toEqual([
			'Sub A()', '    With c', '        If a Then b = 1: Else With e: .Add 3: End With', '        .Add 1', '    End With', 'End Sub',
		]);
	});

	it('keeps the colon after a line number\'s one-line If', () => {
		expect(format('Sub A()', '10 If x Then: y = 1', 'z = 2', 'End Sub')).toEqual(['Sub A()', '10 If x Then: y = 1', '    z = 2', 'End Sub']);
	});

	it('reads a member named Then or Else as a member', () => {
		expect(format('Sub A()', 'With c', 'If a Then With d.Else: End With', '.Add 2', 'End With', 'End Sub')).toEqual([
			'Sub A()', '    With c', '        If a Then With d.Else: End With', '        .Add 2', '    End With', 'End Sub',
		]);
		expect(format('Sub A()', 'If a Then Set x = d.Then', 'y = 1', 'End Sub')).toEqual(['Sub A()', '    If a Then Set x = d.Then', '    y = 1', 'End Sub']);
		expect(format('Sub A()', 'If a Then Set x = d!Then', 'y = 1', 'End Sub')).toEqual(['Sub A()', '    If a Then Set x = d!Then', '    y = 1', 'End Sub']);
	});

	it('reads a type suffix before Then as a suffix', () => {
		expect(format('Sub A()', 'If x! Then With c: .Add 1: End With', 'y = 1', 'End Sub')).toEqual([
			'Sub A()', '    If x! Then With c: .Add 1: End With', '    y = 1', 'End Sub',
		]);
	});

	it('reports no error on any of them', () => {
		expect(errors('Public Sub Main()', '    Dim c As New Collection, d As New Collection, a As Boolean', '    With c', '        If a Then With d: .Add 1: End With', '        .Add 2', '    End With', 'End Sub')).toEqual([]);
		expect(errors('Public Sub Main()', '    Dim c As New Collection, e As New Collection, a As Boolean, outer As Boolean, b As Long', '    If outer Then', '        With c', '            If a Then b = 1: Else With e: .Add 3: End With', '            .Add 1', '        End With', '    End If', 'End Sub')).toEqual([]);
		expect(errors('Public Sub Main()', '    Dim x As Boolean, y As Long, z As Long', '10  If x Then: y = 1', '    z = 2', 'End Sub')).toEqual([]);
		expect(errors('Public Sub Main()', '    Dim c As New Collection, d As Object, a As Boolean', '    With c', '        If a Then With d.Else: End With', '        .Add 2', '    End With', 'End Sub')).toEqual([]);
		expect(errors('Public Sub Main()', '    Dim d As Object, x As Object, a As Boolean, y As Long', '    If a Then Set x = d.Then', '    y = 1', 'End Sub')).toEqual([]);
		expect(errors('Public Sub Main()', '    Dim c As New Collection, x!', '    If x! Then With c: .Add 1: End With', 'End Sub')).toEqual([]);
	});

	it('still reports a second Then before the block-opening one', () => {
		expect(errors('Sub Main()', '    If True Then Then', '    End If', 'End Sub').some((e) => e.startsWith('if-reserved-keyword-in-condition'))).toBe(true);
	});
});

describe('Static before the access word in Smart Enter (issue #1496)', () => {
	it('adds the parentheses to a Static-first header', () => {
		for (const line of ['Static Public Sub Foo', 'Static Private Function Bar As Long', 'Static Friend Property Get P']) {
			expect(procedureHeaderParensEdit(line), line).toEqual({ startCol: line.indexOf(' As') < 0 ? line.length : line.indexOf(' As'), endCol: line.indexOf(' As') < 0 ? line.length : line.indexOf(' As'), newText: '()' });
		}
	});

	it('stops the closer search at a Static-first header', () => {
		const module = ['Sub A()', '', 'Static Public Sub B()', 'End Sub'].map((line) => lexerStrippedLine(line));
		expect(isSmartBlockClosedAhead(module, 0, { endKeyword: 'End Sub' } as never)).toBe(false);
	});
});
