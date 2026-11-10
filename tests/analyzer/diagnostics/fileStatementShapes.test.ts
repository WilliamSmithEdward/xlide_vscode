// File statements in a shape the VBE refuses: an Open Access its mode doesn't
// take, Width without # or two arguments, Print # and Write # without a comma
// after the number, Input and Line Input without variables, Get without a
// variable and Put without a value, Name without As. Measured in Excel,
// Word, PowerPoint and Access 16.0 (build 20527) through the VBE oracle; the
// cases are the file_statement_shape_* fixtures in
// syntax_corpus/oracle/vbe_oracle_cases.json.

import * as fs from 'node:fs';
import * as path from 'node:path';
import { describe, expect, it } from 'vitest';
import { analyzeModule, diagnosticMetadataForCode } from '../../../src/analyzer';
import { analyzeVbaModuleSource } from '../../../src/vbaModuleAnalysis';
import { buildVbaProjectIndex, projectAnalysisOptionsForModule, projectProcedureSignatures } from '../../../src/vbaProjectAnalysis';
import { byCode, expectDiagnostic } from '../../helpers/diagnostics';

interface OracleCase {
	id: string;
	expected: 'accepted' | 'rejected';
	host: string;
	modules: { source: string; type: string }[];
}

const fixtures: OracleCase[] = (JSON.parse(
	fs.readFileSync(path.join(process.cwd(), 'syntax_corpus', 'oracle', 'vbe_oracle_cases.json'), 'utf8'),
) as { cases: OracleCase[] }).cases.filter((kase) => kase.id.startsWith('file_statement_shape_'));

const shapes = (source: string, host = 'excel') => byCode(analyzeModule(source, { host }), 'malformed-statement');
const compileErrors = (source: string, options: Record<string, unknown> = {}) => analyzeModule(source, { host: 'excel', ...options })
	.filter((diagnostic) => diagnosticMetadataForCode(diagnostic.code ?? '')?.diagnosticKind === 'compile-error')
	.map((diagnostic) => `${diagnostic.code}: ${diagnostic.message}`);

function procedure(...lines: string[]): string {
	return ['Option Explicit', 'Sub Main()', '    Dim f As Integer, s As String, n As Long, v As Variant, a(2) As Variant', '    f = FreeFile', ...lines.map((line) => '    ' + line), 'End Sub', ''].join('\r\n');
}

describe('file statements in a shape the VBE refuses', () => {
	it('covers every host', () => {
		expect(new Set(fixtures.map((kase) => kase.host))).toEqual(new Set(['excel', 'word', 'powerpoint', 'access']));
	});

	it.each(fixtures.map((kase) => [kase.id, kase] as const))('%s', (_id, kase) => {
		// Each module as its own kind: a class's or a form's statements are read too.
		const found = kase.modules.flatMap((mod) => byCode(analyzeModule(mod.source, { host: kase.host, ...(mod.type === 'standard' ? {} : { moduleKind: mod.type }) } as never), 'malformed-statement'));
		expect(found.length).toBe(kase.expected === 'rejected' ? 1 : 0);
	});

	it.each([
		['Open "x" For Append Access Read As #f', 'Access Read', 'For Append takes Access Write or Read Write.'],
		['Open "x" For Input Access Read Write As #f', 'Access Read Write', 'For Input takes Access Read.'],
		['Open "x" For Output Access Read As #f', 'Access Read', 'For Output takes Access Write.'],
		['Open "x" For Append Access As #f', 'Access', 'Access takes Read, Write or Read Write.'],
		['Width f, 80', 'Width f, 80', 'Width names its file number with #'],
		['Width #f, 80, 1', 'Width #f, 80, 1', 'Width takes a file number and a width'],
		['Print #f', 'Print #f', 'needs a comma after the file number'],
		['Print #f; 1', 'Print #f;', 'not a semicolon'],
		['Input #f, s & "x"', 'Input #f, s & "x"', 'reads into one or more variables'],
		['Line Input #f, "x"', 'Line Input #f, "x"', 'reads into one variable'],
		['Put #f, 1', 'Put #f, 1', 'a record and a value'],
		['Get #f, 1', 'Get #f, 1', 'a record and a variable'],
		['Get #f, , 1', '1', "Variable required - can't assign to this expression"],
		['Name "a" As', 'Name "a" As', 'Name renames with As'],
	])('reports %j', (line, span, message) => {
		const src = procedure(line);
		expectDiagnostic(src, shapes(src), 'malformed-statement', { span, message });
	});

	it('reads a variable by its shape, Name with a parenthesized path, Open without For, and every part of Width, Get and Put', () => {
		for (const line of ['Input #f, s And s', 'Input #f, s s', 'Name ("a") "b"', 'Name (s) As', 'Open "x" Access As #f', 'Width #, 80', 'Get #f, , v, v', 'Put #f, , v, v',
			'Name "a" = "b"', 'Name "a" = "b" As', 'Get #f, , -1', 'Get #f, , 1 + 2', 'Get #f, , s & "x"']) {
			expect(shapes(procedure(line)), line).toHaveLength(1);
		}
		for (const line of ['Input #f, s$', 'Line Input #f, s$', 'Input #f, a(1), v', 'Name ("a") As "b"', 'Open "x" Access Read As #f']) {
			expect(shapes(procedure('Dim b(2) As String', 'If False Then ' + line, 'If False Then Input #f, b$(1)')), line).toEqual([]);
		}
	});

	it('follows no run-time file state through a statement refused for its shape', () => {
		const opened = procedure('Open "x" For Input As #f', 'Name "x" As', 'Close #f');
		expect(analyzeModule(opened, { host: 'excel' }).map((d) => d.code)).toEqual(expect.not.arrayContaining(['file-already-open', 'file-path-open']));
		expect(shapes(opened)).toHaveLength(1);
	});

	it('reads a one-line If, a colon and a line number', () => {
		for (const line of ['If False Then Print #f', 'If False Then Beep Else Write #f', 'Beep: Input #f', '10 Name "a" "b"']) {
			expect(shapes(procedure(line)), line).toHaveLength(1);
		}
	});

	it('leaves the forms that compile alone', () => {
		for (const line of [
			'Open "x" For Binary Access Read Write Lock Read As #f Len = 128', 'Open "x" For Append Access Read Write As #f', 'Open "x" Access Read As #f',
			'Width #f, 80', 'Print #f,', 'Print #f, s; n', 'Write #f, s, n', 'Input #f, s, n, a(1)', 'Line Input #f, s',
			'Get #f, , v', 'Get f, 1, v', 'Put #f, , v', 'Put f, 1, a(1)', 'Name "a" As "b"', 'Name s & "x" As s', 'Debug.Print s',
		]) {
			expect(compileErrors(procedure('If False Then', '    ' + line, 'End If')), line).toEqual([]);
		}
	});

	it('leaves names that only look like the statements alone', () => {
		// A path that names a variable access has no Access clause.
		for (const open of ['Open access For Input As #f', 'Open (access) For Input As #f', 'Open "x" & access For Input As #f', 'Open CStr(Not access) For Input As #f',
			'Open CStr(Not access) As #f', 'Open "x" & Not access As #f', 'Open CStr(1 Mod access) Access Read As #f',
			...['<>', '<=', '>=', '=<', '=>', '><', '=', '<', '>', '&', '+'].map((op) => `Open "x" ${op} access As #f`)]) {
			expect(shapes(procedure('Dim access As String', 'If False Then ' + open)), open).toEqual([]);
		}
		// A member named access is part of the path.
		const members = ['Option Explicit', 'Private Type R', '    access As String', '    s As String', 'End Type', 'Sub Main()', '    Dim t As R, f As Integer',
			'    f = FreeFile', '    If False Then Open t.access For Input As #f', '    If False Then Open t!access Access Read As #f', '    If False Then Open "x" & t.access For Input As #f',
			'End Sub', ''].join('\r\n');
		expect(shapes(members)).toEqual([]);
		// A member straight after Width or Name, a label, or an assignment to the name itself is no such statement.
		expect(shapes(['Option Explicit', 'Sub Main()', '    Dim Name As String', '    Name$ = "x"', '    Name = "y"', 'End Sub', ''].join('\r\n'))).toEqual([]);
		expect(shapes(['Option Explicit', 'Sub Main(ByVal Width As Object, ByVal Name As Object)', '    Width.Add 1, 2', '    Width!k = 1', '    Width = 300', '    Name.Add 1', '    Name!k = 1',
			'Width:', 'Name:', 'End Sub', ''].join('\r\n'))).toEqual([]);
		expect(shapes(['Option Explicit', 'Sub Main()', '    Dim Width As Double', '    Width# = 300', 'End Sub', 'Sub Other()', '    Dim Width As Single', '    Width! = 1', 'End Sub', ''].join('\r\n'))).toEqual([]);
		// A With block's member is a variable to read into.
		const record = ['Option Explicit', 'Private Type R', '    s As String', 'End Type', 'Sub Main()', '    Dim t As R, f As Integer', '    f = FreeFile',
			'    With t', '        If False Then Input #f, .s', '        If False Then Line Input #f, .s', '    End With', 'End Sub', ''].join('\r\n');
		expect(shapes(record)).toEqual([]);
		// A form's Width and Name are its properties.
		const form = ['Option Explicit', 'Private Sub UserForm_Initialize()', '    Width = 300', '    Name = "x"', 'End Sub', ''].join('\r\n');
		expect(byCode(analyzeModule(form, { host: 'excel', moduleKind: 'userform' } as never), 'malformed-statement')).toEqual([]);
	});

	it('reads the statement wherever the VBE does, past an element list or a procedure named Width or Name', () => {
		const lines = ['Width f, 80 = 80', 'Width (f), 80', 'Width(1) = 2', 'Width(1).Add 1, 2', 'Width(1)(2).Add 1, 2', 'Width#(1) = 300',
			'Name(1).Add 1, 2', 'Name(1)(2).Add 1', 'Name(1) = "x"', 'Name(1).s = "x"', 'Name$(1) = "x"'];
		for (const line of lines) {
			expect(shapes(procedure(line)), line).toHaveLength(1);
		}
		const calls = ['Option Explicit', 'Sub Width(ByVal a As Long, ByVal b As Long)', 'End Sub', 'Sub Name(ByVal a As String, ByVal b As String)', 'End Sub', 'Sub Main()', '    Width 1, 80', '    Name "a", "b"', 'End Sub', ''].join('\r\n');
		expect(shapes(calls)).toHaveLength(2);
	});

	it('refuses Me.s as a place to read into', () => {
		const source = (line: string) => ['Option Explicit', 'Public s As String', 'Sub Load()', '    Dim f As Integer', '    f = FreeFile', `    If False Then ${line}`, 'End Sub', ''].join('\r\n');
		for (const line of ['Input #f, Me.s', 'Line Input #f, Me.s', 'Get #f, , Me.s']) {
			const found = byCode(analyzeModule(source(line), { host: 'excel', moduleKind: 'class' } as never), 'malformed-statement');
			expect(found.map((d) => d.message), line).toEqual([expect.stringContaining("Variable required - can't assign to this expression.")]);
		}
		// Another place that is no variable's shape makes it a Syntax error.
		const mixed = byCode(analyzeModule(source('Input #f, Me.s, 1'), { host: 'excel', moduleKind: 'class' } as never), 'malformed-statement');
		expect(mixed.map((d) => d.message)).toEqual([expect.stringContaining('Syntax error.')]);
	});

	it('reports no run-time file error on a statement refused for its shape', () => {
		const src = procedure('Open "x" For Input As #f', 'Print #f', 'Close #f');
		expect(analyzeModule(src, { host: 'excel' }).map((d) => d.code).filter((code) => code !== 'malformed-statement' && code?.startsWith('file-'))).toEqual([]);
		expect(shapes(src)).toHaveLength(1);
	});

	it('still knows the project\'s procedures when it follows the file states', () => {
		// OtherHelper may close and reopen the file, so nothing is known after it.
		// In the module's first procedure, and in its second, which main missed.
		const first = procedure('Open "x" For Output As #f', 'n = OtherHelper', 'Line Input #f, s', 'Close #f');
		for (const source of [first, first.replace('Sub Main()', 'Sub Other()\r\nEnd Sub\r\nSub Main()')]) {
			const modules = [
				{ moduleName: 'Module1', type: 'standard', source },
				{ moduleName: 'Module2', type: 'standard', source: 'Option Explicit\r\nPublic Function OtherHelper() As Long\r\nEnd Function\r\n' },
			];
			const project = buildVbaProjectIndex(modules, undefined, { conditionalCompilation: { projectConstants: {} } });
			const { diagnostics } = analyzeVbaModuleSource({
				source, moduleName: 'Module1', moduleKind: 'standard',
				...projectAnalysisOptionsForModule(project, 'Module1', projectProcedureSignatures(project)),
			});
			expect(diagnostics.map((d) => d.code)).not.toContain('file-mode-mismatch');
		}
	});
});
