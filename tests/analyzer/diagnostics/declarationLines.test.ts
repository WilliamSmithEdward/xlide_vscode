// DefDec, which is no Def statement, and a Declare without Lib after its
// name. Measured in Excel, Word, PowerPoint and Access 16.0 (build 20527)
// through the VBE oracle; the cases are the declaration_line_* fixtures in
// syntax_corpus/oracle/vbe_oracle_cases.json.

import * as fs from 'node:fs';
import * as path from 'node:path';
import { describe, expect, it } from 'vitest';
import { analyzeModule, diagnosticMetadataForCode } from '../../../src/analyzer';
import { expectDiagnostic } from '../../helpers/diagnostics';

interface OracleCase {
	id: string;
	expected: 'accepted' | 'rejected';
	host: string;
	modules: { source: string }[];
}

const fixtures: OracleCase[] = (JSON.parse(
	fs.readFileSync(path.join(process.cwd(), 'syntax_corpus', 'oracle', 'vbe_oracle_cases.json'), 'utf8'),
) as { cases: OracleCase[] }).cases.filter((kase) => kase.id.startsWith('declaration_line_'));

const compileErrors = (source: string, host = 'excel') => analyzeModule(source, { host })
	.filter((diagnostic) => diagnosticMetadataForCode(diagnostic.code ?? '')?.diagnosticKind === 'compile-error');

const module = (...lines: string[]) => ['Option Explicit', ...lines, ''].join('\r\n');

describe('DefDec and a Declare without Lib', () => {
	it('covers every host', () => {
		expect(new Set(fixtures.map((kase) => kase.host))).toEqual(new Set(['excel', 'word', 'powerpoint', 'access']));
	});

	it.each(fixtures.map((kase) => [kase.id, kase] as const))('%s: one report, or none', (_id, kase) => {
		// A Def statement in a one-line If also gets its letters' Variable not defined, as on main.
		const found = compileErrors(kase.modules[0].source, kase.host).filter((d) => d.code !== 'undeclared-variable');
		expect(found.length).toBe(kase.expected === 'rejected' ? 1 : 0);
	});

	it.each([
		[['DefDec A-Z', 'Sub Main()', 'End Sub'], 'DefDec', 'Expected: line number or label or statement or end of statement'],
		[['Sub Main()', 'End Sub', 'DefDec A-Z'], 'DefDec', 'Syntax error'],
		[['Sub Main()', '    DefDec A-Z', 'End Sub'], 'DefDec', 'Syntax error'],
		[['Private Declare PtrSafe Function F () As Long', 'Sub Main()', 'End Sub'], '(', 'Expected: Lib'],
		[['Private Declare PtrSafe Function F Alias "GetTickCount" Lib "kernel32" () As Long', 'Sub Main()', 'End Sub'], 'Alias', 'Expected: Lib'],
		[['Private Declare PtrSafe Sub S', 'Sub Main()', 'End Sub'], 'S', 'Expected: Lib'],
	])('reports %j', (lines, span, message) => {
		const src = module(...lines);
		const found = compileErrors(src);
		expect(found).toHaveLength(1);
		expectDiagnostic(src, found, found[0].code!, { span, message });
	});

	it('leaves the forms that compile alone', () => {
		for (const lines of [
			['DefInt I', 'DefLngLng L', 'DefCur C'],
			['Private Declare PtrSafe Function F Lib "kernel32" Alias "GetTickCount" () As Long'],
			['Private Declare PtrSafe Function F& Lib "kernel32" Alias "GetTickCount" ()'],
			['Private Declare PtrSafe Function F CDecl Lib "kernel32" Alias "GetTickCount" () As Long'],
			['Private Declare PtrSafe Function F _', '    Lib "kernel32" Alias "GetTickCount" () As Long'],
			['Private Declare PtrSafe Sub S Lib "kernel32" Alias "Sleep" (ByVal ms As Long)'],
		]) {
			expect(compileErrors(module(...lines, 'Sub Main()', 'End Sub')).map((d) => d.message), lines.join(' / ')).toEqual([]);
		}
	});

	it('reads a digit-start Declare name whole, so only the invalid-name check reports it', () => {
		for (const name of ['9Sleep', ...['$', '%', '&', '!', '#', '@', '^'].map((suffix) => `9Sleep${suffix}`)]) {
			const src = module(`Private Declare PtrSafe Function ${name} Lib "kernel32" Alias "GetTickCount" ()`, 'Sub Main()', 'End Sub');
			expect(analyzeModule(src, { host: 'excel' }).filter((d) => d.code === 'invalid-proc-header').map((d) => d.message), name).toEqual([]);
		}
	});

	it('gives DefDec no letters, so a later Def statement may take them', () => {
		const messages = compileErrors(module('DefDec A', 'DefInt A', 'Sub Main()', 'End Sub')).map((d) => d.message);
		expect(messages).toHaveLength(1);
		expect(messages[0]).toContain("'DefDec' is a reserved word");
	});

	it('reads DefDec after a line number, in a one-line If, and above the procedures of its own #If arm', () => {
		const message = (...lines: string[]) => compileErrors(module(...lines)).filter((d) => d.message.includes("'DefDec' is a reserved word and no statement")).map((d) => d.message);
		expect(message('Sub Main()', '10 DefDec:', 'End Sub')).toEqual([expect.stringContaining('Syntax error')]);
		expect(message('Sub Main()', '    If False Then DefDec A-Z', 'End Sub')).toEqual([expect.stringContaining('Syntax error')]);
		expect(message('Sub Main()', '    If True Then Beep Else DefDec A', 'End Sub')).toEqual([expect.stringContaining('Syntax error')]);
		expect(message('#If False Then', 'Sub Other()', 'End Sub', '#Else', 'DefDec A-Z', '#End If', 'Sub Main()', 'End Sub'))
			.toEqual([expect.stringContaining('Expected: line number or label or statement or end of statement')]);
		expect(message('#If False Then', 'Sub Other()', 'End Sub', '#End If', 'DefDec A-Z', 'Sub Main()', 'End Sub'))
			.toEqual([expect.stringContaining('Expected: line number or label or statement or end of statement')]);
		for (const line of ['    Beep: DefDec:', 'foo: DefDec:', '    If True Then Beep: DefDec:', '    DefDec _\r\n:', ' _\r\nDefDec:']) {
			expect(message('Sub Main()', line, 'End Sub'), line).toEqual([expect.stringContaining('Syntax error')]);
		}
	});

	it('reports a DefDec label outside the procedures, which the reserved-label check doesn\'t read', () => {
		const message = (...lines: string[]) => compileErrors(module(...lines)).map((d) => d.message);
		expect(message('DefDec:', 'Sub Main()', 'End Sub')).toEqual([expect.stringContaining("'DefDec' is a reserved word and no statement")]);
		expect(message('Sub Main()', 'End Sub', 'DefDec:')).toEqual([expect.stringContaining('Syntax error')]);
	});

	it('gives a DefDec label one report with CR-only line ends too', () => {
		for (const src of ['Option Explicit\rSub Main()\rDefDec:\rEnd Sub\r', 'Option Explicit\rSub Main()\r\tDefDec:\rEnd Sub\r', 'Option Explicit\nSub Main()\nDefDec:\nEnd Sub\n',
			'Option Explicit\rSub Main()\r _\rDefDec:\rEnd Sub\r', 'Option Explicit\rSub Main()\r _  \rDefDec:\rEnd Sub\r', 'Option Explicit\r\nSub Main()\r\n _\r\nDefDec:\r\nEnd Sub\r\n']) {
			expect(compileErrors(src).filter((d) => d.message.includes('DefDec')), JSON.stringify(src)).toHaveLength(1);
		}
	});

	it('leaves a DefDec label to the reserved-label check', () => {
		const messages = compileErrors(module('Function Main() As Variant', 'DefDec:', '    Main = 1', 'End Function')).map((d) => d.message);
		expect(messages).toEqual(["'DefDec' is a reserved word, so 'DefDec:' cannot be a line label. This is a VBE compile error: Syntax error."]);
	});
});
