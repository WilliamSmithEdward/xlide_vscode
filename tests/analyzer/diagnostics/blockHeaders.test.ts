// A block header without one of its parts: Case Is with no comparison
// operator, Do While, Do Until, Loop While, Loop Until or While with no
// condition, For Each without its
// variable, In or collection, and For without To, its end value or Step's
// value. Each is a Syntax error, measured in Excel, Word, PowerPoint and
// Access 16.0 (build 20527) through the VBE oracle; the cases are the
// block_header_* fixtures in syntax_corpus/oracle/vbe_oracle_cases.json.

import * as fs from 'node:fs';
import * as path from 'node:path';
import { describe, expect, it } from 'vitest';
import { analyzeModule, diagnosticMetadataForCode } from '../../../src/analyzer';
import { byCode, expectDiagnostic } from '../../helpers/diagnostics';

interface OracleCase {
	id: string;
	expected: 'accepted' | 'rejected';
	host: string;
	modules: { source: string }[];
}

const fixtures: OracleCase[] = (JSON.parse(
	fs.readFileSync(path.join(process.cwd(), 'syntax_corpus', 'oracle', 'vbe_oracle_cases.json'), 'utf8'),
) as { cases: OracleCase[] }).cases.filter((kase) => kase.id.startsWith('block_header_'));

const headers = (source: string, host = 'excel') => byCode(analyzeModule(source, { host }), 'malformed-statement');
const compileErrors = (source: string) => analyzeModule(source, { host: 'excel' })
	.filter((diagnostic) => diagnosticMetadataForCode(diagnostic.code ?? '')?.diagnosticKind === 'compile-error')
	.map((diagnostic) => `${diagnostic.code}: ${diagnostic.message}`);

function procedure(...lines: string[]): string {
	return ['Option Explicit', 'Sub Main()', '    Dim v As Variant, n As Long, o As Object, c As New Collection', ...lines.map((line) => '    ' + line), 'End Sub', ''].join('\r\n');
}

describe('a block header without one of its parts', () => {
	it('covers every host', () => {
		expect(new Set(fixtures.map((kase) => kase.host))).toEqual(new Set(['excel', 'word', 'powerpoint', 'access']));
	});

	// `10 Case Is > o.To` is refused for the line number before the first Case,
	// which another check reports; its member o.To is no To of a range.
	it.each(fixtures.map((kase) => [kase.id, kase] as const))('%s', (id, kase) => {
		const source = kase.modules[0].source;
		const all = analyzeModule(source, { host: kase.host }).filter((d) => diagnosticMetadataForCode(d.code ?? '')?.diagnosticKind === 'compile-error');
		expect(all.length > 0).toBe(kase.expected === 'rejected');
		const own = kase.expected === 'rejected' && !id.includes('_numbered_case_member_to_') ? 1 : 0;
		expect(headers(source, kase.host).length).toBe(own);
	});

	it.each([
		[['Select Case n', 'Case Is 5', 'End Select'], 'Is 5', 'Case Is takes a comparison operator before its value'],
		[['Select Case n', 'Case 1, Is', 'End Select'], 'Is', 'Case Is takes a comparison operator'],
		[['Do Until', 'Loop'], 'Do Until', 'Do Until needs a condition.'],
		[['Do', 'Loop While'], 'Loop While', 'Loop While needs a condition.'],
		[['While', 'Wend'], 'While', 'While needs a condition'],
		[['For Each v', 'Next'], 'For Each v', 'needs In and what it walks'],
		[['For Each v In', 'Next'], 'For Each v In', 'needs the collection or array it walks'],
		[['For Each In c', 'Next'], 'For Each In', 'its loop variable before In'],
		[['For n = 1', 'Next'], 'For n = 1', 'needs To and its end value'],
		[['For n = 1 To', 'Next'], 'For n = 1 To', "needs the loop's end value"],
		[['For n = 1 To 2 Step', 'Next'], 'For n = 1 To 2 Step', 'Step needs a value'],
	])('reports %j', (lines, span, message) => {
		const src = procedure(...lines);
		expectDiagnostic(src, headers(src), 'malformed-statement', { span, message });
	});

	it('reads a header after a line number or a label', () => {
		for (const lines of [['10 For n = 1', 'Next'], ['Select Case n', '10 Case Is 5', 'End Select'], ['10 Do Until', 'Loop'], ['10 While', 'Wend'], ['L: For Each v', 'Next']]) {
			expect(headers(procedure(...lines)), lines.join(' / ')).toHaveLength(1);
		}
	});

	it('reads a member named Until, To or In as a member, not the keyword', () => {
		expect(headers(procedure('For n = o.To To', 'Next'))).toHaveLength(1);
		for (const lines of [['10 Do', '20 Loop While o.Until'], ['For n = o.To To 3', 'Next'], ['For Each v In o.In', 'Next'], ['Do Until o!While', 'Loop'],
			['Select Case n', '10 Case Is > o.To', 'End Select'], ['Select Case n', 'Case Is > o!To', 'End Select']]) {
			expect(headers(procedure(...lines)), lines.join(' / ')).toEqual([]);
		}
	});

	it('leaves the forms that compile alone', () => {
		for (const lines of [
			['Select Case n', 'Case Is = 1, Is <> 2, Is < 3, Is > 4, Is <= 5, Is >= 6', 'End Select'],
			['Select Case n', 'Case 1 To 3, Is > 7: Beep', 'End Select'],
			['Select Case n', 'Case Is _', '    > 1', 'End Select'],
			['Do Until n > 3', '    n = n + 1', 'Loop'], ['Do', '    n = n + 1', 'Loop Until (n > 3)'], ['Do: Loop While n < 0'],
			['While (n < 3)', '    n = n + 1', 'Wend'],
			['For Each v In c', 'Next'], ['For Each v In Array(1, 2)', 'Next v'], ['For Each v _', '    In c', 'Next'],
			['For n = 1 To 3', 'Next'], ['For n = LBound(Array(1)) To UBound(Array(1, 2)) Step -1', 'Next n'], ['For n = 1 To _', '    3', 'Next'],
			['10 Do While n < 3', '    n = n + 1', '20 Loop'], ['10 For n = 1 To 2', '20 Next'],
		]) {
			expect(compileErrors(procedure(...lines)), lines.join(' / ')).toEqual([]);
		}
	});

	it('leaves an inactive #If arm alone', () => {
		expect(headers(procedure('#If False Then', 'For Each v', 'Next', 'Do Until', 'Loop', '#End If'))).toEqual([]);
	});
});
