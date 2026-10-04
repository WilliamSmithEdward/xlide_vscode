import {afterEach, expect, it, vi} from 'vitest';
import {resolveDiagnosticCodeActions} from '../src/analyzer/codeActions/diagnosticCodeActions';

afterEach(() => vi.restoreAllMocks());
const boundaryPattern = /^(?:Public|Private|Friend|Static)?\s*(?:Sub|Function|Property\s+(?:Get|Let|Set))\b/i;

it.each([10, 1000].flatMap(count => ['\n', '\r\n', '\r'].map(eol => ({count, eol}))))('bounds context scans past $count unrelated procedures with EOL $eol', ({count, eol}) => {
	const prefix = Array.from({length: count}, (_, i) => ['Sub P' + i + '()', 'Dim n As Long', 'End Sub'].join(eol)).join(eol) + eol;
	const source = prefix + ['Sub Target()', '  Dim value As Long = 42', 'End Sub'].join(eol), offset = source.lastIndexOf('=');
	let reads = 0;
	const original = RegExp.prototype.test;
	vi.spyOn(RegExp.prototype, 'test').mockImplementation(function(this: RegExp, value: string) {
		if (this.source === boundaryPattern.source) { reads++; }
		return original.call(this, value);
	});
	expect(resolveDiagnosticCodeActions(source, {code: 'dim-initializer', span: {start: offset, end: offset + 1}})).toEqual([{
		title: 'Split declaration initializer', kind: 'quickfix', isPreferred: true,
		edits: [{span: {start: offset - 1, end: source.indexOf(eol, offset)}, newText: (eol === '\r' ? '\n' : eol) + '  value = 42'}],
	}]);
	expect(reads).toBe(1);
});

it('keeps the closest boundary decisive for malformed and module-level declarations', () => {
	for (const [prefix, inside] of [['Sub First()\nEnd Sub\n', false], ['Sub First()\nFunction Later()\n', true], ['Function First()\nEnd Property\n', false], ['\ufeffPrivate Property Get Value() As Long\n', true], ['Sub First()\n\n\n', true]] as const) {
		const source = prefix + 'Dim value As Long = 1\n', offset = source.indexOf('=');
		const actions = resolveDiagnosticCodeActions(source, {code: 'dim-initializer', span: {start: offset, end: offset + 1}});
		expect(actions.length).toBe(inside ? 1 : 0);
	}
});

it('retains unsafe initializer refusals and refreshes procedure context after edits', () => {
	for (const declaration of ['Dim value As Long = 1: Stop', 'Dim value As Long = 1, other As Long', "Dim value As Long = 1 ' comment"]) {
		const source = 'Sub S()\n' + declaration + '\nEnd Sub', offset = source.indexOf('=');
		expect(resolveDiagnosticCodeActions(source, {code: 'dim-initializer', span: {start: offset, end: offset + 1}})).toEqual([]);
	}
	for (const prefix of ['Sub S()\n', 'Sub S()\nEnd Sub\n', 'Sub S()\n']) {
		const source = prefix + 'Dim value As Long = 1', offset = source.indexOf('=');
		expect(resolveDiagnosticCodeActions(source, {code: 'dim-initializer', span: {start: offset, end: offset + 1}}).length).toBe(prefix.includes('End Sub') ? 0 : 1);
	}
});
