import { expect, it } from 'vitest';
import { checkRuntimeMemberNotFound } from '../src/analyzer/diagnostics/rules/lateBoundMembers';
import { parseModule } from '../src/analyzer/parser/parseModule';
import { buildModuleSymbols } from '../src/analyzer/symbols/buildModuleSymbols';
import { statementTokensCached } from '../src/analyzer/lexer/tokenHelpers';
function fixture(body: string) {
	const source = ['Option Explicit', 'Sub Go()', 'Dim obj As Object', 'Set obj = New Collection', body, 'End Sub', ''].join('\n');
	const mod = parseModule(source), symbols = buildModuleSymbols('M', 'standard', source, { parsedModule: mod });
	return { source, mod, symbols };
}
function run(f: ReturnType<typeof fixture>) {
	const out: unknown[][] = [];
	checkRuntimeMemberNotFound(f.source, f.mod, f.symbols, { parsedModule: f.mod, memberSurfaceCache: new Map() }, undefined, (...v) => out.push(v)); return out;
}
function refused(f: ReturnType<typeof fixture>, reason: string, count = 1) {
	let from = 0;
	return Array.from({ length: count }, () => { const start = f.source.indexOf('obj.Remove', from) + 4; from = start + 6; return ['runtimeMemberNotFound', "'obj' holds a Collection here: " + reason + '.', { start, end: from }]; });
}
function tooMany(count: number) { return 'its Remove takes at most 1 argument(s), and ' + count + " are passed. This will raise Run-time error '450': Wrong number of arguments or invalid property assignment"; }
for (const branch of ['too-many', 'argument-group']) {
	it.each([10, 100, 1000])('does no comment-filter work for %i values, ' + branch, count => {
		const f = fixture('obj.Remove ' + (branch === 'too-many' ? Array(count).fill('1').join(', ') : 'Array(' + Array(count).fill('1').join(', ') + ')') + " ' trailing comment");
		const procedure = f.mod.members.find(m => m.kind === 'Procedure')!, statement = procedure.body.find(node => f.source.slice(node.span.start, node.span.end).includes('obj.Remove'))!;
		const tokens = statementTokensCached(f.source, statement.span), descriptors = tokens.map(t => Object.getOwnPropertyDescriptor(t, 'kind')!); let filterReads = 0;
		try {
			for (let i = 0; i < tokens.length; i++) { const token = tokens[i], kind = token.kind; Object.defineProperty(token, 'kind', { configurable: true, get() { const stack = new Error().stack ?? ''; if (stack.includes('at Array.filter') && stack.includes('at argumentRefusal')) filterReads++; return kind; } }); }
			expect(run(f)).toEqual(branch === 'too-many' ? refused(f, tooMany(count)) : []); expect(filterReads).toBe(0);
		} finally { for (let i = 0; i < tokens.length; i++) Object.defineProperty(tokens[i], 'kind', descriptors[i]); }
	});
}
it.each(["obj.Remove 1, 2 ' comment", 'obj.Remove 1, _\n2', 'obj.Remove Array(1, 2), 3', 'obj.Remove(1, 2)'])('preserves two-argument grouping: %s', body => {
	const f = fixture(body); expect(run(f)).toEqual(refused(f, tooMany(2)));
});
it('preserves colon-separated call spans', () => { const f = fixture('obj.Remove 1, 2: obj.Remove 1, 2'); expect(run(f)).toEqual(refused(f, tooMany(2), 2)); });
it('preserves missing and unknown named arguments', () => {
	let f = fixture('obj.Remove'); expect(run(f)).toEqual(refused(f, "its Remove needs 'Index', which is not passed. This will raise Run-time error '449': Argument not optional"));
	f = fixture('obj.Remove Other:=1'); expect(run(f)).toEqual(refused(f, "its Remove has no parameter named 'Other'. This will raise Run-time error '448': Named argument not found"));
});
it.each(['obj.Remove Index:=1', 'obj.Remove "a\'b"', 'obj.Remove Array(1, (2 + 3))'])('preserves valid bare argument tokens: %s', body => { expect(run(fixture(body))).toEqual([]); });
it('preserves an omitted positional slot', () => { const f = fixture('obj.Remove ,'); expect(run(f)).toEqual(refused(f, tooMany(2))); });
it('preserves labelled call spans', () => { const f = fixture('Retry: obj.Remove 1, 2'); expect(run(f)).toEqual(refused(f, tooMany(2))); });
