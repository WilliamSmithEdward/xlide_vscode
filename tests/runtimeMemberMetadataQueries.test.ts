import { expect, it } from 'vitest';
import { checkRuntimeMemberNotFound } from '../src/analyzer/diagnostics/rules/lateBoundMembers';
import { parseModule } from '../src/analyzer/parser/parseModule';
import { buildModuleSymbols } from '../src/analyzer/symbols/buildModuleSymbols';
import { analyzeModule } from '../src/analyzer/diagnostics/analyzeModule';
import type { HostMember, HostObjectModel } from '../src/analyzer/host/excelObjectModel';
import type { VbaProjectClassMembers } from '../src/analyzer/symbols/symbolModel';

const messages = {
	worksheetFunction: "WorksheetFunction has no function 'Missing'. The VBE compiles the name; this will raise Run-time error '438': Object doesn't support this property or method.",
	activeSheet: "ActiveSheet has no member 'Missing': neither a Worksheet nor a Chart has one, and no document module of the project declares it. This will raise Run-time error '438': Object doesn't support this property or method.",
	formControls: 'The form Form1 has no control named "Missing". This will raise Run-time error \'-2147024809\': Could not find the specified object.',
};
function fixture(scope: keyof typeof messages, count: number) {
	let names = 0, returns = 0;
	const items: HostMember[] = Array.from({ length: count }, (_, i) => ({ get name() { names++; return 'M' + i; }, kind: 'method' }));
	const model: HostObjectModel = { hostName: 'Excel', source: 'query work model', types: {
		'Excel.Application': { displayName: 'Application', exhaustive: true, members: [{ name: 'WorksheetFunction', kind: 'property', returns: 'Excel.WorksheetFunction' }] },
		'Excel.WorksheetFunction': { displayName: 'WorksheetFunction', members: scope === 'worksheetFunction' ? items : [] },
		'Excel.Worksheet': { displayName: 'Worksheet', members: scope === 'activeSheet' ? items : [] },
		'Excel.Chart': { displayName: 'Chart', members: [] },
	}, aliases: { worksheet: 'Excel.Worksheet', application: 'Excel.Application' }, globals: { WorksheetFunction: 'Excel.WorksheetFunction', ActiveSheet: 'union:Excel.Worksheet|Excel.Chart' } };
	const classes: VbaProjectClassMembers[] = scope === 'formControls' ? [{ name: 'Form1', moduleName: 'Form1', kind: 'userform', exhaustive: true,
		members: Array.from({ length: count }, (_, i) => ({ name: 'Control' + i, moduleName: 'Form1', kind: 'property', get returns() { returns++; return 'MSForms.TextBox'; } })) }] : [];
	const line = scope === 'worksheetFunction' ? 'Debug.Print WorksheetFunction.Missing' : scope === 'activeSheet' ? 'Debug.Print ActiveSheet.Missing' : 'Debug.Print actor.Controls("Missing")';
	const source = ['Option Explicit', 'Sub Go()', ...(scope === 'formControls' ? ['Dim actor As New Form1'] : []), ...Array(count).fill(line), 'End Sub', ''].join('\n');
	return { model, classes, source, reads: () => names + returns };
}
function run(source: string, model: HostObjectModel, classes: VbaProjectClassMembers[], parsed = parseModule(source)) {
	const out: unknown[][] = [], symbols = buildModuleSymbols('M', 'standard', source, { parsedModule: parsed });
	checkRuntimeMemberNotFound(source, parsed, symbols, { model, projectClassMembers: classes, parsedModule: parsed, memberSurfaceCache: new Map() }, undefined, (...v) => out.push(v));
	return out;
}
for (const scope of Object.keys(messages) as (keyof typeof messages)[]) for (const api of ['public', 'full']) {
	it.each([10, 100, 1000])('bounds ' + scope + ' metadata for %i references, ' + api, count => {
		const f = fixture(scope, count); let from = 0;
		const expected = Array.from({ length: count }, () => { const marker = scope === 'formControls' ? '"Missing"' : 'Missing', start = f.source.indexOf(marker, from); from = start + marker.length; return [api === 'public' ? 'runtimeMemberNotFound' : 'runtime-member-not-found', messages[scope], { start, end: from }]; });
		const errors: unknown[] = [];
		const actual = api === 'public' ? run(f.source, f.model, f.classes) : analyzeModule(f.source, { hostModel: f.model, projectClassMembers: f.classes, onInternalError: e => errors.push(e) }).map(d => [d.code, d.message, d.span]);
		expect(actual).toEqual(scope === 'formControls' ? [] : expected); expect(errors).toEqual([]); expect(f.reads()).toBeLessThanOrEqual(count * 5 + 20);
	});
}
it.each(['worksheetFunction', 'activeSheet'] as const)('preserves known %s member and casing', scope => {
	const f = fixture(scope, 3); expect(run(f.source.replaceAll('Missing', 'm1'), f.model, f.classes)).toEqual([]);
});
it('preserves the Excel application gate', () => {
	const f = fixture('worksheetFunction', 3);
	expect(run(f.source, { ...f.model, hostName: 'Word' }, f.classes)).toEqual([]);
	const application = f.model.types['Excel.Application'];
	expect(run(f.source, { ...f.model, types: { ...f.model.types, 'Excel.Application': { ...application, exhaustive: false } } }, f.classes)).toEqual([]);
});
it.each(['class', 'document', 'userform', 'standardModule'] as const)('preserves %s document surface authority', kind => {
	const f = fixture('activeSheet', 3), type: VbaProjectClassMembers = { name: 'Sheet1', moduleName: 'Sheet1', kind, members: [{ name: 'Missing', moduleName: 'Sheet1', kind: 'property' }] };
	expect(run(f.source, f.model, [type])).toHaveLength(kind === 'document' ? 0 : 3);
});
it('preserves shadowed ActiveSheet bindings', () => {
	const f = fixture('activeSheet', 3); expect(run(f.source.replace('Sub Go()', 'Sub Go()\nDim ActiveSheet As Object'), f.model, f.classes)).toEqual([]);
});
it('does not infer runtime Controls from designer names or scalar fields', () => {
	const f = fixture('formControls', 3); expect(run(f.source.replaceAll('"Missing"', '"cOnTrOl1"'), f.model, f.classes)).toEqual([]);
	f.classes[0].members.push({ name: 'Missing', moduleName: 'Form1', kind: 'property', returns: 'Long' });
	expect(run(f.source, f.model, f.classes)).toEqual([]);
});
it.each([0, 1, 2])('does not infer runtime control count from duplicate designer names with index %i', index => {
	const f = fixture('formControls', 2); f.classes[0].members[1].name = 'Control0';
	const text = f.source.replaceAll('"Missing"', String(index)); const out = run(text, f.model, f.classes);
	if (index < 2) expect(out).toEqual([]);
	else expect(out).toEqual([]);
});
it('preserves incomplete form and dynamic added-control guards', () => {
	const f = fixture('formControls', 3); f.classes[0].exhaustive = false; expect(run(f.source, f.model, f.classes)).toEqual([]); f.classes[0].exhaustive = true;
	const named = f.source.replace('Dim actor As New Form1', 'Dim actor As New Form1\nactor.Controls.Add "Forms.TextBox.1", "Missing"'); expect(run(named, f.model, f.classes)).toEqual([]);
	const unknown = f.source.replace('Dim actor As New Form1', 'Dim actor As New Form1\nDim newName As String\nactor.Controls.Add "Forms.TextBox.1", newName'); expect(run(unknown, f.model, f.classes)).toEqual([]);
});
it('keeps designer metadata separate from runtime Controls across invocations', () => {
	const f = fixture('formControls', 3), parsed = parseModule(f.source);
	expect(run(f.source, f.model, f.classes, parsed)).toEqual([]);
	f.classes[0].members.push({ name: 'Missing', moduleName: 'Form1', kind: 'property', returns: 'MSForms.TextBox' });
	expect(run(f.source, f.model, f.classes, parsed)).toEqual([]);
	f.classes[0].members.pop(); expect(run(f.source, f.model, f.classes, parsed)).toEqual([]);
});
it('refreshes host queries with new models and a retained AST', () => {
	for (const scope of ['worksheetFunction', 'activeSheet'] as const) {
		const f = fixture(scope, 3), parsed = parseModule(f.source); expect(run(f.source, f.model, f.classes, parsed)).toHaveLength(3);
		const key = scope === 'worksheetFunction' ? 'Excel.WorksheetFunction' : 'Excel.Worksheet', model = { ...f.model, types: { ...f.model.types, [key]: { ...f.model.types[key], members: [{ name: 'Missing', kind: 'method' as const }] } } };
		expect(run(f.source, model, f.classes, parsed)).toEqual([]); expect(run(f.source, f.model, f.classes, parsed)).toHaveLength(3);
	}
});
it.each(['worksheetFunction', 'activeSheet', 'formControls'] as const)('shares %s queries across procedures within one invocation', scope => {
	const count = 100, f = fixture(scope, count), line = f.source.split('\n').find(value => value.startsWith('Debug.Print'))!;
	const text = ['Option Explicit', ...Array.from({ length: count }, (_, i) => ['Sub P' + i + '()', ...(scope === 'formControls' ? ['Dim actor As New Form1'] : []), line, 'End Sub'].join('\n')), ''].join('\n');
	let from = 0;
	const expected = Array.from({ length: count }, () => { const marker = scope === 'formControls' ? '"Missing"' : 'Missing', start = text.indexOf(marker, from); from = start + marker.length; return ['runtimeMemberNotFound', messages[scope], { start, end: from }]; });
	expect(run(text, f.model, f.classes)).toEqual(scope === 'formControls' ? [] : expected); expect(f.reads()).toBeLessThanOrEqual(count * 5 + 20);
});
