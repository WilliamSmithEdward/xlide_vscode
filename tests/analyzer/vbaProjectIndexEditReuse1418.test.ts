import { describe, expect, it } from 'vitest';
import { ProjectIndex } from '../../src/analyzer/symbols/projectIndex';
import { analyzeVbaModuleSource } from '../../src/vbaModuleAnalysis';
import { projectAnalysisOptionsForModule, projectProcedureSignatures } from '../../src/vbaProjectAnalysis';

// An edit to one module must not rebuild what every other module gave the
// index (issue #1418): class member surfaces, a class's PredeclaredId and
// the exported procedure signatures are kept per module until that module
// changes. These pin both halves: the unchanged modules' parts are the same
// objects after the edit, and the edited module's parts follow its text.

const classSource = (members: string[], predeclared = false): string => [
	'VERSION 1.0 CLASS',
	'BEGIN',
	'  MultiUse = -1  \'True',
	'END',
	`Attribute VB_PredeclaredId = ${predeclared ? 'True' : 'False'}`,
	'Option Explicit',
	...members.flatMap((name) => [`Public Function ${name}(ByVal a As Long) As Long`, `    ${name} = a`, 'End Function']),
	'',
].join('\r\n');

const standardSource = (procedures: string[]): string => [
	'Option Explicit',
	...procedures.flatMap((name) => [`Public Function ${name}(ByVal a As Long) As Long`, `    ${name} = a`, 'End Function']),
	'',
].join('\r\n');

const USER = [
	'Option Explicit',
	'Sub Main()',
	'    Dim c As Widget, n As Long',
	'    Set c = New Widget',
	'    n = c.Spin(1) + Helper(2)',
	'    n = c.Missing(1)',
	'    n = Helper(1, 2)',
	'End Sub',
	'',
].join('\r\n');

function project(): ProjectIndex {
	const index = new ProjectIndex();
	index.setModule({ moduleName: 'Widget', moduleKind: 'class', source: classSource(['Spin']) });
	index.setModule({ moduleName: 'Gadget', moduleKind: 'class', source: classSource(['Turn'], true) });
	index.setModule({ moduleName: 'Tools', moduleKind: 'standard', source: standardSource(['Helper']) });
	index.setModule({ moduleName: 'Other', moduleKind: 'standard', source: standardSource(['Unrelated']) });
	index.setModule({ moduleName: 'User', moduleKind: 'standard', source: USER });
	return index;
}

const surface = (index: ProjectIndex, name: string) => index.projectClassMembers().find((owner) => owner.name === name);

function diagnostics(index: ProjectIndex) {
	return analyzeVbaModuleSource({
		source: USER,
		moduleName: 'User',
		moduleKind: 'standard',
		host: 'excel',
		...projectAnalysisOptionsForModule(index, 'User', projectProcedureSignatures(index)),
	}).diagnostics.map((d) => ({ code: d.code, start: d.span.start, end: d.span.end, message: d.message }));
}

describe('project index parts kept across edits elsewhere (#1418)', () => {
	it('keeps every class surface when a standard module changes', () => {
		const index = project();
		const widget = surface(index, 'Widget');
		const gadget = surface(index, 'Gadget');
		index.setModule({ moduleName: 'Other', moduleKind: 'standard', source: standardSource(['Unrelated', 'Added']) });
		expect(surface(index, 'Widget')).toBe(widget);
		expect(surface(index, 'Gadget')).toBe(gadget);
	});

	it('rebuilds only the edited class, which follows its new text', () => {
		const index = project();
		const gadget = surface(index, 'Gadget');
		index.setModule({ moduleName: 'Widget', moduleKind: 'class', source: classSource(['Spin', 'Missing'], true) });
		expect(surface(index, 'Gadget')).toBe(gadget);
		const widget = surface(index, 'Widget');
		expect(widget?.members.map((member) => member.name)).toEqual(['Spin', 'Missing']);
		expect(widget?.predeclaredId).toBe(true);
		expect(index.modulePredeclaredId('Widget')).toBe(true);
		expect(index.modulePredeclaredId('Gadget')).toBe(true);
	});

	it('keeps the class value facts and the editor surfaces apart', () => {
		const index = project();
		const withFacts = surface(index, 'Widget');
		const withoutFacts = index.projectClassMembers({ includeClassValueFacts: false }).find((owner) => owner.name === 'Widget');
		expect(withoutFacts).not.toBe(withFacts);
		expect(withoutFacts?.members).not.toBe(withFacts?.members);
	});

	it('keeps an unchanged module\'s signatures, and follows the edited one', () => {
		const index = project();
		const helper = index.procedureSignatures().get('helper')?.[0];
		index.setModule({ moduleName: 'Other', moduleKind: 'standard', source: standardSource(['Unrelated', 'Added']) });
		const after = index.procedureSignatures();
		expect(after.get('helper')?.[0]).toBe(helper);
		expect(after.get('added')?.[0]?.name).toBe('Added');
		expect(after.get('other.added')?.[0]?.name).toBe('Added');
		index.setModule({ moduleName: 'Tools', moduleKind: 'standard', source: standardSource([]) });
		expect(index.procedureSignatures().has('helper')).toBe(false);
	});

	it('reports the same findings, at the same spans, after an edit elsewhere', () => {
		const index = project();
		const before = diagnostics(index);
		expect(before.map((d) => d.code)).toEqual(expect.arrayContaining(['member-not-found']));
		index.setModule({ moduleName: 'Other', moduleKind: 'standard', source: standardSource(['Unrelated', 'Added']) });
		expect(diagnostics(index)).toEqual(before);
		index.setModule({ moduleName: 'Widget', moduleKind: 'class', source: classSource(['Spin', 'Missing']) });
		expect(diagnostics(index).map((d) => d.code)).not.toContain('member-not-found');
	});
});
