import { describe, expect, it } from 'vitest';
import { getWordObjectModel } from '../src/analyzer/host/wordObjectModel';
import { getPowerPointObjectModel } from '../src/analyzer/host/powerpointObjectModel';
import { getAccessObjectModel } from '../src/analyzer/host/accessObjectModel';
import { getExcelObjectModel, type HostObjectModel } from '../src/analyzer/host/excelObjectModel';
import { resolveHostGlobal, getHostMembers, resolveHostConstant } from '../src/analyzer/host/hostModel';
import { resolveMemberCompletions } from '../src/analyzer';

// Issue #25. The measured number this replaces: Word's ThisDocument answered
// 1 member - the Sub in its own module - because the only model that existed
// was Excel's and the honest alternative was silence.

const MODELS: ReadonlyArray<[string, () => HostObjectModel]> = [
	['word', getWordObjectModel],
	['powerpoint', getPowerPointObjectModel],
	['access', getAccessObjectModel],
];

// Origin labels are built from hostName (issue #28), so a misspelled name
// would ship in every hover and completion detail of that host.
it('every model names its application for origin labels', () => {
	expect(getExcelObjectModel().hostName).toBe('Excel');
	expect(getWordObjectModel().hostName).toBe('Word');
	expect(getPowerPointObjectModel().hostName).toBe('PowerPoint');
	expect(getAccessObjectModel().hostName).toBe('Access');
});

describe.each(MODELS)('the %s object model', (_host, getModel) => {
	it('declares every global against a type it actually carries', () => {
		const model = getModel();
		for (const [name, qualified] of Object.entries(model.globals)) {
			expect(model.types[qualified] ?? model.aliases[qualified.toLowerCase()], `${name} -> ${qualified}`)
				.toBeDefined();
		}
	});

	it('proves nothing absent: no type claims exhaustiveness', () => {
		const model = getModel();
		for (const [qualified, type] of Object.entries(model.types)) {
			expect(type.exhaustive ?? false, qualified).toBe(false);
		}
	});

	it('carries documentation, not bare signatures', () => {
		const model = getModel();
		const documented = Object.values(model.types)
			.flatMap((type) => type.members)
			.filter((member) => member.doc?.summary).length;
		expect(documented).toBeGreaterThan(500);
	});
});

describe('Word answers as Word', () => {
	it('ThisDocument is a Document with the everyday members', () => {
		const model = getWordObjectModel();
		const documentType = resolveHostGlobal('ThisDocument', model);
		expect(documentType).toBe('Word.Document');
		const names = getHostMembers(documentType!, model).map((member) => member.name);
		for (const expected of ['Content', 'Paragraphs', 'SaveAs2', 'Range', 'Tables']) {
			expect(names, expected).toContain(expected);
		}
		// The issue's measured "1" becomes the real surface.
		expect(names.length).toBeGreaterThan(200);
	});

	it('does not leak Excel: no Cells on a Document, no Worksheets global', () => {
		const model = getWordObjectModel();
		const names = getHostMembers('Word.Document', model).map((member) => member.name);
		expect(names).not.toContain('Cells');
		expect(resolveHostGlobal('ThisWorkbook', model)).toBeUndefined();
	});

	it('Selection. completes through the member resolver', () => {
		const source = 'Sub T()\r\n    Selection.\r\nEnd Sub\r\n';
		const members = resolveMemberCompletions(source, source.indexOf('Selection.') + 10, {
			model: getWordObjectModel(),
		} as never).map((member) => member.name);
		expect(members).toContain('TypeText');
		expect(members).toContain('Collapse');
	});

	it('knows wd* constants with their values', () => {
		const constant = resolveHostConstant('wdMainTextStory', getWordObjectModel());
		expect(constant?.value).toBe(1);
		expect(constant?.type).toBe('WdStoryType');
	});
});

describe('PowerPoint and Access answer as themselves', () => {
	it('ActivePresentation reaches Slides', () => {
		const model = getPowerPointObjectModel();
		const type = resolveHostGlobal('ActivePresentation', model);
		expect(type).toBe('PowerPoint.Presentation');
		expect(getHostMembers(type!, model).map((member) => member.name)).toContain('Slides');
	});

	it('DoCmd carries OpenForm, documented', () => {
		const model = getAccessObjectModel();
		const type = resolveHostGlobal('DoCmd', model);
		expect(type).toBe('Access.DoCmd');
		const openForm = getHostMembers(type!, model).find((member) => member.name === 'OpenForm');
		expect(openForm).toBeDefined();
		expect(openForm?.kind).toBe('method');
	});

	it('the hosts do not bleed into each other or into Excel', () => {
		expect(resolveHostGlobal('ActivePresentation', getWordObjectModel())).toBeUndefined();
		expect(resolveHostGlobal('DoCmd', getPowerPointObjectModel())).toBeUndefined();
		expect(resolveHostGlobal('ActiveDocument', getExcelObjectModel())).toBeUndefined();
	});
});
