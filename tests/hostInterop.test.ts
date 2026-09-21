import { describe, expect, it } from 'vitest';
import * as fs from 'fs';
import * as path from 'path';
import { listReferences, readModulesFromBuffer } from '../src/vba/projectService';
import { analyzeModule } from '../src/analyzer/diagnostics/analyzeModule';
import { resolveMemberCompletions } from '../src/analyzer/completion/memberAccess';
import {
	hostTokenForLibid,
	hostTokensForProject,
	libraryGuidOf,
} from '../src/analyzer/host/hostLibraries';
import { hostObjectModelForTokens } from '../src/analyzer/host/hostRegistry';

// A project can reference another application's library and name its types:
// a Word document with a reference to Excel compiles `Dim xl As
// Excel.Application`, and `xl.Calculate` is Excel's method, not Word's. The
// library GUIDs and the fact that a library's global object is bound bare for
// anyone referencing it were both read from the registered type libraries on
// a machine with Office 16.

const LIBID = {
	excel: '*\\G{00020813-0000-0000-C000-000000000046}#1.9#0#C:\\PROGRA~1\\MICROS~1\\Office16\\EXCEL.EXE#Microsoft Excel 16.0 Object Library',
	word: '*\\G{00020905-0000-0000-C000-000000000046}#8.7#0#C:\\PROGRA~1\\MICROS~1\\Office16\\MSWORD.OLB#Microsoft Word 16.0 Object Library',
	powerpoint: '*\\G{91493440-5A91-11CF-8700-00AA0060263B}#2.12#0#C:\\PROGRA~1\\MICROS~1\\Office16\\MSPPT.OLB#Microsoft PowerPoint 16.0 Object Library',
	access: '*\\G{4AFFC9A0-5F99-101B-AF4E-00AA003F0F07}#9.0#0#C:\\PROGRA~1\\MICROS~1\\Office16\\MSACC.OLB#Microsoft Access 16.0 Object Library',
	office: '*\\G{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}#2.8#0#C:\\Program Files\\Common Files\\Microsoft Shared\\OFFICE16\\MSO.DLL#Microsoft Office 16.0 Object Library',
	stdole: '*\\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\\Windows\\System32\\stdole2.tlb#OLE Automation',
};

const reference = (libid: string, name = 'Lib') => ({ name, libid });

describe('reading a reference', () => {
	it('names the host of every library XLIDE models', () => {
		expect(hostTokenForLibid(LIBID.excel)).toBe('excel');
		expect(hostTokenForLibid(LIBID.word)).toBe('word');
		expect(hostTokenForLibid(LIBID.powerpoint)).toBe('powerpoint');
		expect(hostTokenForLibid(LIBID.access)).toBe('access');
	});

	it('says nothing for a library it has no model for', () => {
		// Silence is the honest answer, the same as for an unmodelled host.
		expect(hostTokenForLibid(LIBID.office)).toBeUndefined();
		expect(hostTokenForLibid(LIBID.stdole)).toBeUndefined();
		expect(hostTokenForLibid('*\\CNormal')).toBeUndefined();
	});

	it('reads the GUID and not the path, which is only a hint', () => {
		// The host resolves a library through the registry by GUID, which is
		// why a file written on one machine loads on another.
		const moved = LIBID.excel.replace('C:\\PROGRA~1\\MICROS~1\\Office16\\EXCEL.EXE', 'D:\\Somewhere\\Else.exe');

		expect(hostTokenForLibid(moved)).toBe('excel');
		expect(libraryGuidOf(moved)).toBe('{00020813-0000-0000-C000-000000000046}');
	});
});

describe('the hosts a project resolves against', () => {
	it('puts the project own host first, then each reference in order', () => {
		expect(hostTokensForProject('word', [
			reference(LIBID.stdole), reference(LIBID.office), reference(LIBID.excel),
		])).toEqual(['word', 'excel']);
	});

	it('keeps the declared order, since VBA resolves an ambiguous name by it', () => {
		expect(hostTokensForProject('access', [reference(LIBID.excel), reference(LIBID.word)]))
			.toEqual(['access', 'excel', 'word']);
		expect(hostTokensForProject('access', [reference(LIBID.word), reference(LIBID.excel)]))
			.toEqual(['access', 'word', 'excel']);
	});

	it('names the host once when the project references its own library', () => {
		// Every project implicitly references its own host, and some declare
		// it as well; it is one library either way.
		expect(hostTokensForProject('excel', [reference(LIBID.excel)])).toEqual(['excel']);
	});

	it('answers with the references alone when the host is unknown', () => {
		expect(hostTokensForProject(undefined, [reference(LIBID.excel)])).toEqual(['excel']);
		expect(hostTokensForProject(undefined, [])).toEqual([]);
	});
});

describe('the merged model', () => {
	it('gives a single host exactly the model it had before', () => {
		// Excel answers undefined so the downstream default keeps riding.
		expect(hostObjectModelForTokens(['excel'])).toBeUndefined();
		expect(hostObjectModelForTokens(['word'])?.hostName).toBe('Word');
	});

	it('keeps the project own host as the one it reports itself to be', () => {
		expect(hostObjectModelForTokens(['word', 'excel'])?.hostName).toBe('Word');
		expect(hostObjectModelForTokens(['word', 'excel'])?.globalType).toBe('Word.Global');
	});

	it("carries both libraries' types, constants and enums", () => {
		const merged = hostObjectModelForTokens(['word', 'excel'])!;

		expect(merged.types['Excel.Application']).toBeDefined();
		expect(merged.types['Word.Document']).toBeDefined();
		expect(merged.constants?.xlOpenXMLWorkbook).toBeDefined();
		expect(merged.constants?.wdFormatPDF).toBeDefined();
	});

	it('lets the project own host win a name both libraries use', () => {
		const wordFirst = hostObjectModelForTokens(['word', 'excel'])!;
		const excelFirst = hostObjectModelForTokens(['excel', 'word'])!;

		expect(wordFirst.aliases.range).toBe('Word.Range');
		expect(excelFirst.aliases.range).toBe('Excel.Range');
	});

	it('is built once per combination, so the analyzer indexes it once', () => {
		expect(hostObjectModelForTokens(['word', 'excel']))
			.toBe(hostObjectModelForTokens(['word', 'excel']));
	});
});

describe('a real document that references another application', () => {
	// Word 16 saved WordExcelInteropFixture.docm with a reference to the Excel
	// library and a Bridge module that drives Excel early-bound. Only Word can
	// write a real reference record, so the fixture is authored by it; the
	// engine reads it natively from here on.
	const modules = readModulesFromBuffer(
		fs.readFileSync(path.join(__dirname, 'fixtures', 'binaries', 'WordExcelInteropFixture.docm')),
		true,
	);
	const bridge = modules.find((one) => one.name === 'Bridge')!;

	it('reads the reference out of the project, alongside the module source', () => {
		const names = (bridge.projectReferences ?? []).map((one) => one.name);

		expect(names).toContain('Excel');
		expect(bridge.source).toContain('Dim xl As Excel.Application');
	});

	it('resolves the document own host and the referenced one, in that order', () => {
		expect(hostTokensForProject('word', bridge.projectReferences ?? []))
			.toEqual(['word', 'excel']);
	});

	it('analyzes its cross-application module clean, and only with the reference', () => {
		const tokens = hostTokensForProject('word', bridge.projectReferences ?? []);
		const options = { moduleName: 'Bridge', moduleKind: 'standard' as const };
		const withReference = analyzeModule(bridge.source!, {
			...options, hostModel: hostObjectModelForTokens(tokens),
		});
		const wordAlone = analyzeModule(bridge.source!, {
			...options, hostModel: hostObjectModelForTokens(['word']),
		});

		// Real cross-application code, so nothing should be reported for it.
		expect(withReference.map((d) => `${d.code}: ${d.message}`)).toEqual([]);
		// And Word alone simply has no opinion, which is why it is silent too.
		expect(wordAlone).toEqual([]);
	});

	it('catches a typo in the referenced library, which Word alone cannot', () => {
		const typo = bridge.source!.replace('wb.SaveAs', 'wb.SaveAsx');
		const tokens = hostTokensForProject('word', bridge.projectReferences ?? []);
		const options = { moduleName: 'Bridge', moduleKind: 'standard' as const };

		expect(analyzeModule(typo, { ...options, hostModel: hostObjectModelForTokens(tokens) })
			.map((d) => d.code)).toContain('member-not-found');
		expect(analyzeModule(typo, { ...options, hostModel: hostObjectModelForTokens(['word']) }))
			.toEqual([]);
	});
});

describe('listing a project references', () => {
	const fixture = path.join(__dirname, 'fixtures', 'binaries', 'WordExcelInteropFixture.docm');

	it('reports each one as Tools > References shows it', () => {
		const refs = listReferences(fixture);
		const excel = refs.find((one) => one.name === 'Excel')!;

		expect(refs.map((one) => one.name)).toEqual(['stdole', 'Normal', 'Office', 'Excel']);
		expect(excel.kind).toBe('registered');
		expect(excel.libid).toContain('{00020813-0000-0000-C000-000000000046}');
	});

	it('answers for a project that references nothing unusual', () => {
		const plain = path.join(__dirname, 'fixtures', 'binaries', 'ShapesFixture.xlsm');

		expect(listReferences(plain).map((one) => one.name)).toEqual(['stdole', 'Office']);
	});
});

describe('completing on a referenced library', () => {
	// The other half of what a reference buys: the editor offering the right
	// members. It reads the same hostModel the diagnostics do.
	const source = [
		'Option Explicit',
		'',
		'Public Sub Bridge()',
		'    Dim xl As Excel.Application',
		'    Set xl = New Excel.Application',
		'    xl.',
		'End Sub',
	].join('\r\n');
	const offset = source.indexOf('xl.') + 'xl.'.length;

	const names = (tokens: Parameters<typeof hostObjectModelForTokens>[0]) =>
		resolveMemberCompletions(source, offset, {
			// The completion context calls it `model`; the editor layer maps
			// its `hostModel` onto it.
			model: hostObjectModelForTokens(tokens),
		}).map((one) => one.name);

	it("offers the referenced application's members", () => {
		const offered = names(['word', 'excel']);

		expect(offered).toContain('Workbooks');
		expect(offered).toContain('Calculate');
		expect(offered).toContain('Quit');
	});

	it('offers nothing for a library the project does not reference', () => {
		expect(names(['word'])).toEqual([]);
	});
});

describe('analyzing cross-application code', () => {
	const source = (body: string) => [
		'Option Explicit',
		'',
		'Public Sub Bridge()',
		`    ${body}`,
		'End Sub',
	].join('\r\n');

	const codes = (body: string, tokens: Parameters<typeof hostObjectModelForTokens>[0]) =>
		analyzeModule(source(body), {
			moduleName: 'M', moduleKind: 'standard',
			hostModel: hostObjectModelForTokens(tokens),
		}).map((d) => d.code);

	it('checks a referenced library members, which a lone host cannot', () => {
		const body = 'Dim xl As Excel.Application\r\n    xl.NoSuchMemberAtAll';

		// Word alone has never heard of Excel.Application, so it says nothing.
		expect(codes(body, ['word'])).not.toContain('member-not-found');
		// With the reference declared, the member is checked against Excel.
		expect(codes(body, ['word', 'excel'])).toContain('member-not-found');
	});

	it('accepts what the referenced library really has', () => {
		const body = 'Dim xl As Excel.Application\r\n'
			+ '    Set xl = New Excel.Application\r\n'
			+ '    xl.Calculate';

		expect(codes(body, ['word', 'excel'])).toEqual([]);
	});

	it('knows a referenced type is an object, so the object rules reach it', () => {
		// Used without Set, which is the whole point of knowing the type.
		const body = 'Dim xl As Excel.Application\r\n    xl.Calculate';

		expect(codes(body, ['word', 'excel'])).toContain('object-variable-not-set');
		expect(codes(body, ['word'])).toEqual([]);
	});

	it('resolves a referenced library constant', () => {
		const body = 'Dim n As Long\r\n    n = xlOpenXMLWorkbook\r\n    Debug.Print n';

		expect(codes(body, ['word', 'excel'])).toEqual([]);
	});

	it('leaves a project that references nothing exactly as it was', () => {
		const body = 'Dim xl As Excel.Application\r\n    xl.NoSuchMemberAtAll';

		expect(codes(body, ['word'])).toEqual([]);
	});
});
