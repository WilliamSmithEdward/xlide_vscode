import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import {
	addReference,
	deleteModule,
	listReferences,
	readModules,
	readModulesFromBuffer,
	removeReference,
} from '../src/vba/projectService';
import {
	buildRegisteredReference,
	HOST_LIBRARIES,
	insertReferenceRecords,
	removeReferenceRecords,
} from '../src/vba/vbaProjectReferences';
import { decompress } from '../src/vba/ovba';
import { librariesNamedIn } from '../src/analyzer/diagnostics/rules/missingReference';
import { openMacroContainer } from '../src/vba/macroContainer';
import { analyzeModule } from '../src/analyzer/diagnostics/analyzeModule';
import { resolveMemberCompletions } from '../src/analyzer/completion/memberAccess';
import { resolveTypeCompletions } from '../src/analyzer/completion/typeCompletion';
import { resolveHover } from '../src/analyzer/hover/resolveHover';
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
		// Without the reference the same module does not compile in the VBE
		// either, and that is the one thing said about it.
		expect(wordAlone.map((d) => d.code)).toEqual(['missing-library-reference']);
	});

	it('catches a typo in the referenced library, which Word alone cannot', () => {
		// A Worksheet member: Workbook and Application are extensible, so a
		// name absent from them is deferred to run time rather than refused,
		// and XLIDE reports nothing for it either.
		const typo = `${bridge.source!}\r\nPublic Sub Typo(ws As Excel.Worksheet)\r\n`
			+ '    ws.CalculateXyz\r\nEnd Sub\r\n';
		const tokens = hostTokensForProject('word', bridge.projectReferences ?? []);
		const options = { moduleName: 'Bridge', moduleKind: 'standard' as const };

		expect(analyzeModule(typo, { ...options, hostModel: hostObjectModelForTokens(tokens) })
			.map((d) => d.code)).toContain('member-not-found');
		// Word alone cannot see the typo; it reports the missing reference
		// instead, which is the thing actually wrong with the project.
		expect(analyzeModule(typo, { ...options, hostModel: hostObjectModelForTokens(['word']) })
			.map((d) => d.code)).toEqual(['missing-library-reference']);
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

describe('naming a library the project does not reference', () => {
	const analyze = (body: string, tokens: Parameters<typeof hostObjectModelForTokens>[0]) =>
		analyzeModule(
			['Option Explicit', '', 'Public Sub Bridge()', `    ${body}`, 'End Sub'].join('\r\n'),
			{ moduleName: 'M', moduleKind: 'standard', hostModel: hostObjectModelForTokens(tokens) },
		);
	const codes = (body: string, tokens: Parameters<typeof hostObjectModelForTokens>[0]) =>
		analyze(body, tokens).map((d) => d.code);

	it('reports an early-bound type from an unreferenced library', () => {
		const found = analyze('Dim xl As Excel.Application', ['word']);

		expect(found.map((d) => d.code)).toContain('missing-library-reference');
		expect(found[0].message).toContain("'Excel' is not referenced by this project");
		expect(found[0].message).toContain('late binding');
	});

	it('says nothing once the project references it', () => {
		expect(codes('Dim xl As Excel.Application', ['word', 'excel']))
			.not.toContain('missing-library-reference');
	});

	it('stays silent on late binding, which needs no reference', () => {
		// The whole point of CreateObject: no name from the library appears.
		expect(codes('Dim xl As Object\r\n    Set xl = CreateObject("Excel.Application")', ['word']))
			.not.toContain('missing-library-reference');
		expect(codes('Dim xl As Object\r\n    Set xl = GetObject(, "Excel.Application")', ['word']))
			.not.toContain('missing-library-reference');
	});

	it('reports New and a qualified constant too, which also need the library', () => {
		expect(codes('Dim xl As Object\r\n    Set xl = New Word.Application', ['excel']))
			.toContain('missing-library-reference');
		expect(codes('Dim n As Long\r\n    n = Word.wdFormatPDF', ['excel']))
			.toContain('missing-library-reference');
	});

	it('leaves a member access on a value alone', () => {
		// `wb.Excel.Thing` is not a library qualifier, and neither is a
		// project name that happens to match further along a chain.
		expect(codes('Dim wb As Object\r\n    wb.Excel.Thing = 1', ['word']))
			.not.toContain('missing-library-reference');
	});

	it('says nothing about a library XLIDE cannot add', () => {
		expect(codes('Dim o As Outlook.Application', ['word']))
			.not.toContain('missing-library-reference');
		expect(codes('Dim s As Scripting.Dictionary', ['word']))
			.not.toContain('missing-library-reference');
	});

	it('carries the library for the quick fix that adds it', () => {
		const found = analyze('Dim xl As Excel.Application', ['word']);
		const missing = found.find((d) => d.code === 'missing-library-reference')!;

		expect(missing.data?.addLibraryReference).toEqual({ library: 'excel' });
	});
});

describe('adding a reference the project is missing', () => {
	// The other half of the diagnostic: the code names Excel, the project does
	// not reference it, and the fix writes the reference rather than sending
	// the user to the VBE's Tools > References dialog.
	let dir: string;
	const copy = (fixture: string): string => {
		const target = path.join(dir, fixture);
		fs.copyFileSync(path.join(__dirname, 'fixtures', 'binaries', fixture), target);
		return target;
	};
	/** The compiled cache Office keeps beside the source. */
	const compiledCache = (file: string): Buffer => openMacroContainer(fs.readFileSync(file))
		.vbaCfb().getStreamInStorage('VBA', '_VBA_PROJECT');

	beforeEach(() => {
		dir = fs.mkdtempSync(path.join(os.tmpdir(), 'xlide-add-reference-'));
	});
	afterEach(() => {
		fs.rmSync(dir, { recursive: true, force: true });
	});

	it('writes it so the project reports it, as Tools > References would', () => {
		const file = copy('ShapesFixture.xlsm');

		expect(addReference(file, 'word')).toEqual({ ok: true, added: true, name: 'Word' });
		const written = listReferences(file).find((one) => one.name === 'Word')!;
		expect(written.kind).toBe('registered');
		expect(written.libid).toContain('{00020905-0000-0000-C000-000000000046}');
		expect(hostTokenForLibid(written.libid)).toBe('word');
	});

	it('leaves the references the project already had', () => {
		const file = copy('ShapesFixture.xlsm');
		addReference(file, 'word');

		expect(listReferences(file).map((one) => one.name)).toEqual(['stdole', 'Office', 'Word']);
	});

	it('is what clears the diagnostic, end to end', () => {
		const file = copy('ShapesFixture.xlsm');
		const analyze = () => {
			const module = readModules(file, true).find((one) => one.source)!;
			return analyzeModule(
				['Option Explicit', '', 'Public Sub Bridge()', '    Dim doc As Word.Document',
					'End Sub'].join('\r\n'),
				{
					moduleName: 'M',
					moduleKind: 'standard',
					hostModel: hostObjectModelForTokens(
						hostTokensForProject('excel', module.projectReferences ?? []),
					),
				},
			).map((d) => d.code);
		};

		expect(analyze()).toContain('missing-library-reference');
		addReference(file, 'word');
		expect(analyze()).not.toContain('missing-library-reference');
	});

	it('adds nothing the second time', () => {
		const file = copy('ShapesFixture.xlsm');
		addReference(file, 'word');

		expect(addReference(file, 'word')).toEqual({ ok: true, added: false, name: 'Word' });
		expect(listReferences(file).filter((one) => one.name === 'Word')).toHaveLength(1);
	});

	it('refuses a library it has no reference to write', () => {
		const file = copy('ShapesFixture.xlsm');

		expect(() => addReference(file, 'outlook')).toThrow(/not a library XLIDE can add/);
		expect(listReferences(file).map((one) => one.name)).toEqual(['stdole', 'Office']);
	});

	it('marks the compiled project stale, or the host goes on ignoring it', () => {
		// Measured against Word 16: a reference written into the dir stream
		// beside an untouched _VBA_PROJECT is invisible to the host, because
		// it runs the compiled project rather than reading the records. A
		// reference is a mutating change, so the cache body goes.
		const file = copy('ShapesFixture.xlsm');
		expect(compiledCache(file).subarray(5).some((byte) => byte !== 0)).toBe(true);

		addReference(file, 'word');

		const after = compiledCache(file);
		expect(after.subarray(0, 5)).toEqual(Buffer.from([0xcc, 0x61, after[2], after[3], 0x00]));
		expect(after.subarray(5).every((byte) => byte === 0)).toBe(true);
	});

	it('knows every application it diagnoses, so each fix can be carried out', () => {
		for (const token of ['word', 'powerpoint', 'access'] as const) {
			const file = copy('ShapesFixture.xlsm');
			addReference(file, token);
			const written = listReferences(file).at(-1)!;

			expect(hostTokenForLibid(written.libid)).toBe(token);
			expect(hostTokensForProject('excel', listReferences(file))).toContain(token);
			fs.rmSync(file);
		}
		// The fourth, from the other side: a document gaining Excel.
		const doc = path.join(dir, 'WordExcelInteropFixture.docm');
		fs.copyFileSync(path.join(__dirname, 'fixtures', 'binaries', 'WordExcelInteropFixture.docm'), doc);
		expect(hostTokensForProject('word', listReferences(doc))).toContain('excel');
	});

	it('refuses the file own host, which the project carries implicitly', () => {
		// Neither Excel nor Word declares its own library in the project's
		// records - Tools > References shows it checked and greyed - so a
		// written one would be a second, redundant copy of it.
		const file = copy('ShapesFixture.xlsm');

		expect(() => addReference(file, 'excel')).toThrow(/already has the Excel object library/);
		expect(listReferences(file).map((one) => one.name)).toEqual(['stdole', 'Office']);
	});
});

describe('analyzing with the references rather than a resolved model', () => {
	// How the editor reaches the analyzer: live diagnostics know the project's
	// host and which libraries it references, and hand over both as tokens
	// rather than building a model. The worker thread has the same pair, since
	// a model does not cross a thread boundary.
	// A parameter rather than a local, so nothing but the type resolution
	// is under test: an unassigned local of a class type is a finding of
	// its own, whichever library the type comes from.
	const source = ['Option Explicit', '', 'Public Sub Bridge(wb As Excel.Workbook)',
		'    Debug.Print wb.Name', 'End Sub'].join('\r\n');
	const codes = (options: Parameters<typeof analyzeModule>[1]) =>
		analyzeModule(source, { moduleName: 'M', moduleKind: 'standard', ...options })
			.map((d) => d.code);

	it('resolves the referenced library beside the host', () => {
		expect(codes({ host: 'word' })).toEqual(['missing-library-reference']);
		expect(codes({ host: 'word', referencedHosts: ['excel'] })).toEqual([]);
	});

	it('leaves the host as the one the module belongs to', () => {
		// `Me` in a document module is the host's document, whatever the
		// project also references: the referenced library adds names, it does
		// not change which application the module runs in.
		const document = ['Option Explicit', '', 'Public Sub P()', '    Debug.Print Me.Name', 'End Sub']
			.join('\r\n');
		const options = { moduleName: 'ThisDocument', moduleKind: 'document' as const, host: 'word' };

		expect(analyzeModule(document, options).map((d) => d.code)).toEqual([]);
		expect(analyzeModule(document, { ...options, referencedHosts: ['excel'] })
			.map((d) => d.code)).toEqual([]);
	});

	it('keeps the Excel defaults for a project that references nothing', () => {
		// No host and no references is how every caller that knows neither
		// arrives, and it has always meant Excel.
		expect(codes({})).toEqual([]);
		expect(codes({ referencedHosts: [] })).toEqual([]);
	});
});

describe('taking a reference away again', () => {
	let dir: string;
	const copy = (fixture: string): string => {
		const target = path.join(dir, fixture);
		fs.copyFileSync(path.join(__dirname, 'fixtures', 'binaries', fixture), target);
		return target;
	};
	/** The reference records themselves, as the host reads them. */
	const dirStream = (file: string): Buffer => decompress(
		openMacroContainer(fs.readFileSync(file)).vbaCfb().getStreamInStorage('VBA', 'dir'),
		'VBA/dir',
	);

	beforeEach(() => {
		dir = fs.mkdtempSync(path.join(os.tmpdir(), 'xlide-remove-reference-'));
	});
	afterEach(() => {
		fs.rmSync(dir, { recursive: true, force: true });
	});

	it('cuts exactly the records an addition wrote, and nothing either side', () => {
		// A reference is a run of records, not one: the name in two encodings
		// and then the record carrying the libid. On the buffer itself, with
		// no writer in the way, putting a block in and taking it out again has
		// to give the original bytes back.
		const dir = dirStream(copy('ShapesFixture.xlsm'));
		const block = buildRegisteredReference(HOST_LIBRARIES.word);

		const grown = insertReferenceRecords(dir, block);
		expect(grown.length).toBe(dir.length + block.length);

		expect(removeReferenceRecords(grown, (one) => one.name === 'Word')).toEqual(dir);
	});

	it('leaves the project the way an untouched save would, through the engine', () => {
		// Not compared against the file Excel wrote: XLIDE's writer normalizes
		// the dir stream's terminator on its first save, so both sides of the
		// comparison are streams it has written.
		const file = copy('ShapesFixture.xlsm');
		addReference(file, 'word');
		const withWord = dirStream(file);

		expect(removeReference(file, 'Word')).toEqual({ ok: true, removed: true, name: 'Word' });
		const withoutWord = dirStream(file);
		expect(withoutWord.length).toBeLessThan(withWord.length);

		// Adding it back lands on the same bytes, which it could not do if the
		// removal had taken a byte too many or left one behind.
		addReference(file, 'word');
		expect(dirStream(file)).toEqual(withWord);
	});

	it('takes the one asked for and leaves the others as they were', () => {
		const file = copy('WordExcelInteropFixture.docm');

		expect(removeReference(file, 'Excel').removed).toBe(true);

		expect(listReferences(file).map((one) => one.name)).toEqual(['stdole', 'Normal', 'Office']);
	});

	it('answers to the host token as well as to the name', () => {
		const file = copy('WordExcelInteropFixture.docm');

		expect(removeReference(file, 'excel')).toEqual({ ok: true, removed: true, name: 'Excel' });
		expect(listReferences(file).map((one) => one.name)).not.toContain('Excel');
	});

	it('cuts a control reference whole, which spans seven records', () => {
		// stdole and Office are plain registered references; MSForms is a
		// control reference carrying the original libid, the control record,
		// its name a second time and the extended record.
		const file = copy('FormFixtureVbide.xlsm');
		const before = listReferences(file).map((one) => one.name);
		expect(before).toContain('MSForms');

		// The forms are what hold the reference in place, so they go first.
		for (const form of readModules(file).filter((one) => one.type === 'userform')) {
			deleteModule(file, form.name);
		}
		expect(removeReference(file, 'MSForms').removed).toBe(true);

		expect(listReferences(file).map((one) => one.name))
			.toEqual(before.filter((name) => name !== 'MSForms'));
		// The project still reads, so nothing either side of the span went.
		expect(readModules(file).length).toBeGreaterThan(0);
	});

	it('refuses to leave a UserForm without the library it needs', () => {
		const file = copy('FormFixtureVbide.xlsm');

		expect(() => removeReference(file, 'MSForms')).toThrow(/needs?\b.*Microsoft Forms/);
		expect(listReferences(file).map((one) => one.name)).toContain('MSForms');
	});

	it('changes nothing for a reference the project does not have', () => {
		const file = copy('ShapesFixture.xlsm');
		const before = dirStream(file);

		expect(removeReference(file, 'powerpoint')).toEqual({ ok: true, removed: false, name: 'PowerPoint' });
		expect(dirStream(file)).toEqual(before);
	});

	it('marks the compiled project stale, as adding one does', () => {
		const file = copy('WordExcelInteropFixture.docm');
		const cache = (): Buffer => openMacroContainer(fs.readFileSync(file))
			.vbaCfb().getStreamInStorage('VBA', '_VBA_PROJECT');
		expect(cache().subarray(5).some((byte) => byte !== 0)).toBe(true);

		removeReference(file, 'Excel');

		expect(cache().subarray(5).every((byte) => byte === 0)).toBe(true);
	});

	it('leaves the modules exactly as they were', () => {
		const file = copy('WordExcelInteropFixture.docm');
		const before = readModules(file, true).map((one) => `${one.name}:${one.source}`);

		removeReference(file, 'Excel');

		expect(readModules(file, true).map((one) => `${one.name}:${one.source}`)).toEqual(before);
	});
});

describe('which libraries a module names', () => {
	// What the removal warns from: the same scan the missing-reference rule
	// runs, so the two can never disagree about what counts as naming one.
	it('finds an early-bound name wherever the compiler has to resolve it', () => {
		expect([...librariesNamedIn('Dim xl As Excel.Application')]).toEqual(['excel']);
		expect([...librariesNamedIn('Set o = New Word.Document')]).toEqual(['word']);
		expect([...librariesNamedIn('n = Word.wdFormatPDF')]).toEqual(['word']);
	});

	it('says nothing about late binding, which needs no reference', () => {
		expect([...librariesNamedIn('Set xl = CreateObject("Excel.Application")')]).toEqual([]);
	});

	it('reports each library once, however often the module names it', () => {
		const source = ['Dim a As Excel.Range', 'Dim b As Excel.Worksheet', 'Dim c As Word.Range'].join('\r\n');

		expect([...librariesNamedIn(source)].sort()).toEqual(['excel', 'word']);
	});
});

describe('which application a label names', () => {
	// Issue #77: a merged model has one hostName - the project's own, by
	// construction - and every label was built from it, so a Word member in
	// an Excel workbook read as Excel's. Resolution was right all along; the
	// one line a developer reads to find out where a member comes from was
	// not.
	const model = hostObjectModelForTokens(['excel', 'word']);
	const hoverDetails = (source: string, needle: string): string[] | undefined =>
		resolveHover(source, source.indexOf(needle) + 2, { model })?.details;
	const wrap = (...lines: string[]) =>
		['Option Explicit', '', 'Public Sub P()', ...lines.map((one) => `    ${one}`), 'End Sub']
			.join('\r\n');

	it("names the referenced application on its own member, not the project's host", () => {
		const source = wrap('Dim wd As Word.Application', 'wd.Visible = True');

		expect(hoverDetails(source, 'Visible = True')).toEqual(['Word host property (read/write)']);
	});

	it("still names the project's own host on its own members", () => {
		const source = wrap('Dim rng As Range', 'rng.Value = 1');

		expect(hoverDetails(source, 'Value = 1')).toEqual(['Excel host property (read/write)']);
	});

	it('labels each library\'s types and enums with that library', () => {
		const offered = resolveTypeCompletions('Dim x As \r\n', 9, { model });
		const detailOf = (name: string) => offered.find((one) => one.name === name)?.detail;

		expect(detailOf('Document')).toBe('Word type');
		expect(detailOf('Worksheet')).toBe('Excel type');
		expect(detailOf('WdSaveFormat')).toBe('Word enum');
		expect(detailOf('XlAxisType')).toBe('Excel enum');
	});

	it('leaves a single host model labelling everything as that host', () => {
		const excelOnly = resolveTypeCompletions('Dim x As \r\n', 9, {});

		expect(excelOnly.find((one) => one.name === 'Worksheet')?.detail).toBe('Excel type');
		expect(excelOnly.find((one) => one.name === 'XlAxisType')?.detail).toBe('Excel enum');
	});
});

describe('a class name two applications share', () => {
	// Range is Word's and Excel's both. VBA resolves the bare name by the
	// reference list with the host at the top, so it is the project's own
	// host's, and the other is reached by qualifying it. Everything a
	// developer reads has to agree with that: offering Word's Range to an
	// Excel project and then checking what they wrote against Excel's is
	// worse than either answer on its own.
	const wordProject = hostObjectModelForTokens(['word', 'excel']);
	const excelProject = hostObjectModelForTokens(['excel', 'word']);
	const wrap = (...lines: string[]) =>
		['Option Explicit', '', 'Public Sub P()', ...lines.map((one) => `    ${one}`), 'End Sub']
			.join('\r\n');
	const memberDetail = (model: typeof wordProject, declared: string, member: string) => {
		const source = wrap(`Dim r As ${declared}`, `r.${member} = 1`);
		return resolveHover(source, source.indexOf(`${member} = 1`) + 2, { model })?.details?.[0];
	};
	const typeDetail = (model: typeof wordProject, declared: string) => {
		const source = wrap(`Dim r As ${declared}`, 'r.Select');
		const bare = declared.split('.').pop()!;
		return resolveHover(source, source.indexOf(`As ${declared}`) + 3 + (declared.length - bare.length) + 1,
			{ model })?.details?.[0];
	};

	it("resolves the bare name to the project's own host, either way round", () => {
		// Bold is Word's, Value2 is Excel's.
		expect(memberDetail(wordProject, 'Range', 'Bold')).toBe('Word host property (read/write)');
		expect(memberDetail(wordProject, 'Range', 'Value2')).toBeUndefined();
		expect(memberDetail(excelProject, 'Range', 'Value2')).toBe('Excel host property (read/write)');
		expect(memberDetail(excelProject, 'Range', 'Bold')).toBeUndefined();
	});

	it('resolves a qualified name to the library that qualifies it', () => {
		expect(memberDetail(excelProject, 'Word.Range', 'Bold')).toBe('Word host property (read/write)');
		expect(memberDetail(wordProject, 'Excel.Range', 'Value2')).toBe('Excel host property (read/write)');
		// And does not quietly fall through to the other one.
		expect(memberDetail(excelProject, 'Word.Range', 'Value2')).toBeUndefined();
		expect(memberDetail(wordProject, 'Excel.Range', 'Bold')).toBeUndefined();
	});

	it('names the qualifier on a qualified type, not the host', () => {
		// The report: `Dim rng As Word.Range` in a workbook hovered as an
		// Excel type, over Word's own description of it.
		expect(typeDetail(excelProject, 'Word.Range')).toBe('Word host type');
		expect(typeDetail(excelProject, 'Excel.Range')).toBe('Excel host type');
		expect(typeDetail(wordProject, 'Excel.Range')).toBe('Excel host type');
	});

	it("offers the bare name as the host's type, which is what it resolves to", () => {
		const offered = (model: typeof wordProject) =>
			resolveTypeCompletions('Dim x As \r\n', 9, { model })
				.filter((one) => one.name === 'Range')
				.map((one) => one.detail);

		expect(offered(wordProject)).toEqual(['Word type']);
		expect(offered(excelProject)).toEqual(['Excel type']);
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
		const body = 'Dim ws As Excel.Worksheet\r\n    ws.NoSuchMemberAtAll';

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
		// Word alone knows no Excel.Application to apply the object rules to,
		// and says only that the library is not referenced.
		expect(codes(body, ['word'])).toEqual(['missing-library-reference']);
	});

	it('resolves a referenced library constant', () => {
		const body = 'Dim n As Long\r\n    n = xlOpenXMLWorkbook\r\n    Debug.Print n';

		expect(codes(body, ['word', 'excel'])).toEqual([]);
	});

	it('reports the missing reference rather than the member it cannot check', () => {
		// The member is unknowable without the library, so the one useful
		// thing to say is that the library is not referenced.
		const body = 'Dim xl As Excel.Application\r\n    xl.NoSuchMemberAtAll';

		expect(codes(body, ['word'])).toEqual(['missing-library-reference']);
	});
});
