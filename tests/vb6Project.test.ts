import { afterEach, describe, expect, it } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { manifestValue, parseVbpManifest, printVbpManifest } from '../src/vba/vb6/vbpProject';
import { resetVb6ProjectCacheForTests } from '../src/vba/vb6/vb6Project';
import {
	getProtectionInfo,
	getWorkbookInfo,
	listModules,
	listSubs,
	readModule,
	readModules,
	validateWorkbook,
	writeModule,
} from '../src/vba/workbookService';
import { hostTokenForFileName } from '../src/analyzer/host/hostRegistry';
import { containerContextValue, MACRO_CONTAINER_GLOB } from '../src/macroContainerUi';

// A VB6 project is a `.vbp` manifest over loose text files. The engine
// answers the same module questions for it that it answers for a workbook,
// so the tree, the analyzer and the agent tools treat it as one more
// container. The fixtures are real projects authored in Visual Basic 6
// (tests/fixtures/vb6/*/NOTICE.md names their source and license).

const FIXTURES = path.join(__dirname, 'fixtures', 'vb6');
const RUN_AS_TI = path.join(FIXTURES, 'RunAsTrustedInstaller', 'Project1.vbp');
const DIABETES = path.join(FIXTURES, 'Diabetes-prediction-1.0', 'MCD_prj.vbp');

const tempDirs: string[] = [];
afterEach(() => {
	resetVb6ProjectCacheForTests();
	for (const dir of tempDirs.splice(0)) {
		fs.rmSync(dir, { recursive: true, force: true });
	}
});

function tempProject(files: Record<string, string>): string {
	const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'xlide-vb6-'));
	tempDirs.push(dir);
	for (const [name, text] of Object.entries(files)) {
		fs.mkdirSync(path.dirname(path.join(dir, name)), { recursive: true });
		fs.writeFileSync(path.join(dir, name), text, 'latin1');
	}
	return dir;
}

const CRLF = '\r\n';

describe('the .vbp manifest', () => {
	it('reads the fixture manifest into modules, references and settings', () => {
		const text = fs.readFileSync(RUN_AS_TI, 'latin1');
		const manifest = parseVbpManifest(text);

		expect(manifest.type).toBe('Exe');
		expect(manifest.name).toBe('RunAsTrustedInstaller');
		expect(manifest.startup).toBe('Form1');
		expect(manifest.exeName32).toBe('RunAsTI.exe');
		expect(manifest.modules).toEqual([
			{ kind: 'form', file: 'Form1.frm', line: 1 },
			{ kind: 'module', name: 'modRunAsTI', file: 'modRunAsTI.bas', line: 3 },
		]);
		expect(manifest.references).toHaveLength(1);
		expect(manifest.references[0]).toMatchObject({
			guid: '{00020430-0000-0000-C000-000000000046}',
			version: '2.0',
			lcid: '0',
			description: 'OLE Automation',
		});
		expect(manifest.references[0].path).toMatch(/stdole2\.tlb$/);
		// A setting the parser never interprets is still readable by key.
		expect(manifestValue(manifest, 'FavorPentiumPro(tm)')).toBe('0');
		expect(manifestValue(manifest, 'MajorVer')).toBe('2');
	});

	it('prints back byte for byte, unknown keys and all', () => {
		for (const file of [RUN_AS_TI, DIABETES]) {
			const text = fs.readFileSync(file, 'latin1');
			expect(printVbpManifest(parseVbpManifest(text))).toBe(text);
		}
	});

	it('understands every member kind, components, and sectioned settings', () => {
		const text = [
			'Type=OleDll',
			'Class=cSocket; ..\\src\\cSocket.cls',
			'Module=modMain; modMain.bas',
			'Form=frmMain.frm',
			'UserControl=ctxWinsock.ctl',
			'PropertyPage=ppGeneral.pag',
			'Designer=dsrReport.dsr',
			'RelatedDoc=ReadMe.txt',
			'Object={831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0; mscomctl.ocx',
			'Reference=*\\G{420B2830-E718-11CF-893D-00A0C9054228}#1.0#0#C:\\Windows\\SysWOW64\\scrrun.dll#Microsoft Scripting Runtime',
			'Startup="Sub Main"',
			'CondComp="DEBUG_MODE = -1"',
			'Frobnicate=yes',
			'',
			'[MS Transaction Server]',
			'AutoRefresh=1',
		].join(CRLF) + CRLF;
		const manifest = parseVbpManifest(text);

		expect(manifest.modules.map((m) => [m.kind, m.name, m.file])).toEqual([
			['class', 'cSocket', '..\\src\\cSocket.cls'],
			['module', 'modMain', 'modMain.bas'],
			['form', undefined, 'frmMain.frm'],
			['usercontrol', undefined, 'ctxWinsock.ctl'],
			['propertypage', undefined, 'ppGeneral.pag'],
			['designer', undefined, 'dsrReport.dsr'],
			['relateddoc', undefined, 'ReadMe.txt'],
		]);
		expect(manifest.objects).toEqual([expect.objectContaining({
			guid: '{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}', version: '2.0', file: 'mscomctl.ocx',
		})]);
		expect(manifest.references[0]).toMatchObject({
			guid: '{420B2830-E718-11CF-893D-00A0C9054228}',
			path: 'C:\\Windows\\SysWOW64\\scrrun.dll',
			description: 'Microsoft Scripting Runtime',
		});
		expect(manifest.startup).toBe('Sub Main');
		expect(manifest.condComp).toBe('DEBUG_MODE = -1');
		// A sectioned key is a setting, never a member, and it keeps its section.
		expect(manifest.lines.find((l) => l.key === 'AutoRefresh')?.section).toBe('MS Transaction Server');
		expect(manifestValue(manifest, 'AutoRefresh')).toBeUndefined();
		expect(printVbpManifest(manifest)).toBe(text);
	});

	it('keeps LF endings and a missing final newline as it found them', () => {
		const text = 'Type=Exe\nForm=Form1.frm\nName="P"';
		expect(printVbpManifest(parseVbpManifest(text))).toBe(text);
	});

	it('refuses a file with no Key=Value line at all', () => {
		expect(() => parseVbpManifest('just some text\r\nand more\r\n')).toThrow(/Not a Visual Basic project file/);
	});
});

describe('a VB6 project through the engine', () => {
	it('lists the modules the manifest names, typed and with their files', () => {
		const modules = listModules(RUN_AS_TI);
		expect(modules).toEqual([
			{ name: 'Form1', type: 'userform', filePath: path.join(path.dirname(RUN_AS_TI), 'Form1.frm') },
			{ name: 'modRunAsTI', type: 'standard', filePath: path.join(path.dirname(RUN_AS_TI), 'modRunAsTI.bas') },
		]);
	});

	it('names a form by its VB_Name attribute, not its file name', () => {
		// The manifest says `Form=MCD.frm`; the file declares `VB_Name = "Form1"`.
		expect(listModules(DIABETES).map((m) => `${m.name}:${m.type}`)).toEqual(['Form1:userform', 'OS:standard']);
	});

	it('reads a form as its own file with the designer block blanked, every offset kept', () => {
		const raw = fs.readFileSync(path.join(path.dirname(RUN_AS_TI), 'Form1.frm'), 'latin1');
		const aligned = readModule(RUN_AS_TI, 'Form1').source;
		expect(aligned.length).toBe(raw.length);
		expect(aligned).not.toMatch(/Begin VB\.Form/);
		expect(aligned).toMatch(/^Attribute VB_Name = "Form1"/m);
		expect(aligned).toMatch(/Private Sub Form_Load\(\)/);
		// The module's lines ARE the file's lines: what an analysis or an
		// agent reports as line 128 is line 128 of Form1.frm.
		expect(aligned.split('\r\n')[127]).toBe(raw.split('\r\n')[127]);

		expect(readModule(RUN_AS_TI, 'Form1', true).source).toBe(raw);
	});

	it('reads a standard module and lists its procedures with lines', () => {
		expect(readModule(DIABETES, 'OS').source).toMatch(/GetVersionEx/);
		const subs = listSubs(DIABETES, 'OS');
		expect(subs).toEqual(expect.arrayContaining([
			expect.objectContaining({ name: 'ItIsWin7', kind: 'Function', line: 26 }),
		]));
		// Lines are the file's own, designer block and attributes counted.
		const formSubs = listSubs(RUN_AS_TI, 'Form1');
		expect(formSubs.map((s) => s.name)).toEqual(
			expect.arrayContaining(['Command1_Click', 'Form_Load', 'AppendLog', 'Form_Unload']),
		);
		expect(formSubs.find((s) => s.name === 'Form_Load')?.line).toBe(128);
	});

	it('reads every module with sources and predeclared-instance facts', () => {
		const modules = readModules(RUN_AS_TI, true);
		const form = modules.find((m) => m.name === 'Form1');
		expect(form?.predeclaredId).toBe(true);
		expect(form?.source).toMatch(/Attribute VB_PredeclaredId = True/);
		// A real VB6 form header carries BeginProperty blocks the header
		// parser cannot yet read, so the controls stay "not known" (absent),
		// never an empty list (roadmap_vb6_support.md, Slice 2).
		expect(form?.implicitMembers).toBeUndefined();
		expect(modules.find((m) => m.name === 'modRunAsTI')?.source).toMatch(/Option Explicit/);
		expect(modules.map((m) => m.filePath)).toEqual([
			path.join(path.dirname(RUN_AS_TI), 'Form1.frm'),
			path.join(path.dirname(RUN_AS_TI), 'modRunAsTI.bas'),
		]);
	});

	it('answers the workbook-shaped questions honestly for a project', () => {
		expect(getProtectionInfo(RUN_AS_TI)).toEqual({ isPasswordProtected: false, isSigned: false });
		const info = getWorkbookInfo(RUN_AS_TI);
		expect(info.sheets).toEqual([]);
		expect(info.modules.map((m) => m.name)).toEqual(['Form1', 'modRunAsTI']);
		expect(validateWorkbook(RUN_AS_TI)).toEqual({ issues: [] });
	});

	it('lists a module whose file is missing, and names the file when it is read', () => {
		const dir = tempProject({
			'P.vbp': ['Type=Exe', 'Module=modGone; modGone.bas', 'Module=modHere; modHere.bas', 'Startup="Sub Main"'].join(CRLF) + CRLF,
			'modHere.bas': ['Attribute VB_Name = "modHere"', 'Sub Main()', 'End Sub'].join(CRLF) + CRLF,
		});
		const vbp = path.join(dir, 'P.vbp');

		expect(listModules(vbp).map((m) => m.name)).toEqual(['modGone', 'modHere']);
		expect(() => readModule(vbp, 'modGone')).toThrow(/Module file not found: .*modGone\.bas/);
		expect(readModules(vbp).map((m) => m.name)).toEqual(['modHere']);
		expect(validateWorkbook(vbp).issues).toEqual([expect.stringMatching(/Missing module file: modGone\.bas/)]);
	});

	it('lists the opaque kinds by name and answers no procedures for them', () => {
		const dir = tempProject({
			'P.vbp': ['Type=Control', 'UserControl=ctxThing.ctl', 'PropertyPage=ppThing.pag', 'Startup="(None)"'].join(CRLF) + CRLF,
			'ctxThing.ctl': ['VERSION 5.00', 'Begin VB.UserControl ctxThing ', '   ClientHeight    =   3000', 'End',
				'Attribute VB_Name = "ctxThing"', 'Option Explicit', 'Private Sub UserControl_Resize()', 'End Sub'].join(CRLF) + CRLF,
			'ppThing.pag': ['VERSION 5.00', 'Begin VB.PropertyPage ppThing ', 'End', 'Attribute VB_Name = "ppThing"'].join(CRLF) + CRLF,
		});
		const vbp = path.join(dir, 'P.vbp');

		expect(listModules(vbp).map((m) => `${m.name}:${m.type}`)).toEqual(['ctxThing:usercontrol', 'ppThing:propertypage']);
		// The code-behind is readable like any module's; the designer is what stays opaque.
		expect(listSubs(vbp, 'ctxThing')).toEqual([{ name: 'UserControl_Resize', kind: 'Sub', line: 7 }]);
		const page = readModule(vbp, 'ppThing').source;
		expect(page).not.toMatch(/Begin VB\.PropertyPage/);
		expect(page).toMatch(/Attribute VB_Name = "ppThing"/);
	});

	it('refuses a manifest that is not one', () => {
		const dir = tempProject({ 'Bad.vbp': 'this is not a project\r\n' });
		expect(() => listModules(path.join(dir, 'Bad.vbp'))).toThrow(/Not a Visual Basic project file/);
	});

	it('rewrites an existing module body, keeping header, line endings and encoding', () => {
		const dir = tempProject({
			'P.vbp': ['Type=Exe', 'Module=modA; modA.bas', 'Startup="Sub Main"'].join(CRLF) + CRLF,
			'modA.bas': ['Attribute VB_Name = "modA"', 'Option Explicit', 'Sub Main()', '    Debug.Print "caf\u00e9"', 'End Sub'].join(CRLF) + CRLF,
		});
		const vbp = path.join(dir, 'P.vbp');
		const file = path.join(dir, 'modA.bas');
		const before = fs.readFileSync(file);
		// Windows-1252 on disk: one byte for e-acute.
		expect(before.includes(Buffer.from([0xe9]))).toBe(true);

		// What an agent hands back is what it read: blank lines where the
		// header was, the attributes, then the code - only the code is new.
		writeModule(vbp, 'modA', '\n\nAttribute VB_Name = "modA"\nOption Explicit\nSub Main()\n    Debug.Print "na\u00efve"\nEnd Sub\n');

		const after = fs.readFileSync(file);
		const text = after.toString('latin1');
		expect(text.startsWith('Attribute VB_Name = "modA"\r\nOption Explicit\r\n')).toBe(true);
		expect(text.match(/Attribute VB_Name/g)).toHaveLength(1);
		expect(text).toContain('Debug.Print "na\u00efve"');
		expect(text).not.toContain('\n' + 'Sub Main()\n'); // CRLF kept
		expect(after.includes(Buffer.from([0xef]))).toBe(true); // i-diaeresis as one byte
		expect(readModule(vbp, 'modA').source).toMatch(/na\u00efve/);
	});

	it('refuses to add, rename or delete a module, naming what to edit instead', () => {
		expect(() => writeModule(RUN_AS_TI, 'modNew', 'Sub X()\nEnd Sub\n')).toThrow(/Adding a module to a VB6 project/);
	});
});

describe('the UI-side facts', () => {
	it('knows a .vbp is a VB6 project', () => {
		expect(hostTokenForFileName('C:\\proj\\App.vbp')).toBe('vb6');
		expect(containerContextValue('C:\\proj\\App.vbp')).toBe('vb6Project');
		expect(MACRO_CONTAINER_GLOB).toContain('vbp');
	});
});
