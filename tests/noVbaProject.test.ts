import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { afterAll, describe, expect, it } from 'vitest';
import {
	addReference,
	addFormModule,
	getModulesAndProtectionInfo,
	getProjectInfo,
	getProtectionInfo,
	hasVbaProject,
	listModules,
	listReferences,
	listSheets,
	listSubs,
	readCells,
	readModule,
	readModules,
	validateProject,
	writeCells,
	writeModule,
} from '../src/vba/projectService';
import {
	MacroContainerError,
	NoVbaProjectError,
	openMacroContainer,
} from '../src/vba/macroContainer';

// A macro-enabled file with no macros in it is an ordinary state, not a
// failure to read one: Excel writes no xl/vbaProject.bin part at all into a
// workbook saved as .xlsm before the first macro exists, and Word, legacy
// Excel and legacy PowerPoint are the same. XLIDE showed all four as
// "Load failed - click to retry", which sent people hunting for a problem
// that was not there.
//
// Every fixture below was saved by the live application itself and never had
// any VBA in it, so the bytes are Office ground truth for the empty state.

const FIXTURES = path.join(__dirname, 'fixtures', 'binaries');
const fixture = (name: string): string => path.join(FIXTURES, name);

const tempRoot = fs.mkdtempSync(path.join(os.tmpdir(), 'xlide-novba-'));
afterAll(() => {
	fs.rmSync(tempRoot, { recursive: true, force: true });
});

function copyOf(name: string): string {
	const target = path.join(tempRoot, `${Math.random().toString(36).slice(2)}-${name}`);
	fs.copyFileSync(fixture(name), target);
	return target;
}

/** Every container shape whose "no VBA yet" path differs, and its host. */
const EMPTY_CONTAINERS: Array<{ file: string; description: string; app: string }> = [
	{ file: 'NoVbaFixture.xlsm', description: 'an Excel workbook', app: 'Excel' },
	{ file: 'NoVbaFixture.docm', description: 'a Word macro-enabled document', app: 'Word' },
	{ file: 'NoVbaFixture.xls', description: 'a legacy Excel workbook (.xls)', app: 'Excel' },
	{ file: 'NoVbaFixture.ppt', description: 'a legacy PowerPoint presentation (.ppt)', app: 'PowerPoint' },
];

describe('a file with no VBA in it yet', () => {
	it.each(EMPTY_CONTAINERS)('$file: the container seam raises the typed state', ({ file, description }) => {
		const container = openMacroContainer(fs.readFileSync(fixture(file)));
		// The file itself read fine; it is only the project that is absent.
		expect(container.description).toBe(description);
		expect(() => container.vbaCfb()).toThrow(NoVbaProjectError);
		try {
			container.vbaCfb();
		} catch (err) {
			expect((err as NoVbaProjectError).containerDescription).toBe(description);
		}
	});

	it.each(EMPTY_CONTAINERS)('$file: listing what it holds answers "nothing", not an error', ({ file }) => {
		expect(listModules(fixture(file))).toEqual([]);
		expect(readModules(fixture(file))).toEqual([]);
		expect(listReferences(fixture(file))).toEqual([]);
		expect(getProtectionInfo(fixture(file))).toEqual({ isPasswordProtected: false, isSigned: false });
		expect(getModulesAndProtectionInfo(fixture(file)))
			.toEqual({ modules: [], isPasswordProtected: false, isSigned: false });
		expect(getProjectInfo(fixture(file))).toMatchObject({ modules: [], isPasswordProtected: false, isSigned: false });
	});

	it.each(EMPTY_CONTAINERS)('$file: nothing is structurally wrong with it', ({ file }) => {
		// Not "VBA project could not be parsed": there is no project to parse,
		// and reporting an issue would read as damage.
		expect(validateProject(fixture(file)).issues).toEqual([]);
	});

	it.each(EMPTY_CONTAINERS)('$file: holds no VBA project, which is not the same as an empty one', ({ file }) => {
		expect(hasVbaProject(fixture(file))).toBe(false);
	});

	it('a blank Access database holds an EMPTY project, not none', () => {
		// Measured on the file Access itself writes: dir, PROJECT, PROJECTwm
		// and _VBA_PROJECT are all there with no modules declared. Both cases
		// list nothing; only this one can take a new module, which is why the
		// two are told apart rather than folded together.
		expect(hasVbaProject('assets/templates/blank.accdb')).toBe(true);
		expect(listModules('assets/templates/blank.accdb')).toEqual([]);
		expect(hasVbaProject('assets/templates/blank.mdb')).toBe(true);
		expect(listModules('assets/templates/blank.mdb')).toEqual([]);
	});
});

describe('asking a file with no VBA for something named', () => {
	it.each(EMPTY_CONTAINERS)('$file: says what the file is and what to do', ({ file, description, app }) => {
		const refusal = (run: () => unknown): string => {
			try {
				run();
			} catch (err) {
				return (err as Error).message;
			}
			throw new Error('expected a refusal');
		};
		for (const message of [
			refusal(() => readModule(fixture(file), 'Module1')),
			refusal(() => listSubs(fixture(file), 'Module1')),
			refusal(() => writeModule(copyOf(file), 'Module1', 'Public Sub P()\r\nEnd Sub\r\n')),
			refusal(() => addFormModule(copyOf(file), 'Form1', '')),
			// A library the file is not the host of: its own is implicit, and
			// that refusal comes first and says something else.
			refusal(() => addReference(copyOf(file), app === 'Word' ? 'excel' : 'word')),
		]) {
			expect(message).toContain('has no VBA in it yet');
			expect(message).toContain(description);
			expect(message).toContain(app);
			// The internal wording never reaches a user or an agent.
			expect(message).not.toContain('vbaProject.bin');
			expect(message).not.toContain('not a valid VBA project');
		}
	});
});

describe('the rest of a file with no VBA', () => {
	it('reads and writes its document surface as usual', () => {
		const target = copyOf('NoVbaFixture.xlsm');
		expect(listSheets(target).sheets.map((sheet) => sheet.name)).toEqual(['Sheet1']);
		expect(readCells(target, 'Sheet1', 'A1').data).toEqual([['no vba']]);
		expect(writeCells(target, 'Sheet1', 'B1', [['written']])).toMatchObject({ ok: true });
		expect(readCells(target, 'Sheet1', 'B1').data).toEqual([['written']]);
		// Still no VBA afterwards: writing cells does not invent a project.
		expect(hasVbaProject(target)).toBe(false);
	});

	it('getProjectInfo still answers the sheet surface', () => {
		const info = getProjectInfo(fixture('NoVbaFixture.xlsm'));
		expect(info.sheets.map((sheet) => sheet.name)).toEqual(['Sheet1']);
		expect(info.modules).toEqual([]);
	});
});

describe('a file that genuinely could not be read', () => {
	it('is still an error, and is not confused with an empty one', () => {
		const damaged = path.join(tempRoot, 'Damaged.xlsm');
		fs.writeFileSync(damaged, Buffer.from('this is not an Office file at all', 'latin1'));
		expect(() => openMacroContainer(fs.readFileSync(damaged))).toThrow(MacroContainerError);
		// The listing surfaces answer [] only for the empty state; a file they
		// cannot read still fails, so the tree still offers its retry.
		expect(() => listModules(damaged)).toThrow();
		expect(() => hasVbaProject(damaged)).toThrow();
		expect(validateProject(damaged).issues).toHaveLength(1);
	});
});
