import { afterEach, describe, expect, it } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { Cfb } from '../src/vba/cfb';
import { openMacroContainer } from '../src/vba/macroContainer';
import { compress, decompress } from '../src/vba/ovba';
import { readDirRecords, REC_PROJECTMODULES } from '../src/vba/vbaProject';
import { addFormModule, createProject, listModules, readFormMarkup, validateProject } from '../src/vba/projectService';
import { hasMsFormsReference, readProjectReferences } from '../src/vba/vbaProjectReferences';

// Creating the first UserForm in a project used to leave it without the
// Microsoft Forms library: the form and its designer storage were written and
// registered, the file read perfectly here, and the host could not
// instantiate the form or compile the project. Every blank template starts in
// that state, in all three hosts that have UserForms, so they are the
// reproduction.

const TEMPLATES = [
	['Excel', 'blank.xlsm', 'Book.xlsm'],
	['Word', 'blank.docm', 'Doc.docm'],
	['PowerPoint', 'blank.pptm', 'Deck.pptm'],
] as const;

function templatePath(name: string): string {
	return path.join('assets', 'templates', name);
}

function dirStream(file: string): Buffer {
	const container = openMacroContainer(fs.readFileSync(file));
	return decompress(container.vbaCfb().getStreamInStorage('VBA', 'dir'), 'VBA/dir');
}

/** Writes a dir stream back, for building the state an older XLIDE left behind. */
function replaceDirStream(file: string, dir: Buffer): void {
	const container = openMacroContainer(fs.readFileSync(file));
	const cfb = container.vbaCfb();
	cfb.writeStreamInStorage('VBA', 'dir', compress(dir));
	fs.writeFileSync(file, container.toFileBytes(cfb));
}

/** The same stream with its Microsoft Forms records cut out. */
function withoutMsForms(dir: Buffer): Buffer {
	const records = readDirRecords(dir);
	const start = records.find((r) => r.id === 0x0016
		&& dir.subarray(r.dataStart, r.dataEnd).toString('latin1') === 'MSForms');
	const modules = records.find((r) => r.id === REC_PROJECTMODULES);
	expect(start, 'the project declares Microsoft Forms').toBeDefined();
	expect(modules, 'the project has a module section').toBeDefined();
	return Buffer.concat([dir.subarray(0, start!.start), dir.subarray(modules!.start)]);
}

describe('a project that gains its first form gains the library it needs', () => {
	let dir: string | undefined;
	afterEach(() => {
		if (dir) { fs.rmSync(dir, { recursive: true, force: true }); }
		dir = undefined;
	});

	function newProject(template: string, name: string): string {
		dir = fs.mkdtempSync(path.join(os.tmpdir(), 'xlide-form-ref-'));
		const file = path.join(dir, name);
		createProject(file, templatePath(template));
		return file;
	}

	it('starts from templates that have no form and no Microsoft Forms reference', () => {
		for (const [host, template] of TEMPLATES) {
			const stream = dirStream(templatePath(template));
			expect(hasMsFormsReference(stream), host).toBe(false);
			expect(readProjectReferences(stream).map((r) => r.name), host).toContain('stdole');
			expect((listModules(templatePath(template)) as Array<{ type: string }>)
				.some((m) => m.type === 'userform'), host).toBe(false);
		}
	});

	for (const [host, template, fileName] of TEMPLATES) {
		it(`adds the reference with the form in a ${host} project, and the form still reads back`, () => {
			const file = newProject(template, fileName);
			expect(hasMsFormsReference(dirStream(file))).toBe(false);
			const before = readProjectReferences(dirStream(file)).map((r) => r.name);

			addFormModule(file, 'EntryForm');

			expect(hasMsFormsReference(dirStream(file))).toBe(true);
			// Everything the project already declared survives, in order.
			expect(readProjectReferences(dirStream(file)).map((r) => r.name)).toEqual([...before, 'MSForms']);
			expect((listModules(file) as Array<{ name: string; type: string }>)
				.find((m) => m.name === 'EntryForm')?.type).toBe('userform');
			expect(readFormMarkup(file, 'EntryForm').markup).toContain('EntryForm');
			expect(validateProject(file).issues).toEqual([]);
		});
	}

	it('does not add a second reference for a second form', () => {
		const file = newProject('blank.xlsm', 'Book.xlsm');
		addFormModule(file, 'FirstForm');
		const afterFirst = readProjectReferences(dirStream(file));
		addFormModule(file, 'SecondForm');
		const afterSecond = readProjectReferences(dirStream(file));
		expect(afterSecond.map((r) => r.name)).toEqual(afterFirst.map((r) => r.name));
		expect(afterSecond.filter((r) => r.name === 'MSForms')).toHaveLength(1);
		expect(validateProject(file).issues).toEqual([]);
	});

	it('reports a form whose project never declared the library', () => {
		const file = newProject('blank.xlsm', 'Book.xlsm');
		addFormModule(file, 'EntryForm');
		expect(validateProject(file).issues).toEqual([]);

		// The state an earlier XLIDE left behind: the form is there, the
		// library reference is not.
		replaceDirStream(file, withoutMsForms(dirStream(file)));
		expect(hasMsFormsReference(dirStream(file))).toBe(false);

		const issues = validateProject(file).issues;
		expect(issues).toHaveLength(1);
		expect(issues[0]).toContain('EntryForm');
		expect(issues[0]).toContain('Microsoft Forms');
		expect(issues[0]).toContain('Tools > References');
		// The message belongs to no single host: Word and PowerPoint forms
		// need the same library.
		expect(issues[0]).not.toContain('Excel');
		// The modules themselves are still readable, which is why nothing else
		// here notices the problem.
		expect((listModules(file) as Array<{ name: string }>).map((m) => m.name)).toContain('EntryForm');
	});
});
