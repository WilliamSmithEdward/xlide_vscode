// Writing the VBA project of an Access database
// (github.com/WilliamSmithEdward/xlide_vscode/issues/65).
//
// Access runs the compiled project rather than the source, so a source write
// counts only once the compiled cache is marked stale; the writer does that,
// and Access recompiles on the next open, which is what its own /decompile
// switch does. A module costs rows in five storage folders plus rows in three
// catalog tables, and every one of them has to agree - a module listed in one
// and missing from another is one Access will show and then refuse to open.
//
// These read the result back through the reader, which walks the same
// structures Access does. The live check, that Access itself opens the file,
// compiles it and runs the new code, is not something CI can host; it was run
// against Access 16.0 for each operation below.

import { describe, expect, it } from 'vitest';
import * as fs from 'fs';
import * as path from 'path';
import { AccessVbaWriter, applyAccessVbaProject } from '../src/vba/access/accessVbaWriter';
import { readAccessCatalog, readAccessStorage } from '../src/vba/access/accessStorage';
import { accessVbaCfb } from '../src/vba/accessDatabase';
import { VbaProject } from '../src/vba/vbaProject';
import { AccessPageStore } from '../src/vba/access/accessPageStore';
import { AccessTable } from '../src/vba/access/accessTableWriter';
import { decodeRowValues } from '../src/vba/access/accessFormat';

const BINARIES = path.join(__dirname, 'fixtures', 'binaries');
const read = (name: string): Buffer => fs.readFileSync(path.join(BINARIES, name));

/** Both engines and both shapes: with and without the compiled `__SRP_` rows. */
const FIXTURES = [
	'AccessFixture.accdb',
	'AccessFixture.mdb',
	'AccessEditedFixture.accdb',
	'AccessEditedFixture.mdb',
	'AccessFormFixture.accdb',
] as const;

/** The one line a rename changes, so the rest of a module can be compared. */
const NAME_ATTRIBUTE = /^Attribute VB_Name = [^\n]*\n/;

const BODY = [
	'Option Compare Database',
	'Option Explicit',
	'',
	'Public Function Answer() As Long',
	'    Answer = 42',
	'End Function',
	'',
].join('\r\n');

/** The project as the rest of XLIDE sees it, read back out of the bytes. */
function project(data: Buffer): VbaProject {
	return VbaProject.parse(accessVbaCfb(data));
}

/** Every table the catalog names still opens, and its row count still adds up. */
function assertConsistent(data: Buffer): void {
	const store = new AccessPageStore(data);
	for (const entry of readAccessCatalog(data)) {
		if (entry.type !== 1) {
			continue;
		}
		let table: AccessTable;
		try {
			table = new AccessTable(store, entry.definitionPage);
		} catch {
			continue;
		}
		// The definition's counter is what Access trusts for a row count, so a
		// write that leaves it behind leaves the table wrong.
		expect(table.rows().length, `${entry.name} row count`).toBe(table.definition.numRows);
	}
	// The storage tree still walks, and every stream still resolves.
	expect(readAccessStorage(data)).toBeDefined();
}

/** What the storage's own listings say the project holds. */
function storageNames(data: Buffer, container: string): string[] {
	const store = new AccessPageStore(data);
	const catalog = readAccessCatalog(data)
		.find((entry) => entry.name === 'MSysAccessStorage')!;
	const table = new AccessTable(store, catalog.definitionPage);
	const rows = table.rows().map((row) => decodeRowValues(row.bytes, table.definition));
	const folder = rows.find((values) => values?.get('Name') === container);
	const id = folder?.get('Id');
	return rows
		.filter((values) => values?.get('ParentId') === id)
		.map((values) => String(values?.get('Name')));
}

describe('replacing a module', () => {
	it.each(FIXTURES)('%s: the new source is what reads back', (file) => {
		const writer = new AccessVbaWriter(read(file));
		const name = writer.moduleNames()[0];
		writer.setModuleText(name, BODY);
		const data = writer.toBuffer();
		const module = project(data).getModule(name);
		expect(module?.source).toContain('Answer = 42');
		expect(module?.source).toContain(`Attribute VB_Name = "${name}"`);
		assertConsistent(data);
	});

	it.each(FIXTURES)('%s: the other modules are untouched', (file) => {
		const data = read(file);
		const before = project(data);
		const writer = new AccessVbaWriter(data);
		const [name, ...others] = writer.moduleNames();
		writer.setModuleText(name, BODY);
		const after = project(writer.toBuffer());
		expect(after.modules.map((module) => module.name)).toEqual(before.modules.map((m) => m.name));
		for (const other of others) {
			expect(after.getModule(other)?.source).toBe(before.getModule(other)?.source);
		}
	});

	it('keeps the attribute block when the new text carries none', () => {
		const writer = new AccessVbaWriter(read('AccessFormFixture.accdb'));
		writer.setModuleText('basket', 'Public Sub P()\r\nEnd Sub\r\n');
		const source = project(writer.toBuffer()).getModule('basket')?.source ?? '';
		// A class module's seven extra attributes are what make it
		// instantiable; a body pasted in from elsewhere must not strip them.
		expect(source).toContain('Attribute VB_Name = "Basket"');
		expect(source).toContain('Attribute VB_PredeclaredId = False');
		expect(source).toContain('Public Sub P()');
	});

	it('marks the compiled project stale so Access recompiles from source', () => {
		const writer = new AccessVbaWriter(read('AccessFixture.accdb'));
		writer.setModuleText('Module1', BODY);
		const streams = accessVbaCfb(writer.toBuffer());
		const cache = streams.getStream('_VBA_PROJECT');
		// A version the host does not know is what makes VBA discard the cache.
		expect(cache.readUInt16LE(2)).toBe(0x0099);
	});

	it('drops the compiled __SRP_ rows, which Access prefers to the p-code', () => {
		const data = read('AccessEditedFixture.accdb');
		const srp = (bytes: Buffer): number => (readAccessStorage(bytes) ?? [])
			.flatMap(function walk(entry): string[] {
				return [entry.name, ...entry.children.flatMap(walk)];
			})
			.filter((name) => name.startsWith('__SRP_')).length;
		expect(srp(data)).toBeGreaterThan(0);
		const writer = new AccessVbaWriter(data);
		writer.setModuleText('MShadow', BODY);
		const written = writer.toBuffer();
		expect(srp(written)).toBe(0);
		assertConsistent(written);
	});
});

describe('adding a module', () => {
	it.each(FIXTURES)('%s: the project lists it and its source reads back', (file) => {
		const writer = new AccessVbaWriter(read(file));
		writer.addModule('MAdded', BODY);
		const data = writer.toBuffer();
		expect(project(data).getModule('MAdded')?.source).toContain('Answer = 42');
		expect(writer.moduleNames()).toContain('MAdded');
		assertConsistent(data);
	});

	it.each(FIXTURES)('%s: every listing the module has to appear in has it', (file) => {
		const writer = new AccessVbaWriter(read(file));
		writer.addModule('MAdded', BODY);
		const data = writer.toBuffer();
		const streams = accessVbaCfb(data);
		// PROJECT is what tells Access the module exists at all, and PROJECTwm
		// carries its name for the VBE.
		const text = streams.getStream('PROJECT').toString('latin1');
		expect(text).toContain('Module=MAdded');
		expect(streams.getStream('PROJECTwm').toString('latin1')).toContain('MAdded');
		// A new module also takes a numbered storage folder under Modules;
		// Access finds it by that name and by no other.
		const before = storageNames(read(file), 'Modules');
		const after = storageNames(data, 'Modules');
		expect(after.length).toBe(before.length + 1);
		// The catalog row is what the navigation pane and AllModules read.
		const catalog = readAccessCatalog(data)
			.filter((entry) => entry.type === -32761).map((entry) => entry.name);
		expect(catalog).toContain('MAdded');
	});

	it('gives a class module the attributes that make it instantiable', () => {
		const writer = new AccessVbaWriter(read('AccessFixture.accdb'));
		writer.addModule('CAdded', 'Public Sub P()\r\nEnd Sub\r\n', 'class');
		const source = project(writer.toBuffer()).getModule('CAdded')?.source ?? '';
		expect(source).toContain('Attribute VB_Name = "CAdded"');
		expect(source).toContain('Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"');
		expect(source).toContain('Attribute VB_Creatable = False');
		expect(project(writer.toBuffer()).getModule('CAdded')?.kind).toBe('other');
	});

	it('refuses a name the project already holds', () => {
		const writer = new AccessVbaWriter(read('AccessFixture.accdb'));
		expect(() => writer.addModule('Module1', BODY)).toThrow(/already exists/);
	});

	it('refuses a name longer than Access allows', () => {
		const writer = new AccessVbaWriter(read('AccessFixture.accdb'));
		expect(() => writer.addModule('M'.repeat(65), BODY)).toThrow(/1 to 64/);
	});
});

describe('renaming a module', () => {
	it.each(FIXTURES)('%s: the new name reaches every listing', (file) => {
		const writer = new AccessVbaWriter(read(file));
		const name = writer.moduleNames()[0];
		writer.renameModule(name, 'MRenamed');
		const data = writer.toBuffer();
		expect(writer.moduleNames()).toContain('MRenamed');
		expect(writer.moduleNames()).not.toContain(name);
		// The name lives in the module's own attribute block too.
		expect(project(data).getModule('MRenamed')?.source)
			.toContain('Attribute VB_Name = "MRenamed"');
		expect(accessVbaCfb(data).getStream('PROJECT').toString('latin1')).not.toContain(`=${name}\r`);
		expect(readAccessCatalog(data).filter((entry) => entry.type === -32761)
			.map((entry) => entry.name)).toContain('MRenamed');
		assertConsistent(data);
	});

	it('keeps the body it had', () => {
		const writer = new AccessVbaWriter(read('AccessFixture.accdb'));
		const before = writer.moduleText('Module1');
		writer.renameModule('Module1', 'MRenamed');
		const after = writer.moduleText('MRenamed');
		expect(after.replace(NAME_ATTRIBUTE, ''))
			.toBe(before.replace(NAME_ATTRIBUTE, ''));
	});
});

describe('deleting a module', () => {
	it.each(FIXTURES)('%s: it leaves no trace in any listing', (file) => {
		const writer = new AccessVbaWriter(read(file));
		writer.addModule('MDoomed', BODY);
		writer.deleteModule('MDoomed');
		const data = writer.toBuffer();
		expect(writer.moduleNames()).not.toContain('MDoomed');
		expect(accessVbaCfb(data).getStream('PROJECT').toString('latin1'))
			.not.toContain('MDoomed');
		expect(readAccessCatalog(data).filter((entry) => entry.type === -32761)
			.map((entry) => entry.name)).not.toContain('MDoomed');
		// The storage folder goes back with it.
		expect(storageNames(data, 'Modules').length)
			.toBe(storageNames(read(file), 'Modules').length);
		assertConsistent(data);
	});

	it('leaves the other modules readable', () => {
		const data = read('AccessFormFixture.accdb');
		const before = project(data);
		const writer = new AccessVbaWriter(data);
		writer.addModule('MDoomed', BODY);
		writer.deleteModule('MDoomed');
		const after = project(writer.toBuffer());
		expect(after.modules.map((module) => module.name))
			.toEqual(before.modules.map((module) => module.name));
		for (const module of before.modules) {
			expect(after.getModule(module.name)?.source).toBe(module.source);
		}
	});

	it('refuses to delete the module behind a form, which has no folder', () => {
		// A document module belongs to its design, not to Modules: it has no
		// storage folder and no catalog row, so removing it means removing the
		// form. Deleting it as a module would take another module's folder.
		const writer = new AccessVbaWriter(read('AccessFormFixture.accdb'));
		expect(() => writer.deleteModule('Form_Calculator'))
			.toThrow(/behind a form or report/);
	});
});

describe('names and source outside the code page', () => {
	// Every ANSI string in the project is in PROJECTCODEPAGE, and a name the
	// page cannot hold is carried in full by the unicode record beside it.
	// Both of these have been user-visible bugs (issue #6).
	const NAMES: Array<[string, string]> = [
		['Cyrillic', 'Модуль1'],
		['Japanese', 'モジュール'],
		['Greek', 'Μονάδα'],
		['Hebrew', 'מודול'],
		['Vietnamese', 'MôĐun'],
		['accented Latin', 'Modül'],
	];

	it.each(NAMES)('adds a module named in %s and finds it again', (_label, name) => {
		const writer = new AccessVbaWriter(read('AccessFixture.accdb'));
		writer.addModule(name, BODY);
		expect(writer.moduleNames()).toContain(name);
		expect(writer.moduleText(name)).toContain('Answer = 42');
		const data = writer.toBuffer();
		// The reader speaks the same name, through its own dir parse.
		expect(project(data).modules.map((module) => module.name)).toContain(name);
		assertConsistent(data);
	});

	it.each(NAMES)('renames a module to %s and back', (_label, name) => {
		const writer = new AccessVbaWriter(read('AccessFixture.accdb'));
		const body = writer.moduleText('Module1').replace(NAME_ATTRIBUTE, '');
		writer.renameModule('Module1', name);
		expect(writer.moduleNames()).toContain(name);
		expect(writer.moduleText(name).replace(NAME_ATTRIBUTE, '')).toBe(body);
		writer.renameModule(name, 'Module1');
		expect(writer.moduleNames()).toContain('Module1');
		expect(writer.moduleText('Module1')).toContain('Attribute VB_Name = "Module1"');
		assertConsistent(writer.toBuffer());
	});

	it('folds the name in the VB_Name attribute, which is ANSI only', () => {
		// The dir stream carries a name twice, ANSI and unicode, so the real
		// name survives there. A module's source has no such twin, so the
		// VB_Name attribute holds the folded projection - and Access takes the
		// module's name from the dir stream, not from the attribute.
		const writer = new AccessVbaWriter(read('AccessFixture.accdb'));
		writer.renameModule('Module1', 'Модуль1');
		expect(writer.moduleNames()).toContain('Модуль1');
		expect(writer.moduleText('Модуль1')).toContain('Attribute VB_Name = "??????1"');
	});

	it('renames between two names the code page cannot hold', () => {
		const writer = new AccessVbaWriter(read('AccessFormFixture.accdb'));
		writer.addModule('Модуль1', BODY);
		writer.renameModule('Модуль1', 'モジュール');
		expect(writer.moduleNames()).toContain('モジュール');
		expect(writer.moduleNames()).not.toContain('Модуль1');
		expect(project(writer.toBuffer()).modules.map((module) => module.name))
			.toContain('モジュール');
		assertConsistent(writer.toBuffer());
	});

	it('round-trips source the code page can hold', () => {
		// cp1252 keeps the em dash, the curly quotes and the euro sign in
		// 0x80-0x9F, where latin-1 has control characters instead.
		const text = 'Public Sub P()\r\n    \' — “quoted” café naïve €\r\nEnd Sub\r\n';
		const writer = new AccessVbaWriter(read('AccessFixture.accdb'));
		writer.setModuleText('Module1', text);
		expect(writer.moduleText('Module1')).toContain('— “quoted” café naïve €');
		expect(project(writer.toBuffer()).getModule('Module1')?.source)
			.toContain('— “quoted” café naïve €');
	});

	it('folds source the code page cannot hold, as the VBE does', () => {
		// A cp1252 project cannot store Cyrillic in a module's source: there is
		// no unicode twin for source the way there is for a name.
		const writer = new AccessVbaWriter(read('AccessFixture.accdb'));
		writer.setModuleText('Module1', 'Public Sub P()\r\n    \' Привет\r\nEnd Sub\r\n');
		const source = writer.moduleText('Module1');
		expect(source).toContain('??????');
		expect(source).not.toContain('Привет');
	});
});

describe('the whole project through the editor', () => {
	it.each(FIXTURES)('%s: a save that changed nothing changes no bytes', (file) => {
		const data = read(file);
		expect(applyAccessVbaProject(data, accessVbaCfb(data)).equals(data)).toBe(true);
	});
});
