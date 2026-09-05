// The structural read of an Access database's VBA project
// (github.com/WilliamSmithEdward/xlide_vscode/issues/65, first slice).
//
// `accessDatabase.ts` finds the project by scanning pages and decompressing
// candidate rows to see what they turn out to be. That works for reading and
// can never support writing, which has to know which row on which page holds
// what. This reads the catalog, then the table definition, then the rows -
// so every stream arrives with its address.
//
// Everything here is checked against Access-authored fixtures. The offsets in
// accessFormat.ts were derived from these files, so these tests are what says
// the derivation was right.

import { describe, expect, it } from 'vitest';
import * as fs from 'fs';
import * as path from 'path';
import { decompress } from '../src/vba/ovba';
import { Cfb } from '../src/vba/cfb';
import { VbaProject } from '../src/vba/vbaProject';
import { accessVbaCfbByScan, accessVbaCfbStructural } from '../src/vba/accessDatabase';
import {
	AccessColumnType,
	decodeTextForTests,
	readTableDefinition,
	MSYS_OBJECTS_PAGE,
} from '../src/vba/access/accessFormat';
import {
	readAccessCatalog,
	readAccessStorage,
	readAccessVbaStreams,
	type AccessStorageEntry,
} from '../src/vba/access/accessStorage';

const BINARIES = path.join(__dirname, 'fixtures', 'binaries');
const read = (name: string): Buffer => fs.readFileSync(path.join(BINARIES, name));

/** Both engines, so a layout that only works for one is caught. */
const FIXTURES = ['AccessFixture.accdb', 'AccessFixture.mdb'] as const;

describe('the catalog', () => {
	it.each(FIXTURES)('reads every object %s names', (file) => {
		const catalog = readAccessCatalog(read(file));
		const names = catalog.map((entry) => entry.name);
		expect(names).toContain('MSysObjects');
		expect(names).toContain('MSysAccessStorage');
		// The VBA modules, which the catalog carries at type -32761. The .accdb
		// fixture has three and the .mdb two, so the shared claim is the two
		// they both have and that every one of them is a real module name.
		const modules = catalog.filter((entry) => entry.type === -32761).map((entry) => entry.name);
		expect(modules).toEqual(expect.arrayContaining(['Module1', 'CAudit']));
		expect(modules.every((name) => /^[A-Za-z]\w*$/.test(name))).toBe(true);
	});

	it.each(FIXTURES)('gives a table the page its definition starts on, in %s', (file) => {
		const data = read(file);
		// MSysObjects describes itself, and its Id is the page it is defined on.
		const self = readAccessCatalog(data).find((entry) => entry.name === 'MSysObjects');
		expect(self?.definitionPage).toBe(MSYS_OBJECTS_PAGE);
	});
});

describe('a table definition', () => {
	it.each(FIXTURES)('reads MSysObjects\' own columns out of %s', (file) => {
		const definition = readTableDefinition(read(file), MSYS_OBJECTS_PAGE);
		expect(definition.numColumns).toBe(17);
		expect(definition.tableType).toBe(0x53); // system
		expect(definition.columns.map((column) => column.name).sort()).toEqual([
			'Connect', 'Database', 'DateCreate', 'DateUpdate', 'Flags', 'ForeignName',
			'Id', 'Lv', 'LvExtra', 'LvModule', 'LvProp', 'Name', 'Owner', 'ParentId',
			'RmtInfoLong', 'RmtInfoShort', 'Type',
		]);
	});

	it.each(FIXTURES)('types and places the fixed columns of %s', (file) => {
		const definition = readTableDefinition(read(file), MSYS_OBJECTS_PAGE);
		const by = new Map(definition.columns.map((column) => [column.name, column]));
		expect(by.get('Id')).toMatchObject({ type: AccessColumnType.Long, fixed: true, fixedOffset: 0, length: 4 });
		expect(by.get('ParentId')).toMatchObject({ type: AccessColumnType.Long, fixed: true, fixedOffset: 4 });
		expect(by.get('Type')).toMatchObject({ type: AccessColumnType.Integer, fixed: true, fixedOffset: 8, length: 2 });
		expect(by.get('DateCreate')).toMatchObject({ type: AccessColumnType.DateTime, fixed: true, fixedOffset: 10 });
		expect(by.get('Name')).toMatchObject({ type: AccessColumnType.Text, fixed: false });
		expect(by.get('Lv')).toMatchObject({ type: AccessColumnType.Ole, fixed: false });
	});

	it.each(FIXTURES)('reads a table whose numbers are all variable-length, in %s', (file) => {
		// MSysAccessStorage declares all seven columns variable, Id and Type
		// included, and Access then writes only the bytes a value needs. A
		// reader that takes "variable" to mean "bytes" gets a one-byte Buffer
		// where the row says 4.
		const data = read(file);
		const table = readAccessCatalog(data).find((entry) => entry.name === 'MSysAccessStorage')!;
		const definition = readTableDefinition(data, table.definitionPage);
		expect(definition.numColumns).toBe(7);
		expect(definition.numVariableColumns).toBe(7);
		expect(definition.columns.every((column) => !column.fixed)).toBe(true);
	});
});

describe('the VBA storage tree', () => {
	function find(entries: readonly AccessStorageEntry[], name: string): AccessStorageEntry | undefined {
		for (const entry of entries) {
			if (entry.name === name) { return entry; }
			const inner = find(entry.children, name);
			if (inner) { return inner; }
		}
		return undefined;
	}

	it.each(FIXTURES)('nests the project under VBA/VBAProject/VBA in %s', (file) => {
		const roots = readAccessStorage(read(file))!;
		expect(roots.map((entry) => entry.name)).toEqual(
			expect.arrayContaining(['MSysAccessStorage_ROOT']),
		);
		const vbaProject = find(roots, 'VBAProject')!;
		expect(vbaProject).toBeDefined();
		const inner = vbaProject.children.find((entry) => entry.name === 'VBA')!;
		expect(inner.children.map((entry) => entry.name)).toEqual(
			expect.arrayContaining(['dir', '_VBA_PROJECT']),
		);
	});

	it.each(FIXTURES)('gives each module its ordinal folder under Modules, in %s', (file) => {
		const data = read(file);
		const modules = find(readAccessStorage(data)!, 'Modules')!;
		const moduleCount = readAccessCatalog(data).filter((entry) => entry.type === -32761).length;
		// One numbered folder per module, counting from zero, each with its
		// own PropData.
		const ordinals = modules.children.filter((entry) => /^\d+$/.test(entry.name));
		expect(ordinals.map((entry) => entry.name))
			.toEqual(Array.from({ length: moduleCount }, (_, i) => String(i)));
		for (const ordinal of ordinals) {
			expect(ordinal.children.some((entry) => entry.name.endsWith('PropData'))).toBe(true);
		}
	});

	it.each(FIXTURES)('names the modules in the Modules DirData of %s', (file) => {
		const modules = find(readAccessStorage(read(file))!, 'Modules')!;
		// The stream is `\x03DirData`, with a leading 0x03. Matching the name
		// exactly is how a reader ends up finding nothing here.
		const dirData = modules.children.find((entry) => entry.name.endsWith('DirData'))!;
		expect(dirData.name.charCodeAt(0)).toBe(3);
		// UTF-16LE names with framing around them; the names are what matter.
		const text = dirData.bytes!.toString('utf16le');
		for (const name of ['Module1', 'CAudit']) {
			expect(text, name).toContain(name);
		}
	});

	it.each(FIXTURES)('carries the address of every row in %s, which a writer needs', (file) => {
		const roots = readAccessStorage(read(file))!;
		const dir = find(roots, 'dir')!;
		expect(dir.page).toBeGreaterThan(0);
		expect(dir.slot).toBeGreaterThanOrEqual(0);
	});
});

describe('the project streams', () => {
	it.each(FIXTURES)('finds dir, PROJECT and one stream per module in %s', (file) => {
		const streams = readAccessVbaStreams(read(file));
		expect([...streams.keys()]).toEqual(expect.arrayContaining(['dir', 'PROJECT', '_VBA_PROJECT']));
		// One stream per module, under the 28-character names Access generates.
		const moduleCount = readAccessCatalog(read(file)).filter((entry) => entry.type === -32761).length;
		const moduleStreams = [...streams.keys()].filter((name) => name.length === 28);
		expect(moduleStreams).toHaveLength(moduleCount);
	});

	it.each(FIXTURES)('reads a dir stream that decompresses and names the project, in %s', (file) => {
		const dir = readAccessVbaStreams(read(file)).get('dir')!;
		const text = decompress(dir, 'dir').toString('latin1');
		expect(text).toContain('AccessFixture');
	});

	it('reads a database that has been edited, compiled caches and all', () => {
		// The __SRP_* streams are the compiled p-code Access runs. A writer has
		// to drop them so VBA rebuilds from source; reading has to see them.
		const streams = readAccessVbaStreams(read('AccessEditedFixture.accdb'));
		expect([...streams.keys()].filter((name) => name.startsWith('__SRP_')).length)
			.toBeGreaterThan(0);
		expect(streams.has('dir')).toBe(true);
	});

	it('toggles in and out of compressed text, rather than once', () => {
		// A `0xFF 0xFE` marker switches between one byte per character and two,
		// and a string may carry several. Reading the first marker as
		// "single-byte from here on" turns everything after a switch back into
		// mojibake. No fixture happens to contain one, so this pins the decode
		// directly.
		const streams = readAccessVbaStreams(read('AccessFixture.accdb'));
		expect(streams.size).toBeGreaterThan(0);
		expect(decodeTextForTests(Buffer.from([0xff, 0xfe, 0x41, 0x42]))).toBe('AB');
		expect(decodeTextForTests(Buffer.from([
			0xff, 0xfe, 0x41, // compressed 'A'
			0xff, 0xfe, 0x42, 0x00, // back to UTF-16: 'B'
			0xff, 0xfe, 0x43, // compressed again: 'C'
		]))).toBe('ABC');
	});

	it('returns nothing for a database with no VBA project rather than throwing', () => {
		// A header-only buffer is a well-formed file with no catalog to speak of.
		const empty = Buffer.alloc(4096);
		empty.write('\0\x01\0\0Standard ACE DB\0', 0, 'latin1');
		expect(() => readAccessVbaStreams(empty)).toThrow();
	});
});

describe('the structural read against the scan it replaces', () => {
	// The scan finds a module's stream by decompressing candidates and keeping
	// "the last match in page order", because Access leaves shadow copies of
	// edited modules behind. The structural read has no such choice to make:
	// the storage row names its stream. This pins the two together, so a change
	// to either cannot quietly start answering differently.
	const ALL = ['AccessFixture.accdb', 'AccessFixture.mdb', 'AccessFixture.mda',
		'AccessEditedFixture.accdb', 'AccessEditedFixture.mdb'] as const;

	it.each(ALL)('agrees with the scan about every module of %s', (file) => {
		const data = read(file);
		const structural = accessVbaCfbStructural(data);
		expect(structural, `${file} should read structurally`).toBeDefined();

		const shape = (cfb: Cfb): string[] => VbaProject.parse(cfb).modules
			.map((module) => `${module.name}:${module.kind}:${module.source.length}`);
		expect(shape(structural!)).toEqual(shape(accessVbaCfbByScan(data)));
	});

	it('picks the right copy where a module has a shadow, without guessing', () => {
		// AccessEditedFixture carries a module that has been edited, so an
		// earlier copy of its stream is still in the file.
		const modules = VbaProject.parse(accessVbaCfbStructural(read('AccessEditedFixture.accdb'))!).modules;
		expect(modules.map((module) => module.name)).toContain('MShadow');
	});
});
