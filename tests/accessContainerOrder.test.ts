import { describe, expect, it } from 'vitest';
import * as fs from 'fs';
import * as path from 'path';
import { readAccessStorage, readTableRows, type AccessStorageEntry } from '../src/vba/access/accessStorage';
import { MSYS_OBJECTS_PAGE, readTableDefinition } from '../src/vba/access/accessFormat';
import {
	accessListOrder,
	accessNameHash,
	dirDataEntries,
	nextFolderName,
	refileInFolderList,
	removeFromFolderList,
} from '../src/vba/access/accessVbaStreams';
import { AccessVbaWriter } from '../src/vba/access/accessVbaWriter';

// Issue #150, from pyOpenVBA 6.3.0, re-measured in Access 16.0 driven over
// COM. AccessTwelveModulesFixture is what Access wrote after Module1 was
// added to a blank database and then Macro1 to Macro11 in one session. The
// expected lists below are the ones Access wrote when it then deleted Macro3,
// added NewOne and renamed Macro5 to Zeta, reopening the database between.

const FIXTURE = path.join(__dirname, 'fixtures', 'binaries', 'AccessTwelveModulesFixture.accdb');

function find(entries: AccessStorageEntry[], name: string): AccessStorageEntry | undefined {
	for (const entry of entries) {
		if (entry.name === name) { return entry; }
		const inner = find(entry.children, name);
		if (inner) { return inner; }
	}
	return undefined;
}

/** The Modules container's folders, DirData list and catalog ids. */
function modulesContainer(data: Buffer): { folders: string[]; listed: string[]; ids: Map<string, number> } {
	const modules = find(readAccessStorage(data) ?? [], 'Modules')!;
	const folders = modules.children.filter((c) => c.type === 1).map((c) => c.name)
		.sort((a, b) => Number(a) - Number(b));
	const listing = modules.children.find((c) => c.name === '\x03DirData')!;
	const listed = dirDataEntries(listing.bytes!).map((e) => `${e.name}@${e.folder}`);
	const objects = readTableRows(data, readTableDefinition(data, MSYS_OBJECTS_PAGE));
	const container = objects.find((row) => row.values.get('Name') === 'Modules' && row.values.get('Type') === 3)!;
	const ids = new Map<string, number>();
	for (const row of objects) {
		if (row.values.get('ParentId') === container.values.get('Id')) {
			ids.set(row.values.get('Name') as string, row.values.get('Id') as number);
		}
	}
	return { folders, listed, ids };
}

const applied = (data: Buffer, act: (writer: AccessVbaWriter) => void): Buffer => {
	const writer = new AccessVbaWriter(data);
	act(writer);
	return writer.toBuffer();
};

describe('the order Access keeps a container\'s lists in (issue #150)', () => {
	it('hashes a name ignoring case and a leading dot, counting a space as 0', () => {
		expect(accessNameHash('Module1')).toBe(accessNameHash('MODULE1'));
		expect(accessNameHash('.Module1')).toBe(accessNameHash('Module1'));
		expect(accessNameHash('A B')).not.toBe(accessNameHash('AB'));
		expect(accessNameHash('A"B')).toBe(accessNameHash('AB'));
		expect(accessNameHash('Module1') & 7).toBe(accessNameHash('Macro7') & 7);
	});

	it('orders Macro1 to Macro11 as Access does, into an empty list and after Module1', () => {
		const macros = Array.from({ length: 11 }, (_, i) => `Macro${i + 1}`);
		expect(accessListOrder([], [], macros).slice(0, 5)).toEqual(['Macro1', 'Macro10', 'Macro2', 'Macro11', 'Macro3']);
		// One session keeps its map in memory; this is the fixture's own list.
		expect(accessListOrder(['Module1'], [], macros)).toEqual([
			'Module1', 'Macro7', 'Macro1', 'Macro10', 'Macro2', 'Macro11', 'Macro3', 'Macro4', 'Macro5', 'Macro6', 'Macro8', 'Macro9',
		]);
	});

	it('names folders in decimal, lowest free first, from 0', () => {
		expect(nextFolderName(new Set())).toBe('0');
		expect(nextFolderName(new Set(['0', '1', '2', '3', '4', '5', '6', '7', '8', '9']))).toBe('10');
		expect(nextFolderName(new Set(['0', '1', '2', '4', '10']))).toBe('3');
	});

	it('reorders a PropData list with two-digit folder lines, and leaves one without the folder alone', () => {
		const line = (folder: string): Buffer => {
			const name = Buffer.from(folder, 'utf16le');
			return Buffer.concat([Buffer.from([0x05, 1 + name.length + 6, name.length]), name, Buffer.from('CB0', 'utf16le')]);
		};
		const list = (folders: string[]): Buffer => Buffer.concat([Buffer.alloc(4), ...folders.map(line)]);
		const stored = ['0', '10', '7', '1', '11', '2', '4', '5', '6', '8', '9', '3'];
		expect(removeFromFolderList(list(stored), '10'))
			.toEqual(list(accessListOrder(stored, ['10'])));
		expect(refileInFolderList(list(stored), '5'))
			.toEqual(list(accessListOrder(stored, ['5'], ['5'])));
		expect(removeFromFolderList(list(['0']), '3')).toEqual(list(['0']));
	});
});

describe('the writer files modules as Access 16.0 does (issue #150)', () => {
	const twelve = fs.readFileSync(FIXTURE);

	it('starts from the database Access wrote', () => {
		const { folders, listed } = modulesContainer(twelve);
		expect(folders).toEqual(['0', '1', '2', '3', '4', '5', '6', '7', '8', '9', '10', '11']);
		expect(listed).toEqual([
			'Module1@0', 'Macro7@7', 'Macro1@1', 'Macro10@10', 'Macro2@2', 'Macro11@11',
			'Macro3@3', 'Macro4@4', 'Macro5@5', 'Macro6@6', 'Macro8@8', 'Macro9@9',
		]);
	});

	it('deletes, reuses the lowest free folder, takes the next id and renames with Access\'s lists', () => {
		const deleted = applied(twelve, (w) => w.deleteModule('Macro3'));
		expect(modulesContainer(deleted).listed).toEqual([
			'Module1@0', 'Macro7@7', 'Macro1@1', 'Macro10@10', 'Macro2@2', 'Macro11@11',
			'Macro4@4', 'Macro5@5', 'Macro6@6', 'Macro8@8', 'Macro9@9',
		]);

		const added = applied(deleted, (w) => w.addModule('NewOne', 'Option Compare Database'));
		const afterAdd = modulesContainer(added);
		expect(afterAdd.listed.at(-1)).toBe('NewOne@3');
		// The highest id plus one: Macro3's freed id is not reused.
		expect(afterAdd.ids.get('NewOne')).toBe(afterAdd.ids.get('Macro11')! + 1);
		expect(afterAdd.ids.get('NewOne')).toBe(-2147483626);

		const renamed = applied(added, (w) => w.renameModule('Macro5', 'Zeta'));
		expect(modulesContainer(renamed).listed).toEqual([
			'Module1@0', 'Macro7@7', 'Macro1@1', 'Macro10@10', 'Macro2@2', 'Macro11@11',
			'Macro4@4', 'Macro6@6', 'Macro8@8', 'Macro9@9', 'NewOne@3', 'Zeta@5',
		]);
	});

	it('writes no PropData line for a new module, which Access adds when it next opens the file', () => {
		const propData = (data: Buffer): Buffer | undefined =>
			find(readAccessStorage(data) ?? [], 'Modules')!.children.find((c) => c.name === 'PropData')?.bytes;
		const before = propData(twelve);
		const after = propData(applied(twelve, (w) => w.addModule('Thirteenth', 'Option Compare Database')));
		expect(after).toEqual(before);
	});

	it('gives a folder past 9 its decimal name', () => {
		const added = applied(twelve, (w) => w.addModule('Thirteenth', 'Option Compare Database'));
		const { folders, listed } = modulesContainer(added);
		expect(folders.at(-1)).toBe('12');
		expect(listed).toContain('Thirteenth@12');
	});
});
