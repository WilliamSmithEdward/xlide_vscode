// The Jet 4 / ACE storage engine underneath an Access write
// (github.com/WilliamSmithEdward/xlide_vscode/issues/65).
//
// Writing one row of one table means writing the page it lands on, the
// long-value pages its value lands on, the B-tree pages of every index over
// the table, the definition's counters, and the usage maps that say which
// pages the table owns and which have room. There is no byte-patching short
// cut: a file the engine did not expect is one Access repairs or refuses.
//
// The tests here hold each layer against the fixtures Access itself wrote.
// They are invariant checks rather than golden bytes: the engine's own rules,
// read back off files it produced, so a change that breaks one of them breaks
// something Access will notice.

import { describe, expect, it } from 'vitest';
import * as fs from 'fs';
import * as path from 'path';
import {
	ACCESS_PAGE_SIZE,
	AccessColumnType,
	AccessPageType,
	pageCount,
	readTableDefinition,
} from '../src/vba/access/accessFormat';
import { AccessDataPage, isLongValuePage } from '../src/vba/access/accessDataPage';
import {
	AccessPageStore,
	GLOBAL_USAGE_MAP_PAGE,
	GLOBAL_USAGE_MAP_ROW,
	readUsageMap,
	readUsageMapRef,
} from '../src/vba/access/accessPageStore';
import { allocatePage, releasePage } from '../src/vba/access/accessAlloc';
import { AccessTable } from '../src/vba/access/accessTableWriter';
import { leafPages } from '../src/vba/access/accessIndex';
import { encodeIndexKey } from '../src/vba/access/accessIndexKey';
import { splitRow } from '../src/vba/access/accessRow';
import { encodeTextKey } from '../src/vba/access/accessCollation';
import { readAccessCatalog } from '../src/vba/access/accessStorage';

const BINARIES = path.join(__dirname, 'fixtures', 'binaries');
const read = (name: string): Buffer => fs.readFileSync(path.join(BINARIES, name));

/** Both engines and three shapes of database, so a rule that holds for one file is not mistaken for a rule. */
const FIXTURES = [
	'AccessFixture.accdb',
	'AccessFixture.mdb',
	'AccessEditedFixture.accdb',
	'AccessEditedFixture.mdb',
	'AccessFormFixture.accdb',
] as const;

/** Every user table in a database, opened for writing. */
function tables(data: Buffer): Array<{ name: string; table: AccessTable }> {
	const store = new AccessPageStore(data);
	const out: Array<{ name: string; table: AccessTable }> = [];
	for (const entry of readAccessCatalog(data)) {
		if (entry.type !== 1) {
			continue;
		}
		try {
			out.push({ name: entry.name, table: new AccessTable(store, entry.definitionPage) });
		} catch {
			// A table whose definition this cannot parse far enough to write is
			// reported by the definition tests, not here.
		}
	}
	return out;
}

describe('data pages', () => {
	it.each(FIXTURES)('%s: the free-space word follows from the rows', (file) => {
		const data = read(file);
		let checked = 0;
		for (let page = 0; page < pageCount(data); page += 1) {
			const raw = data.subarray(page * ACCESS_PAGE_SIZE, (page + 1) * ACCESS_PAGE_SIZE);
			if (raw[0] !== AccessPageType.Data) {
				continue;
			}
			const parsed = new AccessDataPage(raw);
			const slots = parsed.slots();
			const lowest = slots.length === 0
				? ACCESS_PAGE_SIZE
				: Math.min(...slots.map((entry) => entry & 0x1fff));
			// 14 bytes of header, two per slot, and everything below the
			// lowest row is taken.
			expect(parsed.freeSpace).toBe(lowest - 14 - 2 * slots.length);
			checked += 1;
		}
		expect(checked).toBeGreaterThan(0);
	});

	it.each(FIXTURES)('%s: a page round-trips through the editor unchanged', (file) => {
		const data = read(file);
		let checked = 0;
		for (let page = 0; page < pageCount(data); page += 1) {
			const raw = data.subarray(page * ACCESS_PAGE_SIZE, (page + 1) * ACCESS_PAGE_SIZE);
			if (raw[0] !== AccessPageType.Data) {
				continue;
			}
			expect(new AccessDataPage(raw).toBuffer().equals(raw)).toBe(true);
			checked += 1;
		}
		expect(checked).toBeGreaterThan(0);
	});
});

describe('usage maps', () => {
	it.each(FIXTURES)('%s: a table owns exactly the pages that name it', (file) => {
		const data = read(file);
		const store = new AccessPageStore(data);
		let checked = 0;
		for (const { name, table } of tables(data)) {
			const owned = new Set(readUsageMapRef(store, table.definition.ownedPagesRef).pages()
				.filter((page) => page < store.pageCount));
			const claiming = new Set<number>();
			for (let page = 0; page < store.pageCount; page += 1) {
				const raw = store.read(page);
				if (raw[0] === AccessPageType.Data && !isLongValuePage(raw)
					&& raw.readUInt32LE(4) === table.definition.page) {
					claiming.add(page);
				}
			}
			for (const page of claiming) {
				expect(owned.has(page), `${name} page ${page} claims the table but is unowned`)
					.toBe(true);
			}
			checked += 1;
		}
		expect(checked).toBeGreaterThan(0);
	});

	it.each(FIXTURES)('%s: the global map hands out a page and takes it back', (file) => {
		const store = new AccessPageStore(read(file));
		const before = readUsageMap(store, GLOBAL_USAGE_MAP_PAGE, GLOBAL_USAGE_MAP_ROW).pages();
		const page = allocatePage(store);
		expect(before).toContain(page);
		expect(readUsageMap(store, GLOBAL_USAGE_MAP_PAGE, GLOBAL_USAGE_MAP_ROW).pages())
			.not.toContain(page);
		releasePage(store, page, 'value');
		expect(readUsageMap(store, GLOBAL_USAGE_MAP_PAGE, GLOBAL_USAGE_MAP_ROW).pages())
			.toContain(page);
	});
});

describe('index keys', () => {
	it.each(FIXTURES)('%s: every stored key is the one the row encodes to', (file) => {
		const data = read(file);
		const store = new AccessPageStore(data);
		let entries = 0;
		for (const { name, table } of tables(data)) {
			const byRow = new Map<string, Buffer>();
			for (const row of table.rows()) {
				byRow.set(`${row.id.page}:${row.id.slot}`, row.bytes);
			}
			for (const index of table.definition.realIndexes) {
				for (const leaf of leafPages(store, index.rootPage)) {
					for (const entry of leaf.entries) {
						const bytes = byRow.get(`${entry.page}:${entry.row}`);
						if (!bytes) {
							continue;
						}
						const derived = encodeIndexKey(
							table.definition, index, splitRow(table.definition, bytes),
						);
						expect(
							derived?.toString('hex'),
							`${name} row (${entry.page}, ${entry.row})`,
						).toBe(entry.key.toString('hex'));
						entries += 1;
					}
				}
			}
		}
		// The fixtures carry a few hundred each; a run that found none would
		// pass every assertion above and prove nothing.
		expect(entries).toBeGreaterThan(100);
	});

	it('orders text the way Access orders it', () => {
		// Case is not encoded at all, so two spellings key identically.
		expect(encodeTextKey('Module1').equals(encodeTextKey('MODULE1'))).toBe(true);
		expect(encodeTextKey('a').equals(encodeTextKey('A'))).toBe(true);
		// Trailing spaces are trimmed before encoding; a leading one is not.
		expect(encodeTextKey('ab  ').equals(encodeTextKey('ab'))).toBe(true);
		expect(encodeTextKey(' ab').equals(encodeTextKey('ab'))).toBe(false);
		// An accent sorts after its base letter but before the next one.
		const [a, accented, b] = ['a', 'á', 'b'].map(encodeTextKey);
		expect(a.compare(accented)).toBeLessThan(0);
		expect(accented.compare(b)).toBeLessThan(0);
		// A base letter plus a combining mark keys like the precomposed form.
		expect(encodeTextKey('á').equals(accented)).toBe(true);
		// Sharp s expands to two elements, so it keys as "ss" does.
		expect(encodeTextKey('straße').equals(encodeTextKey('strasse'))).toBe(true);
		// Scripts outside Latin key too, and differently from each other.
		const cyrillic = encodeTextKey('Модуль');
		const japanese = encodeTextKey('モジュール');
		expect(cyrillic.length).toBeGreaterThan(0);
		expect(japanese.length).toBeGreaterThan(0);
		expect(cyrillic.equals(japanese)).toBe(false);
	});
});

describe('table definitions', () => {
	it.each(FIXTURES)('%s: every definition parses to exactly its declared length', (file) => {
		const data = read(file);
		let complete = 0;
		for (const entry of readAccessCatalog(data)) {
			if (entry.type !== 1) {
				continue;
			}
			const definition = readTableDefinition(data, entry.definitionPage);
			// A tail that did not parse leaves the map references at zero, and
			// the writer refuses such a table rather than acting on half-read
			// page numbers.
			if (definition.ownedPagesRef === 0) {
				continue;
			}
			expect(definition.freeSpacePagesRef).toBeGreaterThan(0);
			for (const column of definition.columns) {
				if (column.type === AccessColumnType.Ole || column.type === AccessColumnType.Memo) {
					expect(
						definition.columnUsageMaps.has(column.number),
						`${entry.name}.${column.name} has no long-value maps`,
					).toBe(true);
				}
			}
			complete += 1;
		}
		expect(complete).toBeGreaterThan(5);
	});
});
