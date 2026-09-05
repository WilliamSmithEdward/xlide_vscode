import {
	AccessColumnType,
	AccessFormatError,
	readTableDefinitionFrom,
	type AccessColumn,
	type AccessRealIndex,
	type AccessTableDefinition,
} from './accessFormat';
import {
	AccessDataPage,
	PAGE_DATA,
	ROW_DELETED,
	ROW_OVERFLOW,
	decodeRowPointer,
	encodeRowPointer,
} from './accessDataPage';
import { AccessPageStore, readUsageMapRef } from './accessPageStore';
import { addToMap, allocatePage, releasePage, removeFromMap } from './accessAlloc';
import { decodeLongValue, freeLongValue, writeLongValue, type AccessLongValueMaps } from './accessLongValue';
import { AccessBTree } from './accessIndex';
import { encodeIndexKey } from './accessIndexKey';
import { encodeAccessScalar, type AccessScalar } from './accessValue';
import { encodeRow, splitRow, type AccessRawRow } from './accessRow';

/** The row count, a u32 in the table definition's header. */
const TDEF_ROW_COUNT = 0x10;
/** The last AutoNumber handed out, which the next row follows. */
const TDEF_NEXT_AUTONUMBER = 0x14;
/** The last complex-type id handed out. */
const TDEF_LAST_COMPLEX_ID = 0x1c;
/** Column header byte 15, bit 2: the column numbers its own rows. */
const COLUMN_AUTONUMBER = 0x04;

/**
 * Rewriting one row of a table.
 *
 * Only long-value columns can be changed here, which is what writing a VBA
 * module or a form design into `MSysAccessStorage` needs. Jet cannot index an
 * OLE or Memo column, so such a write moves no index key and the B-trees are
 * left alone; the guard below states that rather than assuming it.
 *
 * A row that no longer fits its page moves to another page behind an overflow
 * pointer, and moves back when it fits again, as the engine does. Its home
 * slot does not change, so index entries - which name the home - stay right.
 *
 * Ported from pyOpenVBA's `Table.update_row` and `_place_row`.
 */

/** Where a row lives: the page and slot that are its home. */
export interface AccessRowId {
	page: number;
	slot: number;
}

export class AccessTable {
	readonly definition: AccessTableDefinition;

	/**
	 * The word the engine stamps a chained long value's definition and its
	 * first page with, one per session; any value does, if both match.
	 */
	readonly stamp: number;

	constructor(readonly store: AccessPageStore, definitionPage: number) {
		this.stamp = 0x00500000 | (store.pageCount & 0xffff);
		this.definition = readTableDefinitionFrom(store, definitionPage);
		if (this.definition.ownedPagesRef === 0) {
			throw new AccessFormatError(
				`The definition at page ${definitionPage} did not parse far enough to be written to.`,
			);
		}
	}

	/** Every live row of the table, by home slot, with its bytes. */
	rows(): Array<{ id: AccessRowId; bytes: Buffer }> {
		const out: Array<{ id: AccessRowId; bytes: Buffer }> = [];
		for (const page of readUsageMapRef(this.store, this.definition.ownedPagesRef).pages()) {
			if (page >= this.store.pageCount) {
				continue;
			}
			const raw = this.store.read(page);
			if (raw[0] !== PAGE_DATA || raw.readUInt32LE(4) !== this.definition.page) {
				continue;
			}
			const data = new AccessDataPage(raw);
			for (let slot = 0; slot < data.slotCount; slot += 1) {
				const bytes = this.fetchRow({ page, slot });
				if (bytes) {
					out.push({ id: { page, slot }, bytes });
				}
			}
		}
		return out;
	}

	/**
	 * The bytes of the row whose home is `id`, following an overflow pointer
	 * to wherever the row now lives, or undefined when the slot is dead.
	 */
	fetchRow(id: AccessRowId): Buffer | undefined {
		const home = new AccessDataPage(this.store.read(id.page));
		if (id.slot < 0 || id.slot >= home.slotCount) {
			throw new AccessFormatError(`Page ${id.page} has no slot ${id.slot}.`);
		}
		const bytes = home.row(id.slot);
		if (bytes === undefined) {
			return undefined;
		}
		if ((home.slots()[id.slot] & ROW_OVERFLOW) === 0) {
			return bytes;
		}
		const target = decodeRowPointer(bytes);
		const moved = new AccessDataPage(this.store.read(target.page)).row(target.row, true);
		if (moved === undefined) {
			throw new AccessFormatError(`The overflow row (${id.page}, ${id.slot}) points at nothing.`);
		}
		return moved;
	}

	/**
	 * Replace one long-value column of a row. `data` of zero length, or
	 * undefined, makes the column null, which is how the engine stores an
	 * empty long value.
	 */
	setLongValue(
		id: AccessRowId,
		columnName: string,
		data: Buffer | undefined,
		stamp: number = this.stamp,
	): void {
		const column = this.column(columnName);
		if (column.type !== AccessColumnType.Ole && column.type !== AccessColumnType.Memo) {
			throw new AccessFormatError(`Column ${column.name} is not a long-value column.`);
		}
		for (const index of this.definition.realIndexes) {
			if (index.columns.some((key) => key.number === column.number)) {
				throw new AccessFormatError(
					`Column ${column.name} is part of an index key, which this writer does not maintain.`,
				);
			}
		}
		const bytes = this.fetchRow(id);
		if (bytes === undefined) {
			throw new AccessFormatError(`The row (${id.page}, ${id.slot}) is not live.`);
		}
		const parts = splitRow(this.definition, bytes);
		const maps = this.longValueMaps(column);

		// The engine stores the new value first and gives the old storage back
		// afterwards, so a value never lands on a page the old one still holds.
		const values = new Map(parts.values);
		if (data && data.length > 0) {
			values.set(column.number, writeLongValue(this.store, maps, data, stamp));
		} else {
			values.delete(column.number);
		}
		const old = parts.values.get(column.number);
		if (old) {
			freeLongValue(this.store, maps, decodeLongValue(old));
		}

		const booleans = new Set(
			this.definition.columns
				.filter((c) => c.type === AccessColumnType.Boolean && parts.present.get(c.number))
				.map((c) => c.number),
		);
		this.placeRow(id, encodeRow(this.definition, values, booleans, parts));
	}

	/**
	 * Add a row from named values: a column absent from the map, or mapped to
	 * null, is null. Returns the row's home slot.
	 */
	insertNamedRow(named: Map<string, AccessScalar>): AccessRowId {
		const values = new Map(named);
		this.assignAutoNumbers(values);
		const encoded = this.encodeValues(values);
		return this.insertRow(encoded.values, encoded.booleans);
	}

	/**
	 * Give the row its AutoNumber columns and move the counters the definition
	 * keeps. The counter holds the last number handed out, so a row that names
	 * none takes the one after it; a row that names its own keeps that number
	 * and the counter follows it, to the value written rather than to the
	 * larger of the two. A complex column is flagged AutoNumber too, and its id
	 * comes from its own counter instead.
	 */
	private assignAutoNumbers(values: Map<string, AccessScalar>): void {
		for (const column of this.definition.columns) {
			if (column.type === AccessColumnType.Complex
				|| (column.flags & COLUMN_AUTONUMBER) === 0) {
				continue;
			}
			const given = values.get(column.name);
			if (given === undefined || given === null) {
				const next = (this.definition.nextAutoNumber + 1) >>> 0;
				values.set(column.name, next >= 0x80000000 ? next - 0x100000000 : next);
				this.definition.nextAutoNumber = next;
			} else if (typeof given === 'number' && Number.isInteger(given)) {
				this.definition.nextAutoNumber = given >>> 0;
			} else {
				continue;
			}
			const counter = Buffer.alloc(4);
			counter.writeUInt32LE(this.definition.nextAutoNumber, 0);
			this.patchDefinition(TDEF_NEXT_AUTONUMBER, counter);
		}
		// Every complex column in a row shares one id, handed out from its own
		// counter. A row with no elements still takes one, and an id is not
		// reused after a delete.
		const complex = this.definition.columns.filter((c) => c.type === AccessColumnType.Complex);
		if (complex.length > 0 && complex.every((c) => {
			const given = values.get(c.name);
			return given === undefined || given === null;
		})) {
			this.definition.lastComplexId += 1;
			for (const column of complex) {
				values.set(column.name, this.definition.lastComplexId);
			}
			const counter = Buffer.alloc(4);
			counter.writeUInt32LE(this.definition.lastComplexId, 0);
			this.patchDefinition(TDEF_LAST_COMPLEX_ID, counter);
		}
	}

	/** Change named columns of a row, keeping the bytes of the rest. */
	updateNamedRow(id: AccessRowId, named: Map<string, AccessScalar>): void {
		const bytes = this.fetchRow(id);
		if (bytes === undefined) {
			throw new AccessFormatError(`The row (${id.page}, ${id.slot}) is not live.`);
		}
		const parts = splitRow(this.definition, bytes);
		const values = new Map(parts.values);
		const booleans = new Set(
			this.definition.columns
				.filter((c) => c.type === AccessColumnType.Boolean && parts.present.get(c.number))
				.map((c) => c.number),
		);
		const changed = new Set<number>();
		for (const [name, value] of named) {
			const column = this.column(name);
			changed.add(column.number);
			if (column.type === AccessColumnType.Boolean) {
				if (value) { booleans.add(column.number); } else { booleans.delete(column.number); }
				continue;
			}
			if (value === null || value === undefined) {
				if (this.isLongValue(column)) {
					const old = parts.values.get(column.number);
					if (old) {
						freeLongValue(this.store, this.longValueMaps(column), decodeLongValue(old));
					}
				}
				values.delete(column.number);
				continue;
			}
			const encoded = encodeAccessScalar(column, value);
			if (!this.isLongValue(column)) {
				values.set(column.number, encoded);
				continue;
			}
			// The engine stores the new value first and gives the old storage
			// back afterwards, so a value never lands on a page the old one
			// still holds.
			values.set(column.number, this.storeLongValue(column, encoded));
			const old = parts.values.get(column.number);
			if (old) {
				freeLongValue(this.store, this.longValueMaps(column), decodeLongValue(old));
			}
		}
		// A row written through an index costs the index one of the rows it
		// counts, even when the value does not change.
		const row = encodeRow(this.definition, values, booleans, parts);
		const after = splitRow(this.definition, row);
		this.definition.realIndexes.forEach((index, position) => {
			const tree = this.btree(index);
			const entry = tree.entryOf(id.page, id.slot);
			const next = encodeIndexKey(this.definition, index, after);
			const touched = index.columns.some((key) => changed.has(key.number));
			if (entry && touched) {
				this.rowLeftIndex(position);
			}
			if (entry && next && entry.key.equals(next)) {
				return;
			}
			if (entry) {
				tree.delete(entry);
			}
			if (next) {
				// The distinct-key count grows on inserts only: an update that
				// moves a row to a new key leaves it alone.
				tree.insert(next, id.page, id.slot);
			}
		});
		this.placeRow(id, row);
	}

	private encodeValues(
		named: Map<string, AccessScalar>,
	): { values: Map<number, Buffer>; booleans: Set<number> } {
		const values = new Map<number, Buffer>();
		const booleans = new Set<number>();
		for (const [name, value] of named) {
			const column = this.column(name);
			if (column.type === AccessColumnType.Boolean) {
				if (value) {
					booleans.add(column.number);
				}
				continue;
			}
			if (value === null || value === undefined) {
				continue;
			}
			const encoded = encodeAccessScalar(column, value);
			values.set(column.number, this.isLongValue(column)
				? this.storeLongValue(column, encoded)
				: encoded);
		}
		return { values, booleans };
	}

	private isLongValue(column: AccessColumn): boolean {
		return column.type === AccessColumnType.Ole || column.type === AccessColumnType.Memo;
	}

	/**
	 * Put a long value where it belongs and give back the row's twelve-byte
	 * definition. An empty one is stored as null, which is what the engine
	 * does, so the caller drops the column instead.
	 */
	private storeLongValue(column: AccessColumn, data: Buffer): Buffer {
		if (data.length === 0) {
			throw new AccessFormatError(
				`Column ${column.name}: an empty long value is stored as null; pass null.`,
			);
		}
		return writeLongValue(this.store, this.longValueMaps(column), data, this.stamp);
	}

	/**
	 * Add a row, with each column's bytes by column number - absent for null -
	 * and the Boolean columns that are true. Returns the row's home slot.
	 */
	insertRow(values: Map<number, Buffer>, booleans: Set<number> = new Set()): AccessRowId {
		const store = this.store;
		const row = encodeRow(this.definition, values, booleans);
		const page = this.pageWithRoom(row.length);
		const data = new AccessDataPage(store.read(page));
		const slot = data.addRow(row);
		store.write(page, data.toBuffer());
		const parts = splitRow(this.definition, row);
		for (const index of this.definition.realIndexes) {
			const key = encodeIndexKey(this.definition, index, parts);
			if (key === undefined) {
				continue;
			}
			if (this.btree(index).insert(key, page, slot)) {
				index.entryCount += 1;
				const distinct = Buffer.alloc(4);
				distinct.writeUInt32LE(index.entryCount, 0);
				this.patchDefinition(index.headerAt + 4, distinct);
			}
		}
		this.definition.numRows += 1;
		const count = Buffer.alloc(4);
		count.writeUInt32LE(this.definition.numRows, 0);
		this.patchDefinition(TDEF_ROW_COUNT, count);
		return { page, slot };
	}

	/** One index's tree, with the allocator its splits need. */
	private btree(index: AccessRealIndex): AccessBTree {
		return new AccessBTree(this.store, index.rootPage, this.definition.page, () => {
			const page = allocatePage(this.store);
			addToMap(this.store, index.usageMapRef, page);
			return page;
		});
	}

	/**
	 * Delete one row with its index entries and long values, settling the
	 * pages it leaves as the engine does. `retireEmptyPage` false leaves an
	 * emptied page alive and owned, which is what the engine does when a
	 * catalog row goes.
	 */
	deleteRow(id: AccessRowId, retireEmptyPage = true): void {
		const store = this.store;
		const bytes = this.fetchRow(id);
		if (bytes === undefined) {
			throw new AccessFormatError(`The row (${id.page}, ${id.slot}) is not live.`);
		}
		const parts = splitRow(this.definition, bytes);
		this.definition.realIndexes.forEach((index, position) => {
			const tree = this.btree(index);
			const entry = tree.entryOf(id.page, id.slot);
			// An index that ignores nulls holds no entry for a row whose key
			// is entirely null, and then there is nothing to take out of it.
			if (!entry) {
				return;
			}
			tree.delete(entry);
			this.rowLeftIndex(position);
		});
		for (const column of this.definition.columns) {
			if (column.type !== AccessColumnType.Ole && column.type !== AccessColumnType.Memo) {
				continue;
			}
			const raw = parts.values.get(column.number);
			if (raw) {
				freeLongValue(store, this.longValueMaps(column), decodeLongValue(raw));
			}
		}
		const moved = this.movedTo(id);
		if (moved !== undefined) {
			// The page that held the overflow copy is only written back: the
			// engine neither retires it when it empties nor re-lists it.
			const target = new AccessDataPage(store.read(moved.page));
			target.removeRow(moved.row, true);
			store.write(moved.page, target.toBuffer());
		}
		const page = new AccessDataPage(store.read(id.page));
		page.removeRow(id.slot);
		// A home slot that held only the four-byte pointer to a moved row is
		// written back without re-listing the page.
		this.rowRemoved(id.page, page, moved === undefined, retireEmptyPage);
		this.definition.numRows -= 1;
		const count = Buffer.alloc(4);
		count.writeUInt32LE(this.definition.numRows, 0);
		this.patchDefinition(TDEF_ROW_COUNT, count);
	}

	/**
	 * One row stopped being counted by an index. An index built over existing
	 * rows counts them in its header; each row taken away takes one off, the
	 * count stops at zero, and the distinct-key count is capped at what is
	 * left. An index made before its rows counts none and never changes.
	 */
	private rowLeftIndex(position: number): void {
		const index = this.definition.realIndexes[position];
		if (index.rowCount === 0) {
			return;
		}
		index.rowCount -= 1;
		index.entryCount = Math.min(index.entryCount, index.rowCount);
		const header = Buffer.alloc(8);
		header.writeUInt32LE(index.rowCount, 0);
		header.writeUInt32LE(index.entryCount, 4);
		this.patchDefinition(index.headerAt, header);
	}

	/** Overwrite bytes of the definition's header, which is on its first page. */
	private patchDefinition(offset: number, data: Buffer): void {
		const raw = Buffer.from(this.store.read(this.definition.page));
		data.copy(raw, offset);
		this.store.write(this.definition.page, raw);
	}

	private column(name: string): AccessColumn {
		const found = this.definition.columns.find(
			(c) => c.name.toLowerCase() === name.toLowerCase(),
		);
		if (!found) {
			throw new AccessFormatError(`The table has no column named ${name}.`);
		}
		return found;
	}

	private longValueMaps(column: AccessColumn): AccessLongValueMaps {
		const maps = this.definition.columnUsageMaps.get(column.number);
		if (!maps) {
			throw new AccessFormatError(`Column ${column.name} has no long-value usage maps.`);
		}
		return maps;
	}

	/** Where a row lives when its home slot holds an overflow pointer. */
	private movedTo(id: AccessRowId): { page: number; row: number } | undefined {
		const home = new AccessDataPage(this.store.read(id.page));
		const entry = home.slots()[id.slot];
		if ((entry & ROW_OVERFLOW) === 0 || (entry & ROW_DELETED) !== 0) {
			return undefined;
		}
		const pointer = home.row(id.slot);
		return pointer === undefined ? undefined : decodeRowPointer(pointer);
	}

	/**
	 * Store a row's new bytes: in its home slot when they fit, else on
	 * another page behind an overflow pointer.
	 */
	private placeRow(id: AccessRowId, row: Buffer): void {
		const store = this.store;
		const home = new AccessDataPage(store.read(id.page));
		const moved = this.movedTo(id);
		const [start, end] = home.span(id.slot);
		if (moved === undefined) {
			// In place when the growth fits the page's free space; otherwise
			// the row moves behind a pointer.
			if (row.length - (end - start) <= home.freeSpace) {
				home.replaceRow(id.slot, row);
				store.write(id.page, home.toBuffer());
				return;
			}
			// A page that could not take the growth leaves the free-space map.
			removeFromMap(store, this.definition.freeSpacePagesRef, id.page);
			const targetPage = this.pageWithRoom(row.length, id.page);
			const target = new AccessDataPage(store.read(targetPage));
			const targetSlot = target.addRow(row, ROW_DELETED);
			store.write(targetPage, target.toBuffer());
			const again = new AccessDataPage(store.read(id.page));
			again.replaceRow(id.slot, encodeRowPointer(targetPage, targetSlot), ROW_OVERFLOW);
			store.write(id.page, again.toBuffer());
			return;
		}
		if (row.length - (end - start) <= home.freeSpace) {
			// It fits at home again, the pointer's own bytes counted: bring it
			// back and drop the copy.
			home.replaceRow(id.slot, row, 0);
			store.write(id.page, home.toBuffer());
			const target = new AccessDataPage(store.read(moved.page));
			target.removeRow(moved.row, true);
			this.rowRemoved(moved.page, target);
			return;
		}
		const target = new AccessDataPage(store.read(moved.page));
		const [targetStart, targetEnd] = target.span(moved.row);
		if (row.length - (targetEnd - targetStart) <= target.freeSpace) {
			target.replaceRow(moved.row, row);
			store.write(moved.page, target.toBuffer());
			return;
		}
		target.removeRow(moved.row, true);
		store.write(moved.page, target.toBuffer());
		removeFromMap(store, this.definition.freeSpacePagesRef, moved.page);
		const landingPage = this.pageWithRoom(row.length, id.page);
		const landing = new AccessDataPage(store.read(landingPage));
		const landingSlot = landing.addRow(row, ROW_DELETED);
		store.write(landingPage, landing.toBuffer());
		const again = new AccessDataPage(store.read(id.page));
		again.replaceRow(id.slot, encodeRowPointer(landingPage, landingSlot), ROW_OVERFLOW);
		store.write(id.page, again.toBuffer());
	}

	/**
	 * Write back a page that just lost a row, the way the engine settles it:
	 * a page with rows left rejoins the free-space map; an emptied page is
	 * retired, released and dropped from both maps, unless it is the table's
	 * first data page, which stays.
	 */
	private rowRemoved(
		pageNumber: number,
		page: AccessDataPage,
		settle = true,
		retire = true,
	): void {
		const store = this.store;
		if (!settle) {
			// The row was only a pointer: the page is written back and nothing
			// else moves.
			store.write(pageNumber, page.toBuffer());
			return;
		}
		const owned = readUsageMapRef(store, this.definition.ownedPagesRef).pages();
		const first = owned.length > 0 ? Math.min(...owned) : pageNumber;
		if (retire && page.liveRows === 0 && pageNumber !== first) {
			page.retire();
			store.write(pageNumber, page.toBuffer());
			releasePage(store, pageNumber);
			removeFromMap(store, this.definition.ownedPagesRef, pageNumber);
			removeFromMap(store, this.definition.freeSpacePagesRef, pageNumber);
			return;
		}
		store.write(pageNumber, page.toBuffer());
		addToMap(store, this.definition.freeSpacePagesRef, pageNumber);
	}

	/**
	 * A data page that can take the row, the way the engine picks one: the
	 * free-space map's pages in order, dropping any that cannot hold it, else
	 * a fresh page registered with both maps.
	 */
	private pageWithRoom(rowLength: number, exclude?: number): number {
		const store = this.store;
		const d = this.definition;
		if (!AccessDataPage.empty(d.page).fits(rowLength)) {
			throw new AccessFormatError(`A row of ${rowLength} bytes cannot fit a page.`);
		}
		for (const candidate of readUsageMapRef(store, d.freeSpacePagesRef).pages()) {
			if (candidate >= store.pageCount || candidate === exclude) {
				continue;
			}
			const raw = store.read(candidate);
			if (raw[0] !== PAGE_DATA || raw.readUInt32LE(4) !== d.page) {
				continue;
			}
			if (new AccessDataPage(raw).fits(rowLength)) {
				return candidate;
			}
			removeFromMap(store, d.freeSpacePagesRef, candidate);
		}
		const page = allocatePage(store);
		store.write(page, AccessDataPage.empty(d.page).toBuffer());
		addToMap(store, d.ownedPagesRef, page);
		addToMap(store, d.freeSpacePagesRef, page);
		return page;
	}
}

export type { AccessRawRow };
