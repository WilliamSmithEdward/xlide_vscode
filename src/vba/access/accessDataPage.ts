import { ACCESS_PAGE_SIZE, AccessFormatError } from './accessFormat';

/**
 * In-place editing of a data-shaped page: the slot table and the rows.
 *
 * Ported from pyOpenVBA's `_datapage.py`, whose rules were measured on pages
 * the ACE engine itself wrote:
 *
 * - Rows are laid down from the end of the page toward the slot table, in slot
 *   order, with no gaps: row k occupies [offset(k), offset(k-1)).
 * - Free space is exactly 4096 - 14 - 2 * slots - bytes below the lowest row.
 * - Deleting a row shifts every row below it up over the hole, fixes their
 *   offsets, and leaves the slot in place flagged 0xC000 with its offset at the
 *   boundary it now sits on. The bytes freed at the bottom are not cleared.
 * - Replacing a row shifts the rows below it by the size difference.
 * - Inserting appends a slot and places the row below the lowest one.
 *
 * Matching the engine byte for byte is the point: a page this writes has to be
 * one Access would have written, or the next thing Access does with it is
 * undefined.
 */

export const PAGE_DATA = 0x01;
/** A data-shaped page emptied of rows and given back. */
export const PAGE_RETIRED = 0x09;
export const LVAL_OWNER_TAG = 0x4c41564c;

const OFFSET_FREE_SPACE = 0x02;
const OFFSET_OWNER = 0x04;
const OFFSET_ROW_COUNT = 0x0c;
const OFFSET_ROW_TABLE = 0x0e;

/** Thirteen bits: a boundary can be the page end, which needs bit 12. */
const ROW_OFFSET_MASK = 0x1fff;
export const ROW_DELETED = 0x8000;
export const ROW_OVERFLOW = 0x4000;
/** Both bits together mark a slot that holds nothing at all. */
const DEAD_SLOT = ROW_DELETED | ROW_OVERFLOW;

const INITIAL_FREE_SPACE = ACCESS_PAGE_SIZE - OFFSET_ROW_TABLE;

/** `(row byte, three-byte page)`, how one LVAL chunk points at the next. */
export function encodeRowPointer(page: number, row: number): Buffer {
	if (row < 0 || row > 0xff || page < 0 || page > 0xffffff) {
		throw new AccessFormatError(`Row pointer (${page}, ${row}) is out of range.`);
	}
	const out = Buffer.alloc(4);
	out[0] = row;
	out.writeUIntLE(page, 1, 3);
	return out;
}

/** The `(page, row)` a four-byte reference names. */
export function decodeRowPointer(pointer: Buffer): { page: number; row: number } {
	if (pointer.length < 4) {
		throw new AccessFormatError('A row pointer is four bytes.');
	}
	return { page: pointer.readUIntLE(1, 3), row: pointer[0] };
}

/** Whether the page is a long-value page: a data page owned by the LVAL tag. */
export function isLongValuePage(page: Buffer): boolean {
	return page[0] === PAGE_DATA && page.readUInt32LE(OFFSET_OWNER) === LVAL_OWNER_TAG;
}

export class AccessDataPage {
	private readonly raw: Buffer;

	constructor(page: Buffer) {
		if (page.length !== ACCESS_PAGE_SIZE) {
			throw new AccessFormatError(`A page is ${ACCESS_PAGE_SIZE} bytes, got ${page.length}.`);
		}
		if (page[0] !== PAGE_DATA) {
			throw new AccessFormatError(`Page type 0x${page[0].toString(16)} is not a data page.`);
		}
		this.raw = Buffer.from(page);
	}

	/** A fresh, empty data page owned by the given table (or the LVAL tag). */
	static empty(owner: number): AccessDataPage {
		const raw = Buffer.alloc(ACCESS_PAGE_SIZE);
		raw[0] = PAGE_DATA;
		raw[1] = 0x01;
		raw.writeUInt16LE(INITIAL_FREE_SPACE, OFFSET_FREE_SPACE);
		raw.writeUInt32LE(owner, OFFSET_OWNER);
		return new AccessDataPage(raw);
	}

	get owner(): number {
		return this.raw.readUInt32LE(OFFSET_OWNER);
	}

	get freeSpace(): number {
		return this.raw.readUInt16LE(OFFSET_FREE_SPACE);
	}

	get slotCount(): number {
		return this.raw.readUInt16LE(OFFSET_ROW_COUNT);
	}

	slots(): number[] {
		const out: number[] = [];
		for (let i = 0; i < this.slotCount; i += 1) {
			out.push(this.raw.readUInt16LE(OFFSET_ROW_TABLE + i * 2));
		}
		return out;
	}

	/** `[start, end)` of a slot's row. */
	span(slot: number): [number, number] {
		const slots = this.slots();
		const start = slots[slot] & ROW_OFFSET_MASK;
		const end = slot === 0 ? ACCESS_PAGE_SIZE : (slots[slot - 1] & ROW_OFFSET_MASK);
		return [start, end];
	}

	private lowestOffset(): number {
		const slots = this.slots();
		if (slots.length === 0) {
			return ACCESS_PAGE_SIZE;
		}
		return Math.min(...slots.map((entry) => entry & ROW_OFFSET_MASK));
	}

	/** Whether a row of this length fits, its slot included. */
	fits(rowLength: number): boolean {
		return this.freeSpace >= rowLength + 2;
	}

	/**
	 * A slot's row, or undefined when the slot holds none. An overflow slot
	 * gives back its four-byte pointer; the row it names is read from its own
	 * page with `overflowTarget`, where the 0x8000 bit marks the moved row
	 * rather than a deletion.
	 */
	row(slot: number, overflowTarget = false): Buffer | undefined {
		const entry = this.slots()[slot];
		if ((entry & ROW_DELETED) !== 0 && !overflowTarget) {
			return undefined;
		}
		const [start, end] = this.span(slot);
		if ((entry & ROW_OVERFLOW) !== 0) {
			if (overflowTarget) {
				throw new AccessFormatError(`Slot ${slot} is itself an overflow pointer.`);
			}
			return this.raw.subarray(start, start + 4);
		}
		return this.raw.subarray(start, end);
	}

	/** Slots that still hold a row; a dead slot does not. */
	get liveRows(): number {
		return this.slots().filter((entry) => (entry & DEAD_SLOT) !== DEAD_SLOT).length;
	}

	/** Append a slot for `row` and return its number. */
	addRow(row: Buffer, flags = 0): number {
		if (!this.fits(row.length)) {
			throw new AccessFormatError(
				`A row of ${row.length} bytes does not fit; ${this.freeSpace} bytes are free.`,
			);
		}
		const slot = this.slotCount;
		const start = this.lowestOffset() - row.length;
		if (start < OFFSET_ROW_TABLE + 2 * (slot + 1)) {
			throw new AccessFormatError('The row would overlap the slot table.');
		}
		row.copy(this.raw, start);
		this.raw.writeUInt16LE(slot + 1, OFFSET_ROW_COUNT);
		this.setSlot(slot, flags | start);
		this.setFreeSpace(this.freeSpace - row.length - 2);
		return slot;
	}

	/**
	 * Overwrite a slot's row - a plain row, an overflow pointer, or a row
	 * moved here from another page - shifting the rows below by the size
	 * change. `flags` replaces the slot's flag bits when given.
	 */
	replaceRow(slot: number, row: Buffer, flags?: number): void {
		const entry = this.slots()[slot];
		if ((entry & DEAD_SLOT) === DEAD_SLOT) {
			throw new AccessFormatError(`Slot ${slot} is dead.`);
		}
		let [start, end] = this.span(slot);
		const delta = row.length - (end - start);
		if (delta > this.freeSpace) {
			throw new AccessFormatError(
				`The row grows by ${delta} bytes but only ${this.freeSpace} are free.`,
			);
		}
		if (delta !== 0) {
			this.shiftBelow(start, -delta);
			start -= delta;
		}
		this.setSlot(slot, (flags === undefined ? entry & ~ROW_OFFSET_MASK : flags) | start);
		row.copy(this.raw, start);
		this.setFreeSpace(this.freeSpace - delta);
	}

	/**
	 * Delete a row the way the engine does: close the hole, leave a dead
	 * slot. Works for a plain row and an overflow pointer; the moved row a
	 * pointer refers to carries 0x8000, which is not a deletion there, so
	 * removing one takes `overflowTarget`.
	 */
	removeRow(slot: number, overflowTarget = false): void {
		const entry = this.slots()[slot];
		if ((entry & ROW_DELETED) !== 0 && !overflowTarget) {
			throw new AccessFormatError(`Slot ${slot} is already dead.`);
		}
		if (overflowTarget && (entry & DEAD_SLOT) !== ROW_DELETED) {
			throw new AccessFormatError(`Slot ${slot} does not hold a moved row.`);
		}
		const [start, end] = this.span(slot);
		const length = end - start;
		this.shiftBelow(start, length);
		this.setSlot(slot, DEAD_SLOT | end);
		this.setFreeSpace(this.freeSpace + length);
	}

	/**
	 * What the engine does to a page whose last row went: type 0x09, every slot
	 * dead at the page end, the free word counting only the header and the slot
	 * table, and the rows and owner left where they are.
	 */
	retire(): void {
		this.raw[0] = PAGE_RETIRED;
		for (let slot = 0; slot < this.slotCount; slot += 1) {
			this.setSlot(slot, DEAD_SLOT | ACCESS_PAGE_SIZE);
		}
		this.setFreeSpace(INITIAL_FREE_SPACE - 2 * this.slotCount);
	}

	toBuffer(): Buffer {
		return Buffer.from(this.raw);
	}

	private setSlot(slot: number, value: number): void {
		this.raw.writeUInt16LE(value & 0xffff, OFFSET_ROW_TABLE + slot * 2);
	}

	private setFreeSpace(value: number): void {
		if (value < 0 || value > INITIAL_FREE_SPACE) {
			throw new AccessFormatError(`Free space of ${value} is impossible.`);
		}
		this.raw.writeUInt16LE(value, OFFSET_FREE_SPACE);
	}

	/**
	 * Move every row starting below `boundary` by `delta` (positive moves
	 * toward the page end) and fix its slot offset. A dead slot sitting exactly
	 * at the boundary belongs to the block below and moves with it, even when
	 * that block is empty.
	 */
	private shiftBelow(boundary: number, delta: number): void {
		const lowest = this.lowestOffset();
		if (lowest < boundary) {
			const block = Buffer.from(this.raw.subarray(lowest, boundary));
			block.copy(this.raw, lowest + delta);
		}
		this.slots().forEach((entry, slot) => {
			const offset = entry & ROW_OFFSET_MASK;
			const deadAtBoundary = (entry & DEAD_SLOT) === DEAD_SLOT && offset === boundary;
			if (offset < boundary || deadAtBoundary) {
				this.setSlot(slot, (entry & ~ROW_OFFSET_MASK) | (offset + delta));
			}
		});
	}
}
