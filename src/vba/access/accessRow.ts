import { AccessColumnType, AccessFormatError, type AccessTableDefinition } from './accessFormat';

/**
 * Splitting a row into its columns' bytes and putting one back together.
 *
 * A row is a column count, the fixed-length block, the variable-length data,
 * then - from the end backwards - a null bitmap, the variable-column count and
 * that many offsets plus one, stored in reverse. A bit set in the bitmap means
 * the column HAS a value, and a Boolean column lives only there.
 *
 * Ported from pyOpenVBA's `_rows.py`, kept to what a rewrite needs: the bytes
 * of every column, so an untouched one keeps exactly what it had.
 */

/** A row split into per-column byte slices, with the row itself kept. */
export interface AccessRawRow {
	/** The column count the row was written with, which may not be the table's. */
	columnCount: number;
	/** Each column's bytes by column number; absent when the column is null. */
	values: Map<number, Buffer>;
	/** Whether each column's bitmap bit is set. */
	present: Map<number, boolean>;
	raw: Buffer;
	/** The variable-column boundaries, in column order, with the end. */
	variableOffsets: number[];
}

export function splitRow(definition: AccessTableDefinition, row: Buffer): AccessRawRow {
	if (row.length < 2) {
		throw new AccessFormatError('A row is shorter than its column-count word.');
	}
	const columnCount = row.readUInt16LE(0);
	const maskLength = Math.ceil(columnCount / 8);
	if (row.length < 2 + maskLength) {
		throw new AccessFormatError('A row is shorter than its null mask.');
	}
	const mask = row.subarray(row.length - maskLength);
	const variableOffsets: number[] = [];
	if (definition.numVariableColumns > 0) {
		const countAt = row.length - maskLength - 2;
		if (countAt < 2) {
			throw new AccessFormatError('A row is shorter than its variable-column count.');
		}
		const variableCount = row.readUInt16LE(countAt);
		const tableAt = countAt - 2 * (variableCount + 1);
		if (tableAt < 2) {
			throw new AccessFormatError('A row is shorter than its variable-column offsets.');
		}
		// Memory order is [end, start(n-1), ..., start(0)].
		for (let i = variableCount; i >= 0; i -= 1) {
			variableOffsets.push(row.readUInt16LE(tableAt + i * 2));
		}
	}

	const values = new Map<number, Buffer>();
	const present = new Map<number, boolean>();
	for (const column of definition.columns) {
		const number = column.number;
		if (number >= columnCount) {
			// Added after this row was written: no data and no mask bit.
			present.set(number, false);
			continue;
		}
		const has = (mask[number >> 3] & (1 << (number % 8))) !== 0;
		present.set(number, has);
		if (!has) {
			continue;
		}
		if (column.type === AccessColumnType.Boolean) {
			values.set(number, Buffer.alloc(0));
			continue;
		}
		if (column.fixed) {
			const start = 2 + column.fixedOffset;
			values.set(number, row.subarray(start, start + column.length));
			continue;
		}
		if (column.variableIndex + 1 >= variableOffsets.length) {
			throw new AccessFormatError(
				`Column ${column.name} has variable index ${column.variableIndex} but the row `
				+ `holds ${Math.max(variableOffsets.length - 1, 0)} variable columns.`,
			);
		}
		const start = variableOffsets[column.variableIndex];
		const end = variableOffsets[column.variableIndex + 1];
		if (!(start >= 2 && start <= end && end <= row.length)) {
			throw new AccessFormatError(
				`Column ${column.name} spans ${start}..${end} in a ${row.length}-byte row.`,
			);
		}
		values.set(number, row.subarray(start, end));
	}
	return { columnCount, values, present, raw: Buffer.from(row), variableOffsets };
}

/**
 * Reassemble a row from per-column bytes, keyed by column number; a column
 * absent from `values` is null. `booleans` names the Boolean columns that are
 * true, since a Boolean has no bytes of its own.
 *
 * `template` is the row's previous version. Bytes it holds for columns the
 * definition no longer has - fixed slots, variable slots and mask bits - are
 * carried over, as the engine carries them.
 */
export function encodeRow(
	definition: AccessTableDefinition,
	values: Map<number, Buffer>,
	booleans: Set<number>,
	template?: AccessRawRow,
): Buffer {
	const columns = [...definition.columns].sort((a, b) => a.number - b.number);
	const columnCount = columns.reduce((most, c) => Math.max(most, c.number + 1), 0);
	const fixedSize = columns
		.filter((c) => c.fixed)
		.reduce((size, c) => Math.max(size, c.fixedOffset + c.length), 0);
	let fixedBlock = Buffer.alloc(fixedSize);
	const variable = new Map<number, Buffer>();
	const nullMask = Buffer.alloc(Math.ceil(columnCount / 8));
	const known = new Set(columns.map((c) => c.number));

	if (template && template.raw.length > 0) {
		const oldMaskLength = Math.ceil(template.columnCount / 8);
		const oldVariableCount = Math.max(template.variableOffsets.length - 1, 0);
		if (template.variableOffsets.length > 0 && oldVariableCount !== definition.numVariableColumns) {
			// The engine keeps the old body verbatim and appends a fresh
			// offset table behind it. Only a definition change gets here, and
			// this writer never makes one, so it refuses rather than guessing.
			throw new AccessFormatError(
				`The row was written with ${oldVariableCount} variable columns but the definition `
				+ `has ${definition.numVariableColumns}; rewriting a row across a definition change `
				+ 'is not supported.',
			);
		}
		// The fixed block is the old row's whole pre-variable region, which is
		// never shorter than the definition's fixed size and may carry junk an
		// earlier rewrite left; variable slots are carried by index.
		const oldFixedLength = template.variableOffsets.length > 0
			? template.variableOffsets[0] - 2
			: template.raw.length - 2 - oldMaskLength;
		fixedBlock = Buffer.alloc(Math.max(fixedSize, oldFixedLength));
		template.raw.subarray(2, 2 + fixedBlock.length).copy(fixedBlock);
		const taken = new Set(columns.filter((c) => !c.fixed).map((c) => c.variableIndex));
		for (let index = 0; index < oldVariableCount; index += 1) {
			if (!taken.has(index)) {
				variable.set(index, template.raw.subarray(
					template.variableOffsets[index], template.variableOffsets[index + 1],
				));
			}
		}
		const oldMask = template.raw.subarray(template.raw.length - oldMaskLength);
		for (let number = 0; number < Math.min(template.columnCount, columnCount); number += 1) {
			if (!known.has(number) && (oldMask[number >> 3] & (1 << (number % 8))) !== 0) {
				nullMask[number >> 3] |= 1 << (number % 8);
			}
		}
	}

	const mark = (number: number): void => {
		nullMask[number >> 3] |= 1 << (number % 8);
	};
	for (const column of columns) {
		if (column.type === AccessColumnType.Boolean) {
			if (booleans.has(column.number)) {
				mark(column.number);
			}
			continue;
		}
		const value = values.get(column.number);
		if (value === undefined) {
			continue;
		}
		mark(column.number);
		if (column.fixed) {
			if (value.length !== column.length) {
				throw new AccessFormatError(
					`Column ${column.name}: ${value.length} bytes for a ${column.length}-byte fixed column.`,
				);
			}
			value.copy(fixedBlock, column.fixedOffset);
		} else {
			variable.set(column.variableIndex, value);
		}
	}

	const parts: Buffer[] = [];
	const head = Buffer.alloc(2);
	head.writeUInt16LE(columnCount, 0);
	parts.push(head, fixedBlock);
	let at = 2 + fixedBlock.length;
	const variableCount = definition.numVariableColumns;
	if (variableCount > 0) {
		const offsets: number[] = [];
		for (let index = 0; index < variableCount; index += 1) {
			offsets.push(at);
			const bytes = variable.get(index);
			if (bytes && bytes.length > 0) {
				parts.push(bytes);
				at += bytes.length;
			}
		}
		offsets.push(at);
		if (at > 0x1fff) {
			throw new AccessFormatError('The row is too long for its offset table.');
		}
		const table = Buffer.alloc((offsets.length + 1) * 2);
		offsets.slice().reverse().forEach((offset, i) => table.writeUInt16LE(offset, i * 2));
		table.writeUInt16LE(variableCount, offsets.length * 2);
		parts.push(table);
	}
	parts.push(nullMask);
	return Buffer.concat(parts);
}
