import {
	ACCESS_PAGE_SIZE,
	AccessFormatError,
	AccessPageType,
	MSYS_OBJECTS_PAGE,
	isAccessFile,
	pageCount,
	pageType,
	readDataPageRows,
	readLongValue,
	readTableDefinition,
	type AccessLongValueRef,
	type AccessRow,
	type AccessTableDefinition,
} from './accessFormat';

/**
 * `MSysAccessStorage`, the fake structured storage an Access database keeps its
 * VBA project in.
 *
 * Each row is a folder or a stream: `Id`, `ParentId`, `Name`, `Type` (1 folder,
 * 2 stream) and `Lv`, the long value holding a stream's bytes. Under
 * `MSysAccessStorage_ROOT` sit `VBA`, `Modules`, `Forms` and the rest; the
 * project is at `VBA/VBAProject/VBA`, where the MS-OVBA `dir` stream, `PROJECT`
 * and one stream per module live - the module streams under names like
 * `SJIBUJDNZVEMZDQTVFSNHHEBBVTF`, which Access generates and which mean
 * nothing.
 *
 * This is the structural read for
 * github.com/WilliamSmithEdward/xlide_vscode/issues/65: rows found through the
 * catalog and the table's own definition, rather than by decompressing
 * candidate bytes to see what they turn out to be.
 */

export interface AccessStorageEntry {
	id: number;
	parentId: number;
	name: string;
	/** 1 for a folder, 2 for a stream. */
	type: number;
	/** Where the row lives, which is what a writer needs. */
	page: number;
	slot: number;
	/** A stream's bytes, read from its long value. */
	bytes?: Buffer;
	children: AccessStorageEntry[];
}

/** The catalog's own row for a table: its name and where its definition is. */
export interface AccessTable {
	name: string;
	/** A table's `Id` in the catalog IS the page its definition starts on. */
	definitionPage: number;
	type: number;
}

/** Every table the catalog names. */
export function readAccessCatalog(data: Buffer): AccessTable[] {
	if (!isAccessFile(data)) {
		throw new AccessFormatError('Not a Jet 4 or ACE database.');
	}
	const definition = readTableDefinition(data, MSYS_OBJECTS_PAGE);
	const out: AccessTable[] = [];
	for (const row of readTableRows(data, definition)) {
		const name = row.values.get('Name');
		const id = row.values.get('Id');
		const type = row.values.get('Type');
		if (typeof name !== 'string' || typeof id !== 'number' || typeof type !== 'number') {
			continue;
		}
		out.push({ name, definitionPage: id, type });
	}
	return out;
}

/** Every row of a table, gathered from the data pages its definition owns. */
export function readTableRows(data: Buffer, definition: AccessTableDefinition): AccessRow[] {
	const out: AccessRow[] = [];
	// The definition's usage map names its pages; walking every page and
	// asking which table owns it reaches the same set without decoding the
	// map, and a database this is used on is a few hundred pages.
	for (let page = 0; page < pageCount(data); page += 1) {
		if (pageType(data, page) !== AccessPageType.Data) {
			continue;
		}
		if (data.readUInt32LE(page * ACCESS_PAGE_SIZE + 4) !== definition.page) {
			continue;
		}
		out.push(...readDataPageRows(data, page, definition));
	}
	return out;
}

/**
 * The storage tree, with every stream's bytes read. Returns undefined when the
 * database has no `MSysAccessStorage` at all, which is a database with no VBA
 * project rather than a broken one.
 */
export function readAccessStorage(data: Buffer): AccessStorageEntry[] | undefined {
	const table = readAccessCatalog(data)
		.find((entry) => entry.name.toLowerCase() === 'msysaccessstorage');
	if (!table) {
		return undefined;
	}
	const definition = readTableDefinition(data, table.definitionPage);
	const entries: AccessStorageEntry[] = [];
	for (const row of readTableRows(data, definition)) {
		const id = row.values.get('Id');
		const name = row.values.get('Name');
		if (typeof id !== 'number' || typeof name !== 'string') {
			continue;
		}
		const parentId = row.values.get('ParentId');
		const type = row.values.get('Type');
		const lv = row.values.get('Lv');
		const entry: AccessStorageEntry = {
			id,
			parentId: typeof parentId === 'number' ? parentId : -1,
			name,
			type: typeof type === 'number' ? type : 0,
			page: row.page,
			slot: row.slot,
			children: [],
		};
		if (lv && typeof lv === 'object' && 'kind' in lv && lv.kind === 'longValue') {
			entry.bytes = readLongValue(data, lv as AccessLongValueRef);
		}
		entries.push(entry);
	}

	const byId = new Map(entries.map((entry) => [entry.id, entry]));
	const roots: AccessStorageEntry[] = [];
	for (const entry of entries) {
		const parent = byId.get(entry.parentId);
		if (parent && parent !== entry) {
			parent.children.push(entry);
		} else {
			roots.push(entry);
		}
	}
	return roots;
}

/**
 * The VBA project's streams, keyed by the name the storage gives them: `dir`,
 * `PROJECT`, `_VBA_PROJECT` and one per module.
 *
 * Found by walking to `VBA/VBAProject/VBA` when the tree has it, and falling
 * back to every stream under any folder named `VBA` when it does not - an
 * Access database that has been through a version upgrade does not always keep
 * the same nesting.
 */
export function readAccessVbaStreams(data: Buffer): Map<string, Buffer> {
	const roots = readAccessStorage(data);
	const out = new Map<string, Buffer>();
	if (!roots) {
		return out;
	}
	const found = findVbaProjectFolder(roots);
	if (!found) {
		return out;
	}
	// `PROJECT` and `PROJECTwm` are siblings of the folder holding `dir`, not
	// children of it, so the level above comes too.
	for (const entry of [...(found.parent?.children ?? []), ...found.folder.children]) {
		if (entry.bytes && entry.type === STREAM) {
			out.set(entry.name, entry.bytes);
		}
	}
	return out;
}

/** A storage row's Type: 1 is a folder, 2 a stream. */
const STREAM = 2;

/** The folder whose streams are the project: it is the one holding `dir`. */
function findVbaProjectFolder(
	roots: readonly AccessStorageEntry[],
): { folder: AccessStorageEntry; parent?: AccessStorageEntry } | undefined {
	const stack: Array<{ entry: AccessStorageEntry; parent?: AccessStorageEntry }> =
		roots.map((entry) => ({ entry }));
	while (stack.length > 0) {
		const { entry, parent } = stack.pop()!;
		if (entry.children.some((child) => child.name === 'dir' && child.bytes)) {
			return parent ? { folder: entry, parent } : { folder: entry };
		}
		stack.push(...entry.children.map((child) => ({ entry: child, parent: entry })));
	}
	return undefined;
}
