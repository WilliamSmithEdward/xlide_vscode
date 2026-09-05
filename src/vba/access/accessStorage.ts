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
import { parseAccessDesign, type AccessDesign } from './accessDesign';

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
export const ACCESS_STORAGE_FOLDER = 1;
export const ACCESS_STORAGE_STREAM = 2;
const STREAM = ACCESS_STORAGE_STREAM;

/** Every entry of the tree, in no particular order. */
export function flattenAccessStorage(
	roots: readonly AccessStorageEntry[],
): AccessStorageEntry[] {
	const out: AccessStorageEntry[] = [];
	const stack = [...roots];
	while (stack.length > 0) {
		const entry = stack.pop()!;
		out.push(entry);
		stack.push(...entry.children);
	}
	return out;
}

/** The folder whose streams are the VBA project, and the one above it. */
export function accessVbaProjectFolder(
	roots: readonly AccessStorageEntry[],
): { folder: AccessStorageEntry; parent?: AccessStorageEntry } | undefined {
	return findVbaProjectFolder(roots);
}

/** A form or a report, with the design the file describes it by. */
export interface AccessDesignEntry {
	name: string;
	kind: 'form' | 'report';
	/** The numbered folder under `Forms` or `Reports`. */
	ordinal: string;
	design: AccessDesign;
	/** The design's other streams, kept so the object can be rebuilt. */
	typeInfo?: Buffer;
	propData?: Buffer;
}

/**
 * Every form and report the database describes
 * (github.com/WilliamSmithEdward/xlide_vscode/issues/67).
 *
 * A design sits in a `Blob` under a numbered folder under `Forms` or
 * `Reports`, the way a module's streams sit under `Modules`. The blob parses
 * into the objects the designer shows: the sections, then the controls.
 */
export function readAccessDesigns(data: Buffer): AccessDesignEntry[] {
	const roots = readAccessStorage(data);
	if (!roots) {
		return [];
	}
	const out: AccessDesignEntry[] = [];
	for (const [folder, kind] of [['Forms', 'form'], ['Reports', 'report']] as const) {
		const container = findChild(roots, folder);
		for (const ordinal of container?.children ?? []) {
			if (!/^\d+$/.test(ordinal.name)) {
				continue;
			}
			const blob = ordinal.children.find((entry) => entry.name === 'Blob')?.bytes;
			if (!blob || blob.length === 0) {
				continue;
			}
			out.push({
				name: designName(container!, ordinal.name) ?? ordinal.name,
				kind,
				ordinal: ordinal.name,
				design: parseAccessDesign(blob),
				...streamOf(ordinal, 'TypeInfo'),
				...streamOf(ordinal, 'PropData'),
			});
		}
	}
	return out;
}

function streamOf(
	ordinal: AccessStorageEntry,
	name: 'TypeInfo' | 'PropData',
): Record<string, Buffer> {
	const bytes = ordinal.children.find((entry) => entry.name === name)?.bytes;
	return bytes ? { [name === 'TypeInfo' ? 'typeInfo' : 'propData']: bytes } : {};
}

/**
 * The object's name, from the container's `\x03DirData`. It lists the names in
 * ordinal order, each UTF-16LE and framed, so the name is read out of the run
 * that follows the ordinal's own entry.
 */
function designName(container: AccessStorageEntry, ordinal: string): string | undefined {
	const dirData = container.children.find((entry) => entry.name.endsWith('DirData'))?.bytes;
	if (!dirData) {
		return undefined;
	}
	const names = [...dirData.toString('utf16le').matchAll(/[\p{L}_][\p{L}\p{N}_ ]*/gu)]
		.map((match) => match[0].trim())
		.filter((name) => name.length > 0);
	return names[Number(ordinal)];
}

function findChild(roots: readonly AccessStorageEntry[], name: string): AccessStorageEntry | undefined {
	const stack = [...roots];
	while (stack.length > 0) {
		const entry = stack.pop()!;
		if (entry.name === name) {
			return entry;
		}
		stack.push(...entry.children);
	}
	return undefined;
}

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
