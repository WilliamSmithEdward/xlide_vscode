// Binary PowerPoint (.ppt) VBA container: locate, extract, and write back
// the embedded VBA project storage.
//
// A .ppt is a CFB whose 'PowerPoint Document' stream is a flat sequence of
// records addressed through a persist model: the 'Current User' stream's
// CurrentUserAtom points at the newest UserEditAtom, each UserEditAtom
// points at a PersistDirectoryAtom (and at the previous edit), and the
// directories map persist object ids to absolute byte offsets in the
// stream. The VBA project is an ExOleObjStg record (0x1011) - a zlib
// deflate of a CFB, truncated without a proper stream end - whose persist
// id the DocumentContainer's VbaInfoAtom names.
//
// Writing back replaces that record in place. Every structure that carries
// absolute offsets - the persist directories, the user-edit chain, and the
// CurrentUserAtom - is shifted by the size delta; nothing else in the file
// addresses by offset (that is what the persist model is for).

import * as zlib from 'zlib';
import { Cfb } from './cfb';

export class PptContainerError extends Error {}

const CFB_MAGIC = Buffer.from([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]);

const RT_DOCUMENT_CONTAINER = 0x03e8;
const RT_VBA_INFO = 0x03ff;
const RT_VBA_INFO_ATOM = 0x0400;
const RT_USER_EDIT_ATOM = 0x0ff5;
const RT_PERSIST_DIRECTORY_ATOM = 0x1772;
const RT_EX_OLE_OBJ_STG = 0x1011;
const RT_CURRENT_USER_ATOM = 0x0ff6;

interface RecordHeader {
	verInstance: number;
	recType: number;
	recLen: number;
	/** Offset of the 8-byte header itself. */
	start: number;
	/** Offset just past the record's payload. */
	end: number;
}

function readHeader(stream: Buffer, offset: number): RecordHeader | undefined {
	if (offset < 0 || offset + 8 > stream.length) {
		return undefined;
	}
	const verInstance = stream.readUInt16LE(offset);
	const recType = stream.readUInt16LE(offset + 2);
	const recLen = stream.readUInt32LE(offset + 4);
	if (recLen > stream.length - (offset + 8)) {
		return undefined;
	}
	return { verInstance, recType, recLen, start: offset, end: offset + 8 + recLen };
}

function* topLevelRecords(stream: Buffer): Generator<RecordHeader> {
	let offset = 0;
	for (;;) {
		const header = readHeader(stream, offset);
		if (!header) {
			return;
		}
		yield header;
		offset = header.end;
	}
}

// --------------------------------------------------------- persist machinery

interface PersistState {
	/** persist id -> absolute offset of the referenced record, newest edit wins. */
	persistOffsets: Map<number, number>;
	/** Every UserEditAtom's header, newest first. */
	userEdits: RecordHeader[];
	/** Every PersistDirectoryAtom reachable from the edit chain. */
	persistDirectories: RecordHeader[];
}

/**
 * Walks the user-edit chain from CurrentUserAtom.offsetToCurrentEdit and
 * folds the persist directories, newest edit first, so the first directory
 * that names a persist id supplies its offset.
 */
function walkPersistChain(doc: Buffer, offsetToCurrentEdit: number): PersistState {
	const state: PersistState = { persistOffsets: new Map(), userEdits: [], persistDirectories: [] };
	const seen = new Set<number>();
	let editOffset = offsetToCurrentEdit;
	while (editOffset > 0 && !seen.has(editOffset)) {
		seen.add(editOffset);
		const edit = readHeader(doc, editOffset);
		if (!edit || edit.recType !== RT_USER_EDIT_ATOM) {
			break;
		}
		state.userEdits.push(edit);
		// UserEditAtom payload: lastSlideIdRef(4) version(4) offsetLastEdit(4)
		// offsetPersistDirectory(4) docPersistIdRef(4) persistIdSeed(4) ...
		const payload = edit.start + 8;
		const offsetLastEdit = doc.readUInt32LE(payload + 8);
		const offsetPersistDirectory = doc.readUInt32LE(payload + 12);
		const directory = readHeader(doc, offsetPersistDirectory);
		if (directory && directory.recType === RT_PERSIST_DIRECTORY_ATOM) {
			state.persistDirectories.push(directory);
			for (const [persistId, offset] of readPersistDirectory(doc, directory)) {
				if (!state.persistOffsets.has(persistId)) {
					state.persistOffsets.set(persistId, offset);
				}
			}
		}
		editOffset = offsetLastEdit;
	}
	return state;
}

/** Decodes a PersistDirectoryAtom: runs of (persistId:20, cPersist:12) + offsets. */
function* readPersistDirectory(doc: Buffer, directory: RecordHeader): Generator<[number, number]> {
	let cursor = directory.start + 8;
	const end = directory.end;
	while (cursor + 4 <= end) {
		const info = doc.readUInt32LE(cursor);
		cursor += 4;
		const firstId = info & 0xfffff;
		const count = info >>> 20;
		for (let i = 0; i < count && cursor + 4 <= end; i++) {
			yield [firstId + i, doc.readUInt32LE(cursor)];
			cursor += 4;
		}
	}
}

/** CurrentUserAtom.offsetToCurrentEdit lives at payload offset 8. */
function currentEditOffset(currentUser: Buffer): number | undefined {
	const header = readHeader(currentUser, 0);
	if (!header || header.recType !== RT_CURRENT_USER_ATOM) {
		return undefined;
	}
	// Payload: size(4) headerToken(4) offsetToCurrentEdit(4) ...
	return currentUser.readUInt32LE(header.start + 8 + 8);
}

/** The VBA project's persist id, from the DocumentContainer's VbaInfoAtom. */
function vbaPersistId(doc: Buffer, state: PersistState): number | undefined {
	for (const [, offset] of state.persistOffsets) {
		const record = readHeader(doc, offset);
		if (!record || record.recType !== RT_DOCUMENT_CONTAINER) {
			continue;
		}
		const vbaInfo = findDescendant(doc, record, RT_VBA_INFO);
		const vbaInfoAtom = vbaInfo && findDescendant(doc, vbaInfo, RT_VBA_INFO_ATOM);
		if (vbaInfoAtom && vbaInfoAtom.recLen >= 4) {
			return doc.readUInt32LE(vbaInfoAtom.start + 8);
		}
	}
	return undefined;
}

function findDescendant(doc: Buffer, container: RecordHeader, recType: number): RecordHeader | undefined {
	let cursor = container.start + 8;
	while (cursor + 8 <= container.end) {
		const child = readHeader(doc, cursor);
		if (!child || child.end > container.end) {
			return undefined;
		}
		if (child.recType === recType) {
			return child;
		}
		if ((child.verInstance & 0xf) === 0xf) {
			const nested = findDescendant(doc, child, recType);
			if (nested) {
				return nested;
			}
		}
		cursor = child.end;
	}
	return undefined;
}

// ------------------------------------------------------------------ locating

interface VbaStorageLocation {
	/** Header of the ExOleObjStg record holding the project. */
	record: RecordHeader;
	/** The embedded CFB, inflated when the record is compressed. */
	storage: Buffer;
}

function decodeStorageRecord(doc: Buffer, record: RecordHeader): Buffer | undefined {
	const body = doc.subarray(record.start + 8, record.end);
	const recInstance = record.verInstance >> 4;
	if (recInstance === 0) {
		return body.subarray(0, 8).equals(CFB_MAGIC) ? body : undefined;
	}
	if (body.length < 4) {
		return undefined;
	}
	try {
		// PowerPoint truncates the deflate stream without a proper stream
		// end; the sync-flush finish accepts exactly the declared bytes.
		const inflated = zlib.inflateSync(body.subarray(4), {
			finishFlush: zlib.constants.Z_SYNC_FLUSH,
		});
		return inflated.subarray(0, 8).equals(CFB_MAGIC) ? inflated : undefined;
	} catch {
		return undefined;
	}
}

function isVbaProjectStorage(storage: Buffer): boolean {
	try {
		const inner = Cfb.fromBytes(storage);
		return inner.hasStreamInStorage('VBA', 'dir') || inner.hasStream('dir');
	} catch {
		return false;
	}
}

/**
 * The current VBA project storage: resolved through the persist chain when
 * the file is well-formed, else by scanning every ExOleObjStg record and
 * keeping the last that parses as a VBA project (append-order means newest).
 */
function locateVbaStorage(cfb: Cfb): VbaStorageLocation | undefined {
	const doc = cfb.getStream('PowerPoint Document');
	try {
		const editOffset = currentEditOffset(cfb.getStream('Current User'));
		if (editOffset !== undefined) {
			const state = walkPersistChain(doc, editOffset);
			const persistId = vbaPersistId(doc, state);
			const offset = persistId === undefined ? undefined : state.persistOffsets.get(persistId);
			const record = offset === undefined ? undefined : readHeader(doc, offset);
			if (record && record.recType === RT_EX_OLE_OBJ_STG) {
				const storage = decodeStorageRecord(doc, record);
				if (storage && isVbaProjectStorage(storage)) {
					return { record, storage };
				}
			}
		}
	} catch {
		// Malformed persist chain: fall through to the scan.
	}
	let found: VbaStorageLocation | undefined;
	const walk = (from: number, to: number): void => {
		let cursor = from;
		while (cursor + 8 <= to) {
			const record = readHeader(doc, cursor);
			if (!record || record.end > to) {
				return;
			}
			if ((record.verInstance & 0xf) === 0xf) {
				walk(record.start + 8, record.end);
			} else if (record.recType === RT_EX_OLE_OBJ_STG) {
				const storage = decodeStorageRecord(doc, record);
				if (storage && isVbaProjectStorage(storage)) {
					found = { record, storage };
				}
			}
			cursor = record.end;
		}
	};
	walk(0, doc.length);
	return found;
}

// ----------------------------------------------------------------- public API

/** The embedded VBA project of a binary .ppt, as a parsed CFB. */
export function pptVbaCfb(cfb: Cfb): Cfb {
	const location = locateVbaStorage(cfb);
	if (!location) {
		throw new PptContainerError(
			'Presentation contains no VBA project (no ExOleObjStg record holding one).',
		);
	}
	return Cfb.fromBytes(location.storage);
}

/**
 * Splices a mutated VBA project storage back into the presentation: the
 * ExOleObjStg record is rebuilt (deflated, recInstance 1) at its current
 * position, and every absolute offset past the edit point - persist
 * directory entries, the user-edit chain, and the CurrentUserAtom - shifts
 * by the size delta. Returns the updated outer CFB, ready to serialize.
 */
export function pptWriteVbaStorage(cfb: Cfb, storage: Buffer): Cfb {
	const location = locateVbaStorage(cfb);
	if (!location) {
		throw new PptContainerError('Presentation has no VBA project record to replace.');
	}
	const doc = cfb.getStream('PowerPoint Document');
	const { record } = location;

	const deflated = zlib.deflateSync(storage, { level: 6 });
	const body = Buffer.alloc(4 + deflated.length);
	body.writeUInt32LE(storage.length, 0);
	deflated.copy(body, 4);
	const header = Buffer.alloc(8);
	header.writeUInt16LE((record.verInstance & 0xf) | 0x10, 0); // recVer kept, recInstance = 1
	header.writeUInt16LE(RT_EX_OLE_OBJ_STG, 2);
	header.writeUInt32LE(body.length, 4);

	const delta = header.length + body.length - (record.end - record.start);
	const updated = Buffer.concat([doc.subarray(0, record.start), header, body, doc.subarray(record.end)]);

	// Shift every absolute offset that pointed past the replaced record's
	// start. The record starts cannot themselves move unless they sit after
	// the edit point, so walking the UPDATED stream finds each carrier at
	// its (possibly shifted) location while the stored offset VALUES still
	// need the delta applied when they referenced later bytes.
	const shift = (value: number): number => (value > record.start ? value + delta : value);
	for (const top of topLevelRecords(updated)) {
		if (top.recType === RT_USER_EDIT_ATOM) {
			const payload = top.start + 8;
			updated.writeUInt32LE(shift(updated.readUInt32LE(payload + 8)), payload + 8);
			updated.writeUInt32LE(shift(updated.readUInt32LE(payload + 12)), payload + 12);
		} else if (top.recType === RT_PERSIST_DIRECTORY_ATOM) {
			let cursor = top.start + 8;
			while (cursor + 4 <= top.end) {
				const info = updated.readUInt32LE(cursor);
				cursor += 4;
				const count = info >>> 20;
				for (let i = 0; i < count && cursor + 4 <= top.end; i++) {
					updated.writeUInt32LE(shift(updated.readUInt32LE(cursor)), cursor);
					cursor += 4;
				}
			}
		}
	}

	const currentUser = Buffer.from(cfb.getStream('Current User'));
	const cuHeader = readHeader(currentUser, 0);
	if (cuHeader && cuHeader.recType === RT_CURRENT_USER_ATOM) {
		const at = cuHeader.start + 8 + 8;
		currentUser.writeUInt32LE(shift(currentUser.readUInt32LE(at)), at);
	}

	cfb.writeStream('PowerPoint Document', updated);
	cfb.writeStream('Current User', currentUser);
	return cfb;
}
