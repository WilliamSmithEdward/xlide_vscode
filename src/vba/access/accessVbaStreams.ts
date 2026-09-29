import { encodeCodePage } from '../codePages';
import { dirRecord, readDirRecords } from '../vbaProject';
import { AccessFormatError } from './accessFormat';

/**
 * The byte and text edits a module add, rename or delete makes to the streams
 * the VBA project keeps beside its modules.
 *
 * A module's name lives in eight places: two dir records, the module's own
 * `Attribute VB_Name`, `PROJECT`, `PROJECTwm`, the container's `\x03DirData`,
 * the catalog row and the navigation pane. This file owns the four that are
 * stream bytes; `accessVbaWriter.ts` owns which rows they belong to.
 *
 * Ported from pyOpenVBA's `_vba.py` and `_storage.py`.
 */

const CRLF = '\r\n';
const QUOTE = '"';

// --- dir stream records ------------------------------------------------------
const REC_PROJECTMODULES = 0x000f;
const REC_TERMINATOR = 0x0010;
const REC_MODULENAME = 0x0019;
const REC_MODULESTREAMNAME = 0x001a;
const REC_MODULEDOCSTRING = 0x001c;
const REC_MODULEHELPCONTEXT = 0x001e;
const REC_MODULETYPE_PROCEDURAL = 0x0021;
const REC_MODULETYPE_CLASS = 0x0022;
const REC_MODULEEND = 0x002b;
const REC_MODULEEND2 = 0x002c;
const REC_MODULEOFFSET = 0x0031;
const REC_MODULESTREAMNAME_UNICODE = 0x0032;
const REC_MODULENAME_UNICODE = 0x0047;
const REC_MODULEDOCSTRING_UNICODE = 0x0048;

export type AccessModuleKind = 'module' | 'class';

/**
 * Access's class-module base, measured off a class the VBE added. A class
 * stream without it loads but will not instantiate.
 */
const CLASS_BASE = '0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}';
const CLASS_ATTRIBUTES: ReadonlyArray<[string, string]> = [
	['VB_Base', QUOTE + CLASS_BASE + QUOTE],
	['VB_GlobalNameSpace', 'False'],
	['VB_Creatable', 'False'],
	['VB_PredeclaredId', 'False'],
	['VB_Exposed', 'False'],
	['VB_TemplateDerived', 'False'],
	['VB_Customizable', 'False'],
];

/** Every object's storage folder holds this, unchanging, 13 bytes. */
export const PROP_DATA = Buffer.from('00000000020000000000000000', 'hex');
/** A folder's line in a container's `PropData` opens with this tag. */
const FOLDER_TAG = 0x05;
/** Both container lists open with four bytes before their first entry. */
const LIST_HEADER = 4;
/** One entry in a `\x03DirData` payload: the tag, then the payload length. */
const ENTRY_TAG = 4;
const ENTRY_TRAILER = 4;
const STREAM_NAME_LENGTH = 28;

/** The attributes a module's source opens with; a class carries seven more. */
export function attributeLines(name: string, kind: AccessModuleKind): string[] {
	const lines = [`Attribute VB_Name = ${QUOTE}${name}${QUOTE}`];
	if (kind === 'class') {
		lines.push(...CLASS_ATTRIBUTES.map(([field, value]) => `Attribute ${field} = ${value}`));
	}
	return lines;
}

/** A module's leading `Attribute` block and everything after it. */
export function splitModuleSource(text: string): { attributes: string[]; body: string[] } {
	const lines = text.split(CRLF);
	let at = 0;
	while (at < lines.length && lines[at].startsWith('Attribute ')) {
		at += 1;
	}
	return { attributes: lines.slice(0, at), body: lines.slice(at) };
}

/**
 * The eleven records a module contributes to the dir stream. A character the
 * code page cannot hold folds to `?` in the ANSI record and stays exact in the
 * Unicode one beside it, which is what the VBE writes.
 */
export function moduleDirBlock(
	name: string,
	streamName: string,
	cookie: Buffer,
	kind: AccessModuleKind,
	codePage: number,
): Buffer {
	return Buffer.concat([
		dirRecord(REC_MODULENAME, encodeCodePage(name, codePage)),
		dirRecord(REC_MODULENAME_UNICODE, Buffer.from(name, 'utf16le')),
		dirRecord(REC_MODULESTREAMNAME, encodeCodePage(streamName, codePage)),
		dirRecord(REC_MODULESTREAMNAME_UNICODE, Buffer.from(streamName, 'utf16le')),
		dirRecord(REC_MODULEDOCSTRING, Buffer.alloc(0)),
		dirRecord(REC_MODULEDOCSTRING_UNICODE, Buffer.alloc(0)),
		dirRecord(REC_MODULEOFFSET, Buffer.alloc(4)),
		dirRecord(REC_MODULEHELPCONTEXT, Buffer.alloc(4)),
		dirRecord(REC_MODULEEND2, cookie),
		dirRecord(kind === 'class' ? REC_MODULETYPE_CLASS : REC_MODULETYPE_PROCEDURAL, Buffer.alloc(0)),
		dirRecord(REC_MODULEEND, Buffer.alloc(0)),
	]);
}

function setModuleCount(dir: Buffer, delta: number): Buffer {
	const out = Buffer.from(dir);
	for (const rec of readDirRecords(out)) {
		if (rec.id === REC_PROJECTMODULES && rec.dataEnd - rec.dataStart === 2) {
			out.writeUInt16LE(out.readUInt16LE(rec.dataStart) + delta, rec.dataStart);
			break;
		}
	}
	return out;
}

/** Insert a module's block before the terminator and count it. */
export function addToDir(dir: Buffer, block: Buffer): Buffer {
	let at: number | undefined;
	for (const rec of readDirRecords(dir)) {
		if (rec.id === REC_TERMINATOR) {
			at = rec.start;
		}
	}
	if (at === undefined) {
		throw new AccessFormatError('The dir stream has no terminator.');
	}
	return setModuleCount(
		Buffer.concat([dir.subarray(0, at), block, dir.subarray(at)]), 1,
	);
}

/** Drop a module's block and take one off the module count. */
export function removeFromDir(dir: Buffer, name: string, codePage: number): Buffer {
	const want = encodeCodePage(name, codePage);
	let start: number | undefined;
	let end: number | undefined;
	for (const rec of readDirRecords(dir)) {
		if (rec.id === REC_MODULENAME) {
			if (dir.subarray(rec.dataStart, rec.dataEnd).equals(want)) {
				start = rec.start;
			} else if (start !== undefined && end === undefined) {
				end = rec.start;
			}
		} else if (rec.id === REC_MODULEEND && start !== undefined && end === undefined
			&& rec.start > start) {
			end = rec.start + 6;
		}
	}
	if (start === undefined || end === undefined) {
		throw new AccessFormatError(`The dir stream has no module block for ${name}.`);
	}
	return setModuleCount(Buffer.concat([dir.subarray(0, start), dir.subarray(end)]), -1);
}

/** Rewrite a module's two name records. */
export function renameInDir(dir: Buffer, oldName: string, newName: string, codePage: number): Buffer {
	let out = Buffer.from(dir);
	for (const [id, encode] of [
		[REC_MODULENAME, (text: string) => encodeCodePage(text, codePage)],
		[REC_MODULENAME_UNICODE, (text: string) => Buffer.from(text, 'utf16le')],
	] as const) {
		const header = dirRecord(id, encode(oldName));
		const at = out.indexOf(header);
		if (at < 0) {
			throw new AccessFormatError(
				`The dir stream has no 0x${id.toString(16)} record for ${oldName}.`,
			);
		}
		out = Buffer.concat([
			out.subarray(0, at), dirRecord(id, encode(newName)), out.subarray(at + header.length),
		]);
	}
	return out;
}

/** Where a module's MODULEOFFSET payload starts in the dir stream. */
export function moduleOffsetAt(dir: Buffer, name: string, codePage: number): number {
	const want = encodeCodePage(name, codePage);
	let seen = false;
	for (const rec of readDirRecords(dir)) {
		if (rec.id === REC_MODULENAME) {
			seen = dir.subarray(rec.dataStart, rec.dataEnd).equals(want);
		} else if (rec.id === REC_MODULEOFFSET && seen) {
			return rec.dataStart;
		}
	}
	throw new AccessFormatError(`The dir stream has no MODULEOFFSET for ${name}.`);
}

// --- PROJECTwm ---------------------------------------------------------------

function projectWmEntry(name: string, codePage: number): Buffer {
	return Buffer.concat([
		encodeCodePage(name, codePage), Buffer.alloc(1),
		Buffer.from(name, 'utf16le'), Buffer.alloc(2),
	]);
}

export function addToProjectWm(payload: Buffer, name: string, codePage: number): Buffer {
	return Buffer.concat([
		payload.subarray(0, payload.length - 2), projectWmEntry(name, codePage), Buffer.alloc(2),
	]);
}

export function removeFromProjectWm(payload: Buffer, name: string, codePage: number): Buffer {
	const want = projectWmEntry(name, codePage);
	const at = payload.indexOf(want);
	if (at < 0) {
		throw new AccessFormatError(`PROJECTwm holds no entry for ${name}.`);
	}
	return Buffer.concat([payload.subarray(0, at), payload.subarray(at + want.length)]);
}

export function renameProjectWm(
	payload: Buffer, oldName: string, newName: string, codePage: number,
): Buffer {
	const want = projectWmEntry(oldName, codePage);
	const at = payload.indexOf(want);
	if (at < 0) {
		throw new AccessFormatError(`PROJECTwm holds no entry for ${oldName}.`);
	}
	return Buffer.concat([
		payload.subarray(0, at), projectWmEntry(newName, codePage),
		payload.subarray(at + want.length),
	]);
}

// --- the PROJECT stream ------------------------------------------------------

/**
 * Access lists a standard module as `Module=` and a class as `Class=`, both in
 * the same block, and gives each a window rectangle under `[Workspace]`.
 *
 * A project with no modules yet has no such line to sit beside. The block
 * belongs between the project's `ID` and its `Name`, which is where Access
 * writes the first one, so an empty project anchors on `ID=` instead.
 */
export function addToProject(text: string, name: string, kind: AccessModuleKind): string {
	const lines = text.split(CRLF);
	let last = -1;
	lines.forEach((line, index) => {
		if (line.startsWith('Module=') || line.startsWith('Class=')
			|| line.startsWith(`${DOC_CLASS}=`)) {
			last = index;
		}
	});
	if (last < 0) {
		last = lines.findIndex((line) => line.startsWith('ID='));
	}
	if (last < 0) {
		throw new AccessFormatError('PROJECT has neither a module list nor an ID to add beside.');
	}
	lines.splice(last + 1, 0, (kind === 'class' ? 'Class=' : 'Module=') + name);
	if (lines.some((line) => line.trim() === '[Workspace]')) {
		lines.splice(lines.length - 1, 0, `${name}=38, 38, 1786, 1030, `);
	}
	return lines.join(CRLF);
}

export function removeFromProject(text: string, name: string): string {
	return text.split(CRLF)
		.filter((line) => line !== `Module=${name}` && line !== `Class=${name}`
			&& !line.startsWith(`${DOC_CLASS}=${name}/`)
			&& !line.startsWith(`${name}=`))
		.join(CRLF);
}

/**
 * How `PROJECT` names the module behind a form or report: `DocClass=` and the
 * module's name, then a slash and a flag word Access owns. A design's module
 * is listed this way and never as `Module=` or `Class=`, and Access reads a
 * `DocClass` naming a module the project no longer has as a corrupt project.
 */
const DOC_CLASS = 'DocClass';

/**
 * The `Module=`, `Class=` or `DocClass=` line and the `[Workspace]` line. The
 * stream's lines end CR LF, so the end anchor has to allow the CR.
 */
export function renameProject(text: string, oldName: string, newName: string): string {
	const quoted = oldName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
	let out = text;
	for (const keyword of ['Module', 'Class']) {
		out = out.replace(
			new RegExp(`^${keyword}=${quoted}(?=\\r?$)`, 'gm'), `${keyword}=${newName}`,
		);
	}
	out = out.replace(
		new RegExp(`^${DOC_CLASS}=${quoted}(?=/)`, 'gm'), `${DOC_CLASS}=${newName}`,
	);
	return out.replace(new RegExp(`^${quoted}=`, 'gm'), `${newName}=`);
}

// --- the order Access keeps a container's lists in ---------------------------
// Access loads a container's `\x03DirData` into an MSVC std::unordered_map
// keyed by object name, and its `PropData` folder list into another keyed by
// folder name, and writes each back in the map's own order whenever it
// changes (issue #150, from pyOpenVBA 6.3.0, which read the hash and the map
// out of MSACCESS.EXE 16.0). Checked here against Access 16.0 driven over
// COM: a blank database given Module1 and then Macro1 to Macro11, one
// deleted, one added, one renamed, lists every step exactly as replayed.

/** ASCII the name hash skips: controls other than tab, both quotes, `~` and DEL. */
function hashSkips(code: number): boolean {
	return (code < 0x20 && code !== 0x09) || code === 0x22 || code === 0x27 || code === 0x7e || code === 0x7f;
}

/** cp1252 small letters whose capitals differ in more than the case bit. */
const CP1252_CAPITALS: Readonly<Record<number, number>> = { 0x9a: 0x8a, 0x9c: 0x8c, 0x9e: 0x8e };

/**
 * The 16-bit hash Access files a name under in a container's lists. Each
 * character that counts adds its low five bits: `h = (h << 5) + (h >> 13) +
 * 1 + bits` in 16 bits, so case never matters. A leading `.` is skipped; a
 * space or tab counts as 0. From the first character past ASCII, the rest of
 * the name is taken a byte at a time in the ANSI code page (cp1252 on Western
 * Windows), upper-cased, every byte above 1 counting.
 */
export function accessNameHash(name: string): number {
	const text = name.startsWith('.') ? name.slice(1) : name;
	const values: number[] = [];
	for (let at = 0; at < text.length; at += 1) {
		const code = text.charCodeAt(at);
		if (code >= 0x80) {
			for (const byte of encodeCodePage(text.slice(at), 1252)) {
				if (byte > 1) {
					values.push((CP1252_CAPITALS[byte] ?? byte) & 0x1f);
				}
			}
			break;
		}
		if (!hashSkips(code)) {
			values.push(code === 0x20 || code === 0x09 ? 0 : code & 0x1f);
		}
	}
	let hashed = 0;
	for (const value of values) {
		hashed = (((hashed << 5) & 0xffff) + (hashed >> 13) + 1 + value) & 0xffff;
	}
	return hashed;
}

/** A new map's bucket count; below 512 buckets the map grows eightfold. */
const FIRST_BUCKETS = 8;
const EIGHTFOLD_BELOW = 512;

/**
 * Access's map of a container list's keys, in the order it keeps them. One
 * list holds every key with each bucket's keys together. A new key goes in
 * front of the first key of its bucket, or at the end when the bucket is
 * empty. An insert that would leave more keys than buckets first grows the
 * map to a power of two, eightfold while under 512 buckets, and the rehash
 * walks the list moving each key to the front of its new bucket, the buckets
 * in the order the walk first meets them.
 */
class ContainerMap {
	keys: string[] = [];
	private buckets = FIRST_BUCKETS;

	constructor(stored: readonly string[]) {
		for (const key of stored) {
			this.insert(key);
		}
	}

	private bucket(key: string): number {
		return accessNameHash(key) & (this.buckets - 1);
	}

	insert(key: string): void {
		if (this.keys.length + 1 > this.buckets) {
			let wanted = this.keys.length + 1;
			if (this.buckets < EIGHTFOLD_BELOW) {
				wanted = Math.max(wanted, this.buckets * 8);
			}
			let size = 1;
			while (size < wanted) { size *= 2; }
			this.buckets = size;
			const chains = new Map<number, string[]>();
			for (const moved of this.keys) {
				const chain = chains.get(this.bucket(moved));
				if (chain) { chain.unshift(moved); } else { chains.set(this.bucket(moved), [moved]); }
			}
			this.keys = [...chains.values()].flat();
		}
		const bucket = this.bucket(key);
		const first = this.keys.findIndex((other) => this.bucket(other) === bucket);
		if (first < 0) {
			this.keys.push(key);
		} else {
			this.keys.splice(first, 0, key);
		}
	}
}

/**
 * The keys of a container list after Access loads it as `stored`, erases
 * `remove` and inserts `add`, in the order it writes them. Loading inserts the
 * stored keys in stored order, so an erase can move keys it never named. A
 * rename is an erase and an insert.
 */
export function accessListOrder(
	stored: readonly string[],
	remove: readonly string[] = [],
	add: readonly string[] = [],
): string[] {
	const map = new ContainerMap(stored);
	for (const key of remove) {
		const at = map.keys.indexOf(key);
		if (at >= 0) { map.keys.splice(at, 1); }
	}
	for (const key of add) {
		map.insert(key);
	}
	return map.keys;
}

// --- the container's `\x03DirData` -------------------------------------------
// `<u32 0>` and then one entry each:
//
//     04 <u8 payload length> <name UTF-16> <u32 folder>
//
// where the payload length counts the name's bytes plus the four of the folder
// number. The trailing four bytes name the object's storage folder, not a
// terminator: a module that reused a freed folder carries the reused name.

function dirDataPrefix(name: string): Buffer {
	const text = Buffer.from(name, 'utf16le');
	return Buffer.concat([Buffer.from([ENTRY_TAG, text.length + ENTRY_TRAILER]), text]);
}

function dirDataEntry(name: string, folder: string): Buffer {
	const trailer = Buffer.alloc(ENTRY_TRAILER);
	trailer.writeUInt32LE(Number(folder), 0);
	return Buffer.concat([dirDataPrefix(name), trailer]);
}

/** The names in stored order, each name's whole entry, and whatever follows the last. */
function dirDataParts(payload: Buffer): { names: string[]; entries: Map<string, Buffer>; tail: Buffer } {
	const names: string[] = [];
	const entries = new Map<string, Buffer>();
	let at = LIST_HEADER;
	while (at + 2 <= payload.length && payload[at] === ENTRY_TAG) {
		const end = at + 2 + payload[at + 1];
		const name = payload.subarray(at + 2, end - ENTRY_TRAILER).toString('utf16le');
		names.push(name);
		entries.set(name, payload.subarray(at, end));
		at = end;
	}
	return { names, entries, tail: payload.subarray(at) };
}

function dirDataIn(payload: Buffer, order: readonly string[], entries: ReadonlyMap<string, Buffer>, tail: Buffer): Buffer {
	return Buffer.concat([payload.subarray(0, LIST_HEADER), ...order.map((name) => entries.get(name)!), tail]);
}

export function dirDataEntries(payload: Buffer): Array<{ name: string; folder: string }> {
	const out: Array<{ name: string; folder: string }> = [];
	let at = LIST_HEADER;
	while (at + 2 <= payload.length && payload[at] === ENTRY_TAG) {
		const size = payload[at + 1];
		const body = payload.subarray(at + 2, at + 2 + size);
		const folder = body.readUInt32LE(body.length - ENTRY_TRAILER);
		out.push({
			name: body.subarray(0, body.length - ENTRY_TRAILER).toString('utf16le'),
			folder: String(folder),
		});
		at += 2 + size;
	}
	return out;
}

/** List a new object where Access puts it (see `accessListOrder`). */
export function addToDirData(payload: Buffer, name: string, folder: string): Buffer {
	const { names, entries, tail } = dirDataParts(payload);
	entries.set(name, dirDataEntry(name, folder));
	return dirDataIn(payload, accessListOrder(names, [], [name]), entries, tail);
}

/** Drop an entry, its four folder bytes included, leaving the rest as Access orders them. */
export function removeFromDirData(payload: Buffer, name: string): Buffer {
	const { names, entries, tail } = dirDataParts(payload);
	if (!entries.has(name)) {
		throw new AccessFormatError(`DirData holds no entry for ${name}.`);
	}
	return dirDataIn(payload, accessListOrder(names, [name]), entries, tail);
}

/** Rename an entry, keeping its folder. Access erases and inserts, so the entry moves. */
export function renameDirData(payload: Buffer, oldName: string, newName: string): Buffer {
	const { names, entries, tail } = dirDataParts(payload);
	const old = entries.get(oldName);
	if (!old) {
		throw new AccessFormatError(`DirData holds no entry for ${oldName}.`);
	}
	entries.set(newName, Buffer.concat([dirDataPrefix(newName), old.subarray(old.length - ENTRY_TRAILER)]));
	return dirDataIn(payload, accessListOrder(names, [oldName], [newName]), entries, tail);
}

// --- the container's `PropData` folder list ------------------------------------
// Access adds an object's line here the next time it opens the database, not
// when the object is made (measured: a module added in one session appears
// in the list only after the next open), so adding writes none. A line is
//
//     05 <1 + 2n + 6> <2n> <folder name UTF-16> "CB0" UTF-16
//
// with 2n the name's size in bytes: `05 09 02` for folder `4`, `05 0b 04` for
// folder `10`.

/** The folders in stored order, each folder's whole line, and whatever follows the last. */
function folderListParts(payload: Buffer): { folders: string[]; lines: Map<string, Buffer>; tail: Buffer } {
	const folders: string[] = [];
	const lines = new Map<string, Buffer>();
	let at = LIST_HEADER;
	while (at + 3 <= payload.length && payload[at] === FOLDER_TAG) {
		const end = at + 2 + payload[at + 1];
		const folder = payload.subarray(at + 3, at + 3 + payload[at + 2]).toString('utf16le');
		folders.push(folder);
		lines.set(folder, payload.subarray(at, end));
		at = end;
	}
	return { folders, lines, tail: payload.subarray(at) };
}

/**
 * The list with `folder`'s line erased, and inserted again when `again`, in
 * the order Access writes it after. A list without the line is left alone,
 * as Access leaves it.
 */
function inFolderList(payload: Buffer, folder: string, again: boolean): Buffer {
	const { folders, lines, tail } = folderListParts(payload);
	if (!lines.has(folder)) {
		return Buffer.from(payload);
	}
	const order = accessListOrder(folders, [folder], again ? [folder] : []);
	return Buffer.concat([payload.subarray(0, LIST_HEADER), ...order.map((name) => lines.get(name)!), tail]);
}

export function removeFromFolderList(payload: Buffer, folder: string): Buffer {
	return inFolderList(payload, folder, false);
}

/** A renamed design's line: `DoCmd.Rename` erases and inserts it, so it moves. */
export function refileInFolderList(payload: Buffer, folder: string): Buffer {
	return inFolderList(payload, folder, true);
}

/**
 * The name Access gives a new object's storage folder: the lowest free
 * decimal number from 0, in every container (measured against Access 16.0
 * over COM, issue #150). Access will not find an object in a folder by any
 * other name: `AllModules(i).Name` fails on a module in the wrong one while
 * the VBE still lists and runs it.
 */
export function nextFolderName(taken: ReadonlySet<string>): string {
	let number = 0;
	while (taken.has(String(number))) {
		number += 1;
	}
	return String(number);
}

/** A module's storage row name: 28 random capitals, unused. */
export function newStreamRowName(taken: ReadonlySet<string>, random: () => number): string {
	for (;;) {
		let name = '';
		for (let i = 0; i < STREAM_NAME_LENGTH; i += 1) {
			name += String.fromCharCode(65 + Math.floor(random() * 26));
		}
		if (!taken.has(name)) {
			return name;
		}
	}
}
