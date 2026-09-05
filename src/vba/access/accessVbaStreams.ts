import { encodeCodePage } from '../codePages';
import { readDirRecords } from '../vbaProject';
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
/** One folder's line in `Modules/PropData`, before its name. */
const FOLDER_ENTRY = Buffer.from('050902', 'hex');
const FOLDER_SUFFIX = Buffer.from('CB0', 'utf16le');
/** One entry in a `\x03DirData` payload: the tag, then the payload length. */
const ENTRY_TAG = 4;
const ENTRY_TRAILER = 4;
/**
 * The character a container's folder names start from, measured per container.
 * `Modules` carries four streams of its own - `PropData`, `PropDataCopy`,
 * `\x03DirData` and `\x03DirDataCopy` - and its folders start at `4`; `Forms`,
 * `Reports` and `Scripts` start at `0`.
 */
const FOLDER_BASE: Readonly<Record<string, number>> = {
	Modules: 4, Forms: 0, Reports: 0, Scripts: 0,
};
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

function record(id: number, payload: Buffer): Buffer {
	const head = Buffer.alloc(6);
	head.writeUInt16LE(id, 0);
	head.writeUInt32LE(payload.length, 2);
	return Buffer.concat([head, payload]);
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
		record(REC_MODULENAME, encodeCodePage(name, codePage)),
		record(REC_MODULENAME_UNICODE, Buffer.from(name, 'utf16le')),
		record(REC_MODULESTREAMNAME, encodeCodePage(streamName, codePage)),
		record(REC_MODULESTREAMNAME_UNICODE, Buffer.from(streamName, 'utf16le')),
		record(REC_MODULEDOCSTRING, Buffer.alloc(0)),
		record(REC_MODULEDOCSTRING_UNICODE, Buffer.alloc(0)),
		record(REC_MODULEOFFSET, Buffer.alloc(4)),
		record(REC_MODULEHELPCONTEXT, Buffer.alloc(4)),
		record(REC_MODULEEND2, cookie),
		record(kind === 'class' ? REC_MODULETYPE_CLASS : REC_MODULETYPE_PROCEDURAL, Buffer.alloc(0)),
		record(REC_MODULEEND, Buffer.alloc(0)),
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
		const header = record(id, encode(oldName));
		const at = out.indexOf(header);
		if (at < 0) {
			throw new AccessFormatError(
				`The dir stream has no 0x${id.toString(16)} record for ${oldName}.`,
			);
		}
		out = Buffer.concat([
			out.subarray(0, at), record(id, encode(newName)), out.subarray(at + header.length),
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

// --- the container's `\x03DirData` and folder list ---------------------------
//
// `<u32 0>` and then one entry each:
//
//     04 <u8 payload length> <name UTF-16> <u32 folder>
//
// where the payload length counts the name's bytes plus the four of the folder
// number. The trailing four bytes name the object's storage folder, not a
// terminator: a five-module project whose folders are 0, 4, 5, 6, 7 carries
// exactly those, and a module that reused a freed folder carries the reused
// name.

function dirDataPrefix(name: string): Buffer {
	const text = Buffer.from(name, 'utf16le');
	return Buffer.concat([Buffer.from([ENTRY_TAG, text.length + ENTRY_TRAILER]), text]);
}

export function dirDataEntries(payload: Buffer): Array<{ name: string; folder: string }> {
	const out: Array<{ name: string; folder: string }> = [];
	let at = 4;
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

export function addToDirData(payload: Buffer, name: string, folder: string): Buffer {
	const trailer = Buffer.alloc(ENTRY_TRAILER);
	trailer.writeUInt32LE(Number(folder), 0);
	return Buffer.concat([payload, dirDataPrefix(name), trailer]);
}

/** Drop an entry, the four bytes that belong to it included. */
export function removeFromDirData(payload: Buffer, name: string): Buffer {
	const prefix = dirDataPrefix(name);
	const at = payload.indexOf(prefix);
	if (at < 0) {
		throw new AccessFormatError(`DirData holds no entry for ${name}.`);
	}
	return Buffer.concat([
		payload.subarray(0, at), payload.subarray(at + prefix.length + ENTRY_TRAILER),
	]);
}

/** Rewrite an entry's name, leaving the folder it names alone. */
export function renameDirData(payload: Buffer, oldName: string, newName: string): Buffer {
	const prefix = dirDataPrefix(oldName);
	const at = payload.indexOf(prefix);
	if (at < 0) {
		throw new AccessFormatError(`DirData holds no entry for ${oldName}.`);
	}
	return Buffer.concat([
		payload.subarray(0, at), dirDataPrefix(newName), payload.subarray(at + prefix.length),
	]);
}

export function addToFolderList(payload: Buffer, folder: string): Buffer {
	return Buffer.concat([payload, FOLDER_ENTRY, Buffer.from(folder, 'utf16le'), FOLDER_SUFFIX]);
}

export function removeFromFolderList(payload: Buffer, folder: string): Buffer {
	const entry = Buffer.concat([FOLDER_ENTRY, Buffer.from(folder, 'utf16le'), FOLDER_SUFFIX]);
	const at = payload.indexOf(entry);
	return at < 0
		? Buffer.from(payload)
		: Buffer.concat([payload.subarray(0, at), payload.subarray(at + entry.length)]);
}

/**
 * The name Access gives a new object's storage folder. It is computed, not
 * chosen, and Access will not find an object in a folder by any other name:
 * `AllModules(i).Name` fails on a module in the wrong one while the VBE still
 * lists and runs it. Names are allocated lowest-free from the container's base.
 */
export function nextFolderName(container: string, taken: ReadonlySet<string>): string {
	let code = '0'.charCodeAt(0) + (FOLDER_BASE[container] ?? 0);
	while (taken.has(String.fromCharCode(code))) {
		code += 1;
	}
	return String.fromCharCode(code);
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
