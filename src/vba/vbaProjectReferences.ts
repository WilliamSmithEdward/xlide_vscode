// [MS-OVBA 2.3.4.2.2] The library references a VBA project declares.
//
// The engine reuses the dir stream's reference section verbatim when it
// rewrites a project, which is right for every edit that does not change what
// the project depends on. Adding a UserForm does change it: a project holding
// a form must reference the Microsoft Forms library, or the form cannot be
// instantiated and nothing in the project compiles. A project that gained its
// first form without one looks broken in its host while reading perfectly
// here. This is host-independent - Excel, Word and PowerPoint projects all
// need it, and none of XLIDE's blank templates ships with it.
//
// The record shapes below were measured against workbooks Excel itself wrote
// (tests/fixtures/binaries/FormFixtureVbide.xlsm is the oracle, and a test
// rebuilds its Microsoft Forms block byte for byte from its own parsed
// fields). Nothing here is guessed from prose.

import { dirRecord, readDirRecords, REC_PROJECTMODULES } from './vbaProject';

const REC_REFERENCE_NAME = 0x0016;
const REC_REFERENCE_NAME_UNICODE = 0x003e;
const REC_REFERENCE_REGISTERED = 0x000d;
const REC_REFERENCE_PROJECT = 0x000e;
const REC_REFERENCE_CONTROL = 0x002f;
const REC_REFERENCE_ORIGINAL = 0x0033;
const REC_REFERENCE_CONTROL_EXTENDED = 0x0030;

/** The Microsoft Forms 2.0 type library, which every UserForm needs. */
export const MSFORMS_REFERENCE_NAME = 'MSForms';
export const MSFORMS_TYPELIB_GUID = '{0D452EE1-E08F-101A-852E-02608C4D0BB4}';
/**
 * The path inside a libid is a hint: the host resolves the library through the
 * registry by GUID and version, which is why a file written on one machine
 * loads on another whose FM20.DLL sits elsewhere. This is the path Excel itself
 * writes, measured on a machine whose registry points at a different location
 * entirely, so it is what the format carries rather than what is installed.
 */
export const MSFORMS_DEFAULT_PATH = 'C:\\WINDOWS\\system32\\FM20.DLL';
const MSFORMS_DESCRIPTION = 'Microsoft Forms 2.0 Object Library';
/** A control reference's twiddled libid is the null one; the real identity is the extended half. */
const NULL_LIBID = '*\\G{00000000-0000-0000-0000-000000000000}#0.0#0##';

export interface VbaProjectReference {
	/** The name the project knows it by, as `Tools > References` shows it. */
	name: string;
	kind: 'registered' | 'project' | 'control';
	/** The libid as written, `*\G{guid}#major.minor#lcid#path#description`. */
	libid: string;
}

/** A reference with the byte span its records occupy in the dir stream. */
export interface VbaProjectReferenceSpan {
	reference: VbaProjectReference;
	/** First byte of the reference's records, its name records included. */
	start: number;
	/** One past the last byte of them. */
	end: number;
}

/**
 * Every reference with the span its records occupy, which is what removing
 * one needs: a reference is not one record but a run of them - the name in
 * both encodings, then the record carrying the libid, and for a control
 * reference the original libid, the control record, the name again and the
 * extended record ([MS-OVBA] 2.3.4.2.2.3).
 *
 * The name records are optional in the format, so a span starts at the run of
 * names in front of the body record where there is one, and at the body
 * record itself where there is not. Scanning stops at the module section,
 * which is where the reference section ends.
 */
export function readProjectReferenceSpans(dir: Buffer): VbaProjectReferenceSpan[] {
	const out: VbaProjectReferenceSpan[] = [];
	let name = '';
	let namesStart: number | undefined;
	// A control reference spans several records, so its entry is pushed at
	// the first and its end moved along as the rest arrive.
	let control: number | undefined;
	for (const record of readDirRecords(dir)) {
		const body = dir.subarray(record.dataStart, record.dataEnd);
		if (record.id === REC_PROJECTMODULES) {
			break;
		}
		switch (record.id) {
			case REC_REFERENCE_NAME:
				name = body.toString('latin1');
				namesStart ??= record.start;
				continue;
			case REC_REFERENCE_NAME_UNICODE:
				namesStart ??= record.start;
				continue;
			case REC_REFERENCE_REGISTERED:
				// SizeOfLibid then the libid; the trailing reserved fields are ignored.
				out.push({
					reference: { name, kind: 'registered', libid: sizedString(body) },
					start: namesStart ?? record.start,
					end: record.end,
				});
				break;
			case REC_REFERENCE_PROJECT:
				out.push({
					reference: { name, kind: 'project', libid: sizedString(body) },
					start: namesStart ?? record.start,
					end: record.end,
				});
				break;
			case REC_REFERENCE_ORIGINAL:
				// The record's own size IS the libid's, so the body is the libid.
				control = out.length;
				out.push({
					reference: { name, kind: 'control', libid: body.toString('latin1') },
					start: namesStart ?? record.start,
					end: record.end,
				});
				break;
			case REC_REFERENCE_CONTROL:
				if (control !== undefined) { out[control].end = record.end; }
				break;
			case REC_REFERENCE_CONTROL_EXTENDED:
				if (control !== undefined) { out[control].end = record.end; }
				control = undefined;
				break;
			default:
				break;
		}
		namesStart = undefined;
	}
	return out;
}

/**
 * Every reference the dir stream declares, in order. The spans carry the same
 * list, so this is what a caller that only wants the list gets.
 */
export function readProjectReferences(dir: Buffer): VbaProjectReference[] {
	return readProjectReferenceSpans(dir).map((span) => span.reference);
}

/**
 * Cuts a reference's records out of a dir stream, leaving every other byte as
 * it was. Returns the stream unchanged when nothing matches.
 */
export function removeReferenceRecords(
	dir: Buffer,
	matches: (reference: VbaProjectReference) => boolean,
): Buffer {
	const cuts = readProjectReferenceSpans(dir)
		.filter((span) => matches(span.reference))
		.sort((a, b) => a.start - b.start);
	if (cuts.length === 0) {
		return dir;
	}
	const parts: Buffer[] = [];
	let kept = 0;
	for (const span of cuts) {
		parts.push(dir.subarray(kept, span.start));
		kept = span.end;
	}
	parts.push(dir.subarray(kept));
	return Buffer.concat(parts);
}

/** Whether the project already declares the Microsoft Forms library. */
export function hasMsFormsReference(dir: Buffer): boolean {
	return readProjectReferences(dir).some(
		(reference) => reference.name.toLowerCase() === MSFORMS_REFERENCE_NAME.toLowerCase(),
	);
}

function sizedString(body: Buffer): string {
	if (body.length < 4) {
		return '';
	}
	const size = body.readUInt32LE(0);
	return size > 0 && 4 + size <= body.length ? body.subarray(4, 4 + size).toString('latin1') : '';
}

function sized(text: string): Buffer {
	const bytes = Buffer.from(text, 'latin1');
	const size = Buffer.alloc(4);
	size.writeUInt32LE(bytes.length, 0);
	return Buffer.concat([size, bytes]);
}

/** A reference's name, in both the code page and Unicode forms the format carries. */
function nameRecords(name: string): Buffer {
	return Buffer.concat([
		dirRecord(REC_REFERENCE_NAME, Buffer.from(name, 'latin1')),
		dirRecord(REC_REFERENCE_NAME_UNICODE, Buffer.from(name, 'utf16le')),
	]);
}

/** A GUID in the mixed-endian layout OLE stores: three little-endian fields, then eight bytes as written. */
export function guidBytes(guid: string): Buffer {
	const hex = guid.replace(/[{}-]/g, '');
	if (!/^[0-9a-fA-F]{32}$/.test(hex)) {
		throw new Error(`Not a GUID: ${guid}`);
	}
	const raw = Buffer.from(hex, 'hex');
	return Buffer.concat([
		Buffer.from([raw[3], raw[2], raw[1], raw[0]]),
		Buffer.from([raw[5], raw[4]]),
		Buffer.from([raw[7], raw[6]]),
		raw.subarray(8),
	]);
}

export interface ControlReference {
	name: string;
	/** The library as registered: what Tools > References resolves. */
	libidOriginal: string;
	/** The twiddled libid, which for a control reference is the null one. */
	libidTwiddled?: string;
	/** The current resolution. Excel rewrites this to its own `.exd` cache once the VBE builds one. */
	libidExtended?: string;
	/** The type library's GUID, which is how the reference actually resolves. */
	typeLibGuid: string;
	cookie?: number;
}

/**
 * The records a control reference occupies: its name, the original libid, the
 * control record, the name again, and the extended record carrying the type
 * library's own GUID.
 */
export function buildControlReference(reference: ControlReference): Buffer {
	const twiddled = reference.libidTwiddled ?? NULL_LIBID;
	const extended = reference.libidExtended ?? reference.libidOriginal;
	const reserved1and2 = Buffer.alloc(6);
	const control = Buffer.concat([sized(twiddled), reserved1and2]);
	const tail = Buffer.alloc(4);
	tail.writeUInt32LE(reference.cookie ?? 1, 0);
	const extendedBody = Buffer.concat([
		sized(extended),
		Buffer.alloc(6),
		guidBytes(reference.typeLibGuid),
		tail,
	]);
	return Buffer.concat([
		nameRecords(reference.name),
		dirRecord(REC_REFERENCE_ORIGINAL, Buffer.from(reference.libidOriginal, 'latin1')),
		dirRecord(REC_REFERENCE_CONTROL, control),
		nameRecords(reference.name),
		dirRecord(REC_REFERENCE_CONTROL_EXTENDED, extendedBody),
	]);
}

/** The libid text for a type library, as the format spells it. */
export function libid(guid: string, version: string, path: string, description: string): string {
	return `*\\G${guid}#${version}#0#${path}#${description}`;
}

/**
 * The Office application type libraries a project can be given a reference
 * to, so its VBA can name another application's types early-bound.
 *
 * Every field was read from the registered type library: the GUID, the
 * version its own `GetLibAttr` declares, the path the registry resolves it
 * at, and the library's own description. The Excel entry is cross-checked
 * against the reference Word itself wrote into WordExcelInteropFixture.docm,
 * which matches field for field.
 *
 * Probing `LoadRegTypeLib` upwards from 0.0 is NOT how to get the version:
 * it answers with the first that loads, which is 1.0 for Excel where the
 * library declares, and Office writes, 1.9.
 */
export interface HostLibrary {
	/** The name the project knows it by, and the qualifier VBA writes. */
	name: string;
	guid: string;
	version: string;
	path: string;
	description: string;
}

export const HOST_LIBRARIES: Readonly<Record<string, HostLibrary>> = Object.freeze({
	excel: {
		name: 'Excel',
		guid: '{00020813-0000-0000-C000-000000000046}',
		version: '1.9',
		path: 'C:\\Program Files\\Microsoft Office\\root\\Office16\\EXCEL.EXE',
		description: 'Microsoft Excel 16.0 Object Library',
	},
	word: {
		name: 'Word',
		guid: '{00020905-0000-0000-C000-000000000046}',
		version: '8.7',
		path: 'C:\\Program Files\\Microsoft Office\\root\\Office16\\MSWORD.OLB',
		description: 'Microsoft Word 16.0 Object Library',
	},
	powerpoint: {
		name: 'PowerPoint',
		guid: '{91493440-5A91-11CF-8700-00AA0060263B}',
		version: '2.12',
		path: 'C:\\Program Files\\Microsoft Office\\root\\Office16\\MSPPT.OLB',
		description: 'Microsoft PowerPoint 16.0 Object Library',
	},
	access: {
		name: 'Access',
		guid: '{4AFFC9A0-5F99-101B-AF4E-00AA003F0F07}',
		version: '9.0',
		path: 'C:\\Program Files\\Microsoft Office\\root\\Office16\\MSACC.OLB',
		description: 'Microsoft Access 16.0 Object Library',
	},
});

/**
 * The records a REGISTERED reference occupies: its name, then the libid with
 * the two reserved fields [MS-OVBA 2.3.4.2.2.5]. Simpler than a control
 * reference, which carries a twiddled libid and the library's GUID as well.
 */
export function buildRegisteredReference(library: HostLibrary): Buffer {
	const text = libid(library.guid, library.version, library.path, library.description);
	return Buffer.concat([
		nameRecords(library.name),
		dirRecord(REC_REFERENCE_REGISTERED, Buffer.concat([
			sized(text),
			// Reserved1 (4 bytes) and Reserved2 (2 bytes), both zero.
			Buffer.alloc(6),
		])),
	]);
}

/** The Microsoft Forms reference records, pointing at `fm20Path` for the library. */
export function buildMsFormsReference(fm20Path = MSFORMS_DEFAULT_PATH): Buffer {
	return buildControlReference({
		name: MSFORMS_REFERENCE_NAME,
		libidOriginal: libid(MSFORMS_TYPELIB_GUID, '2.0', fm20Path, MSFORMS_DESCRIPTION),
		typeLibGuid: MSFORMS_TYPELIB_GUID,
	});
}

/**
 * Puts reference records into a dir stream, immediately before the module
 * section where the format requires them. Returns the stream unchanged when
 * it has no module section to sit in front of.
 */
export function insertReferenceRecords(dir: Buffer, block: Buffer): Buffer {
	const modules = readDirRecords(dir).find((r) => r.id === REC_PROJECTMODULES);
	if (!modules) {
		return dir;
	}
	return Buffer.concat([dir.subarray(0, modules.start), block, dir.subarray(modules.start)]);
}
