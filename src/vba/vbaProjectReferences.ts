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

import { readDirRecords, REC_PROJECTMODULES } from './vbaProject';

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

/** Every reference the dir stream declares, in order. */
export function readProjectReferences(dir: Buffer): VbaProjectReference[] {
	const out: VbaProjectReference[] = [];
	let name = '';
	for (const record of readDirRecords(dir)) {
		const body = dir.subarray(record.dataStart, record.dataEnd);
		switch (record.id) {
			case REC_REFERENCE_NAME:
				name = body.toString('latin1');
				break;
			case REC_REFERENCE_REGISTERED:
				// SizeOfLibid then the libid; the trailing reserved fields are ignored.
				out.push({ name, kind: 'registered', libid: sizedString(body) });
				break;
			case REC_REFERENCE_PROJECT:
				out.push({ name, kind: 'project', libid: sizedString(body) });
				break;
			case REC_REFERENCE_ORIGINAL:
				// The record's own size IS the libid's, so the body is the libid.
				out.push({ name, kind: 'control', libid: body.toString('latin1') });
				break;
			default:
				break;
		}
	}
	return out;
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

function record(id: number, body: Buffer): Buffer {
	const head = Buffer.alloc(6);
	head.writeUInt16LE(id, 0);
	head.writeUInt32LE(body.length, 2);
	return Buffer.concat([head, body]);
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
		record(REC_REFERENCE_NAME, Buffer.from(name, 'latin1')),
		record(REC_REFERENCE_NAME_UNICODE, Buffer.from(name, 'utf16le')),
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
		record(REC_REFERENCE_ORIGINAL, Buffer.from(reference.libidOriginal, 'latin1')),
		record(REC_REFERENCE_CONTROL, control),
		nameRecords(reference.name),
		record(REC_REFERENCE_CONTROL_EXTENDED, extendedBody),
	]);
}

/** The libid text for a type library, as the format spells it. */
export function libid(guid: string, version: string, path: string, description: string): string {
	return `*\\G${guid}#${version}#0#${path}#${description}`;
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
