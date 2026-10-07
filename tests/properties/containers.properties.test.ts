import { describe, it } from 'vitest';
import fc from 'fast-check';
import * as fs from 'fs';
import * as path from 'path';
import * as zlib from 'zlib';
import { AccessFormatError } from '../../src/vba/access/accessFormat';
import { AccessDatabaseError } from '../../src/vba/accessDatabase';
import { Cfb, CfbError } from '../../src/vba/cfb';
import { InflateError, inflate, inflateRaw } from '../../src/vba/inflate';
import { MacroContainerError } from '../../src/vba/macroContainer';
import { NoVbaProjectError } from '../../src/vba/noVbaProject';
import { OvbaError, compress, decompress } from '../../src/vba/ovba';
import { PptContainerError } from '../../src/vba/pptContainer';
import { readModulesFromBuffer } from '../../src/vba/projectService';
import { VbaProjectLockedError } from '../../src/vba/projectProtection';
import { VbaProject, VbaProjectError } from '../../src/vba/vbaProject';
import { XlsxError } from '../../src/vba/xlsx';
import { ZipArchive, ZipError } from '../../src/vba/zip';
import { PROPERTY_RUNS, propertySettings } from './settings';

// Properties of the readers that take a macro container apart: ZIP, deflate,
// the compound file, MS-OVBA compression and the VBA project records. A
// container is whatever file the user opens, often somebody else's, so each
// reader must either read it or refuse it with its own error, and never run
// away. Every round trip below is the reader against its own writer (or
// against zlib), and every "refuses" property feeds it real fixtures with
// bytes changed.
//
// The Fuzz workflow raises XLIDE_PROPERTY_RUNS; see ./settings.ts.

const FIXTURES = path.join(__dirname, '..', 'fixtures', 'binaries');
const fixture = (name: string): Buffer => fs.readFileSync(path.join(FIXTURES, name));

/** The errors the container readers raise for a file they cannot read. */
const REFUSALS = [
	AccessDatabaseError,
	AccessFormatError,
	CfbError,
	InflateError,
	MacroContainerError,
	NoVbaProjectError,
	OvbaError,
	PptContainerError,
	VbaProjectError,
	VbaProjectLockedError,
	XlsxError,
	ZipError,
];

/**
 * zlib's refusal of a corrupt stream (Z_DATA_ERROR, Z_BUF_ERROR): the desktop
 * codec inflates with node:zlib, and its errors carry a Z_ code. A buffer too
 * large to allocate carries a different code and is not a refusal.
 */
function isZlibRefusal(error: unknown): boolean {
	const code = (error as { code?: unknown } | undefined)?.code;
	return error instanceof Error && typeof code === 'string' && code.startsWith('Z_');
}

/** No single input may take longer than this; a slower one is a runaway. */
const PER_INPUT_MS = 5000;

/** Runs a reader over hostile bytes: it may refuse, but only with its own error, and promptly. */
function readsOrRefuses(read: () => void): void {
	const started = Date.now();
	try {
		read();
	} catch (error) {
		if (!REFUSALS.some((refusal) => error instanceof refusal) && !isZlibRefusal(error)) {
			throw error;
		}
	}
	const took = Date.now() - started;
	if (took > PER_INPUT_MS) {
		throw new Error(`took ${took} ms`);
	}
}

/**
 * A real file with some of its bytes changed. Half the edits land in the
 * first 4 KiB, where headers, sector tables and directories are; an edit is a
 * random byte, a 32-bit field forced to a boundary value (lengths, sector
 * numbers, offsets), or the file cut short.
 */
function mutated(seed: Buffer): fc.Arbitrary<Buffer> {
	const offset = fc.oneof(
		fc.integer({ min: 0, max: Math.min(4095, seed.length - 1) }),
		fc.integer({ min: 0, max: seed.length - 1 }),
	);
	const edit = fc.oneof(
		fc.record({ kind: fc.constant('byte' as const), at: offset, value: fc.integer({ min: 0, max: 255 }) }),
		fc.record({
			kind: fc.constant('u32' as const),
			at: offset,
			value: fc.constantFrom(0, 1, 0x7fffffff, 0x80000000, 0xfffffffa, 0xfffffffc, 0xfffffffd, 0xfffffffe, 0xffffffff),
		}),
		fc.record({ kind: fc.constant('cut' as const), at: offset, value: fc.constant(0) }),
	);
	return fc.array(edit, { minLength: 1, maxLength: 6 }).map((edits) => {
		let bytes = Buffer.from(seed);
		for (const { kind, at, value } of edits) {
			if (kind === 'byte' && at < bytes.length) {
				bytes[at] = value;
			} else if (kind === 'u32' && at + 4 <= bytes.length) {
				bytes.writeUInt32LE(value, at);
			} else if (kind === 'cut') {
				bytes = bytes.subarray(0, at);
			}
		}
		return bytes;
	});
}

/** Every module's source of a container, the way the extension lists a project. */
function readEveryModule(bytes: Buffer): void {
	for (const entry of readModulesFromBuffer(bytes, true)) {
		void entry.source;
	}
}

/** A compound file read stream by stream, then as a VBA project with every module's source. */
function readVbaProject(bytes: Buffer): void {
	const cfb = Cfb.fromBytes(bytes);
	for (const name of cfb.listStreams()) {
		cfb.getStream(name);
	}
	for (const module of VbaProject.parse(cfb).modules) {
		void module.source;
	}
}

const bytes = (maxLength: number): fc.Arbitrary<Buffer> =>
	fc.uint8Array({ maxLength }).map((array) => Buffer.from(array));

/** Bytes that compress: runs and repeats, which is what back-references are for. */
const repetitive = fc
	.array(fc.tuple(bytes(16), fc.integer({ min: 1, max: 300 })), { maxLength: 20 })
	.map((runs) => Buffer.concat(runs.map(([run, times]) => Buffer.concat(Array(times).fill(run)))));

describe('MS-OVBA compression', () => {
	it('decompresses what it compresses', () => {
		fc.assert(fc.property(fc.oneof(bytes(20000), repetitive), (data) => {
			const back = decompress(compress(data));
			if (!back.equals(data)) {
				throw new Error(`${data.length} bytes came back as ${back.length}`);
			}
		}), propertySettings());
	});

	it('reads or refuses any stream', () => {
		fc.assert(fc.property(bytes(8000), fc.option(fc.integer({ min: 0, max: 20000 }), { nil: undefined }), (data, maxBytes) => {
			readsOrRefuses(() => decompress(Buffer.concat([Buffer.from([1]), data]), 'fuzz', maxBytes));
		}), propertySettings());
	});
});

describe('inflate', () => {
	it('matches zlib on whatever zlib deflates', () => {
		fc.assert(fc.property(fc.oneof(bytes(20000), repetitive), fc.integer({ min: 0, max: 9 }), (data, level) => {
			if (!inflateRaw(zlib.deflateRawSync(data, { level })).equals(data)) {
				throw new Error('raw deflate came back different');
			}
			if (!inflate(zlib.deflateSync(data, { level })).equals(data)) {
				throw new Error('zlib-wrapped deflate came back different');
			}
		}), propertySettings());
	});

	it('reads or refuses any stream, whatever size it is told to expect', () => {
		fc.assert(fc.property(
			bytes(4000),
			fc.option(fc.oneof(fc.nat(100000), fc.constantFrom(0x7fffffff, 0xffffffff)), { nil: undefined }),
			fc.boolean(),
			(data, expectedSize, allowTruncated) => {
				readsOrRefuses(() => inflateRaw(data, { expectedSize, allowTruncated }));
				readsOrRefuses(() => inflate(data, { expectedSize, allowTruncated }));
			},
		), propertySettings());
	});
});

describe('ZIP', () => {
	const entryName = fc.stringMatching(/^[A-Za-z0-9_][A-Za-z0-9_./-]{0,40}$/);

	it('reads back what it writes', () => {
		fc.assert(fc.property(fc.uniqueArray(fc.tuple(entryName, fc.oneof(bytes(5000), repetitive)), {
			maxLength: 8,
			selector: ([name]) => name,
		}), (files) => {
			const zip = ZipArchive.read(ZipArchive.read(fixture('NoVbaFixture.xlsm')).toBytes());
			for (const name of zip.names()) {
				zip.delete(name);
			}
			for (const [name, data] of files) {
				zip.write(name, data);
			}
			const back = ZipArchive.read(zip.toBytes());
			for (const [name, data] of files) {
				if (!back.read(name).equals(data)) {
					throw new Error(`${name} came back different`);
				}
			}
			if (back.names().length !== files.length) {
				throw new Error(`${files.length} entries came back as ${back.names().length}`);
			}
		}), propertySettings());
	});

	it('reads or refuses a changed workbook package', () => {
		const seed = fixture('FormFixture.xlsm');
		fc.assert(fc.property(mutated(seed), (data) => {
			readsOrRefuses(() => {
				const zip = ZipArchive.read(data);
				for (const name of zip.names()) {
					zip.read(name);
				}
			});
		}), propertySettings());
	});
});

describe('compound file', () => {
	const streamName = fc.stringMatching(/^[A-Za-z0-9_ ]{1,31}$/);

	it('reads back the streams it writes, small and large', () => {
		fc.assert(fc.property(fc.uniqueArray(fc.tuple(streamName, fc.oneof(bytes(6000), repetitive)), {
			maxLength: 6,
			selector: ([name]) => name.toLowerCase(),
		}), (streams) => {
			const cfb = Cfb.createEmpty();
			for (const [name, data] of streams) {
				cfb.addStream(name, data);
			}
			const back = Cfb.fromBytes(cfb.toBytes());
			for (const [name, data] of streams) {
				if (!back.getStream(name).equals(data)) {
					throw new Error(`${name} (${data.length} bytes) came back different`);
				}
			}
		}), propertySettings());
	});

	it('reads or refuses a changed VBA project', () => {
		const seed = ZipArchive.read(fixture('FormFixture.xlsm')).read('xl/vbaProject.bin');
		fc.assert(fc.property(mutated(seed), (data) => {
			readsOrRefuses(() => readVbaProject(data));
		}), propertySettings());
	});
});

describe('macro containers', () => {
	// One of each container the extension opens: OOXML for each host, a
	// legacy compound file, a PowerPoint 97 file with compressed storage, and
	// an Access database.
	const SEEDS = [
		'FormFixture.xlsm',
		'WordFixture.docm',
		'PowerPointFixture.pptm',
		'ExcelFixture.xla',
		'WordFixture.doc',
		'PowerPointFixture.ppt',
		'AccessFixture.mdb',
	];

	for (const name of SEEDS) {
		it(`reads or refuses a changed ${name}`, () => {
			const seed = fixture(name);
			readEveryModule(seed);
			fc.assert(fc.property(mutated(seed), (data) => {
				readsOrRefuses(() => readEveryModule(data));
			}), propertySettings(Math.max(1, Math.round(PROPERTY_RUNS / 4))));
		});
	}
});
