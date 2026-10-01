// Damaged compound files and ZIP packages, as the property tests in
// tests/properties found them. Each reader refuses with its own error; none
// reads past the end of the file or keeps allocating.

import * as fs from 'fs';
import * as path from 'path';
import { describe, expect, it } from 'vitest';
import { Cfb, CfbError } from '../src/vba/cfb';
import { ZipArchive, ZipError } from '../src/vba/zip';
import { readModulesFromBuffer } from '../src/vba/projectService';

const FIXTURES = path.join(__dirname, 'fixtures', 'binaries');
const SECTOR = 512;

/** An empty compound file: the header, a FAT sector and a directory sector. */
function emptyCfb(): Buffer {
	return Buffer.from(Cfb.createEmpty().toBytes());
}

describe('a damaged compound file', () => {
	it('refuses a DIFAT chain that loops, whatever count the header gives', () => {
		// One more sector, whose next-DIFAT pointer is itself. The header asks
		// for four billion DIFAT sectors; following the loop that often
		// collected 127 entries a step until memory ran out.
		const bytes = Buffer.concat([emptyCfb(), Buffer.alloc(SECTOR, 0xff)]);
		const self = (bytes.length - SECTOR - SECTOR) / SECTOR;
		bytes.writeUInt32LE(self, bytes.length - 4);
		bytes.writeUInt32LE(self, 68);
		bytes.writeUInt32LE(0xffffffff, 72);
		expect(() => Cfb.fromBytes(bytes)).toThrow(CfbError);
		expect(() => Cfb.fromBytes(bytes)).toThrow(/Cycle in the DIFAT chain/);
	});

	it('invents no sector past the end of the file: a DIFAT sector there ends the DIFAT', () => {
		const bytes = emptyCfb();
		bytes.writeUInt32LE(1000, 68);
		bytes.writeUInt32LE(1, 72);
		expect(Cfb.fromBytes(bytes).listStreams()).toEqual([]);
	});

	// Issue #340: a file cut short reads the streams that lie before the cut.
	it('reads the VBA modules of a legacy file cut a sector or more short', () => {
		for (const [name, cut] of [['XlsFixture.xls', 59392], ['XlsFixture.xls', 66048], ['WordFixture.doc', 68096]] as const) {
			const whole = fs.readFileSync(path.join(FIXTURES, name));
			const expected = readModulesFromBuffer(whole).map((m) => [m.name, m.code]);
			expect(readModulesFromBuffer(whole.subarray(0, cut)).map((m) => [m.name, m.code]), `${name} cut at ${cut}`).toEqual(expected);
		}
	});

	it('refuses a mini stream that runs into a removed sector, rather than reading zeros', () => {
		// WordFixture.doc cut at 67584 bytes loses the last 41 bytes of `dir`.
		const whole = fs.readFileSync(path.join(FIXTURES, 'WordFixture.doc'));
		expect(() => readModulesFromBuffer(whole.subarray(0, 67584))).toThrow();
	});

	it('refuses a stream whose bytes the cut removed, rather than reading zeros', () => {
		const cfb = Cfb.createEmpty();
		cfb.addStream('Big', Buffer.from('x'.repeat(8192)));
		const bytes = cfb.toBytes();
		const cut = Cfb.fromBytes(bytes.subarray(0, bytes.length - 3 * SECTOR));
		expect(() => cut.getStream('Big')).toThrow(CfbError);
	});

	it('refuses a sector size the format does not have', () => {
		for (const [offset, shift] of [[30, 0], [30, 31], [32, 0]]) {
			const bytes = emptyCfb();
			bytes.writeUInt16LE(shift, offset);
			expect(() => Cfb.fromBytes(bytes)).toThrow(/Unsupported sector sizes/);
		}
	});

	it('still reads a file whose last sector is cut short', () => {
		const cfb = Cfb.createEmpty();
		cfb.addStream('S', Buffer.from('y'.repeat(5000)));
		const bytes = cfb.toBytes();
		const cut = bytes.subarray(0, bytes.length - 100);
		expect(Cfb.fromBytes(cut).getStream('S').length).toBe(5000);
	});
});

describe('a damaged ZIP package', () => {
	const workbook = (): Buffer => Buffer.from(fs.readFileSync(path.join(FIXTURES, 'NoVbaFixture.xlsm')));
	const eocdOf = (bytes: Buffer): number => bytes.lastIndexOf(Buffer.from([0x50, 0x4b, 0x05, 0x06]));

	it('refuses a central directory that starts past the end', () => {
		const bytes = workbook();
		bytes.writeUInt32LE(bytes.length - 2, eocdOf(bytes) + 16);
		expect(() => ZipArchive.read(bytes)).toThrow(ZipError);
	});

	// Issue #340: a damaged part VBA does not need leaves the package
	// readable; reading that part, or writing the package, is refused.
	it('reads a package with one entry whose local header is past the end', () => {
		const bytes = workbook();
		const central = bytes.readUInt32LE(eocdOf(bytes) + 16);
		bytes.writeUInt32LE(bytes.length - 2, central + 42);
		const zip = ZipArchive.read(bytes);
		const damaged = zip.names()[0];
		expect(() => zip.read(damaged)).toThrow(/Bad local header signature/);
		expect(zip.read(zip.names()[1]).length).toBeGreaterThan(0);
		expect(() => zip.toBytes()).toThrow(ZipError);
	});

	it('refuses to read an entry whose data runs past the end', () => {
		const bytes = workbook();
		const central = bytes.readUInt32LE(eocdOf(bytes) + 16);
		bytes.writeUInt32LE(bytes.length, central + 20);
		const zip = ZipArchive.read(bytes);
		expect(() => zip.read(zip.names()[0])).toThrow(/runs past the end/);
	});
});
