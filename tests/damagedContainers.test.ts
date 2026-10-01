// Damaged compound files and ZIP packages, as the property tests in
// tests/properties found them. Each reader refuses with its own error; none
// reads past the end of the file or keeps allocating.

import * as fs from 'fs';
import * as path from 'path';
import { describe, expect, it } from 'vitest';
import { Cfb, CfbError } from '../src/vba/cfb';
import { ZipArchive, ZipError } from '../src/vba/zip';

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

	it('refuses a sector past the end of the file rather than inventing one', () => {
		const bytes = emptyCfb();
		bytes.writeUInt32LE(1000, 68);
		bytes.writeUInt32LE(1, 72);
		expect(() => Cfb.fromBytes(bytes)).toThrow(/past the end of the file/);
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

	it('refuses an entry whose local header is past the end', () => {
		const bytes = workbook();
		const central = bytes.readUInt32LE(eocdOf(bytes) + 16);
		bytes.writeUInt32LE(bytes.length - 2, central + 42);
		expect(() => ZipArchive.read(bytes)).toThrow(ZipError);
	});
});
