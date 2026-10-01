// A version 4 compound file: 4096-byte sectors. The header is still 512
// bytes but takes the whole first sector, so sector n starts at
// (n + 1) * 4096 ([MS-CFB] 2.2), not 512 + n * 4096 as for version 3's
// 512-byte sectors.
//
// The fixture was written by Windows itself, StgCreateStorageEx with
// STGOPTIONS { SectorSize: 4096 } (through pywin32), so it is the reference
// implementation's layout rather than a reading of the specification. It
// holds a stream below the mini-stream cutoff, one above it, and a storage
// with two more; each stream holds `pattern(seed, length)`.

import * as fs from 'fs';
import * as path from 'path';
import { describe, expect, it } from 'vitest';
import { Cfb } from '../src/vba/cfb';

const FIXTURE = path.join(__dirname, 'fixtures', 'binaries', 'Version4Storage.cfb');

/** The bytes the fixture's streams were written with; they differ sector to sector. */
function pattern(seed: number, length: number): Buffer {
	const out = Buffer.alloc(length);
	for (let i = 0; i < length; i++) {
		out[i] = (seed * 31 + i * 7 + (i >> 9)) & 0xff;
	}
	return out;
}

const STREAMS: Array<[string | undefined, string, Buffer]> = [
	[undefined, 'Small', pattern(1, 100)],
	[undefined, 'Big', pattern(2, 10000)],
	['VBA', 'dir', pattern(3, 5000)],
	['VBA', 'Module1', pattern(4, 300)],
];

describe('a version 4 compound file', () => {
	it('is what the test says it is', () => {
		const bytes = fs.readFileSync(FIXTURE);
		expect(bytes.readUInt16LE(26)).toBe(4);
		expect(bytes.readUInt16LE(30)).toBe(12);
	});

	it('reads every stream, in the mini stream and in regular sectors', () => {
		const cfb = Cfb.fromBytes(fs.readFileSync(FIXTURE));
		for (const [storage, name, data] of STREAMS) {
			expect(cfb.getStreamIn(storage, name).equals(data), `${storage ?? ''}/${name}`).toBe(true);
		}
	});

	it('is written back as version 3 with the same streams', () => {
		const written = Cfb.fromBytes(fs.readFileSync(FIXTURE)).toBytes();
		expect(written.readUInt16LE(26)).toBe(3);
		const back = Cfb.fromBytes(written);
		for (const [storage, name, data] of STREAMS) {
			expect(back.getStreamIn(storage, name).equals(data), `${storage ?? ''}/${name}`).toBe(true);
		}
	});
});
