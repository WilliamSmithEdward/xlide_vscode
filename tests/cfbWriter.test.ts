// The compound-file writer past the header's 109 FAT sectors (issue #136)
// and the timestamps of storages it creates (issue #137).

import { describe, expect, it } from 'vitest';
import { Cfb } from '../src/vba/cfb';

const SECTOR = 512;

function headerField(bytes: Buffer, offset: number): number {
	return bytes.readUInt32LE(offset);
}

describe('compound-file writer: DIFAT chain (issue #136)', () => {
	it('writes a container past 109 FAT sectors and reads it back byte for byte', () => {
		// 109 FAT sectors cover 109 * 128 sectors = 7,143,424 bytes; nine
		// megabytes of streams need more, so the DIFAT continues in sectors.
		const cfb = Cfb.createEmpty();
		const big = Buffer.alloc(9 * 1024 * 1024);
		for (let i = 0; i < big.length; i += 4) {
			big.writeUInt32LE((i * 2654435761) >>> 0, i);
		}
		cfb.addStream('Big', big);
		cfb.addStream('Small', Buffer.from('hello'));
		const bytes = cfb.toBytes();
		const nFat = headerField(bytes, 44);
		expect(nFat).toBeGreaterThan(109);
		expect(headerField(bytes, 72)).toBe(Math.ceil((nFat - 109) / 127));
		expect(headerField(bytes, 68)).not.toBe(0xfffffffe);
		const back = Cfb.fromBytes(bytes);
		expect(back.getStream('Big').equals(big)).toBe(true);
		expect(back.getStream('Small').toString()).toBe('hello');
		// Every FAT sector named by the DIFAT chain is marked FATSECT and every
		// DIFAT sector DIFSECT, so a second pass reads the same file again.
		expect(Cfb.fromBytes(back.toBytes()).getStream('Big').equals(big)).toBe(true);
	});

	it('still writes a small container with the header DIFAT alone', () => {
		const cfb = Cfb.createEmpty();
		cfb.addStream('S', Buffer.from('x'.repeat(10000)));
		const bytes = cfb.toBytes();
		expect(headerField(bytes, 72)).toBe(0);
		expect(headerField(bytes, 68)).toBe(0xfffffffe);
		expect(bytes.length % SECTOR).toBe(0);
		expect(Cfb.fromBytes(bytes).getStream('S').length).toBe(10000);
	});
});

describe('compound-file writer: new storage timestamps (issue #137)', () => {
	it('stamps a storage it creates with the current time and leaves streams at zero', () => {
		const cfb = Cfb.createEmpty();
		cfb.addStorageAtPath([], 'Forms');
		cfb.addStreamToStorage('Forms', 'f', Buffer.from('data'));
		const bytes = cfb.toBytes();
		const dirFirst = headerField(bytes, 48);
		const directory = bytes.subarray(SECTOR * (1 + dirFirst));
		const entries: Array<{ name: string; type: number; created: bigint; modified: bigint }> = [];
		for (let i = 0; i < 4; i++) {
			const entry = directory.subarray(i * 128, (i + 1) * 128);
			const nameLength = entry.readUInt16LE(64);
			if (nameLength === 0) { continue; }
			entries.push({
				name: entry.subarray(0, nameLength - 2).toString('utf16le'),
				type: entry.readUInt8(66),
				created: entry.readBigUInt64LE(100),
				modified: entry.readBigUInt64LE(108),
			});
		}
		const storage = entries.find((e) => e.name === 'Forms')!;
		const stream = entries.find((e) => e.name === 'f')!;
		expect(storage.type).toBe(1);
		const nowTicks = (BigInt(Date.now()) + 11644473600000n) * 10000n;
		expect(storage.created).toBeGreaterThan(nowTicks - 600000000000n);
		expect(storage.created).toBeLessThanOrEqual(nowTicks + 10000000n);
		expect(storage.modified).toBe(storage.created);
		expect(stream.created).toBe(0n);
		expect(Cfb.fromBytes(bytes).getStreamInStorage('Forms', 'f').toString()).toBe('data');
	});
});
