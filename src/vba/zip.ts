// Minimal ZIP reader/writer for OOXML packages (.xlsm/.xlsb/.xlam).
//
// Only what Office writes is supported: stored (0) and deflate (8) entries, no
// encryption, no spanning. Entries the caller does not modify are carried over
// with their ORIGINAL compressed bytes, so an edit to one part cannot perturb
// any other part (and saving stays fast on large projects).

import * as zlib from 'zlib';

const SIG_LOCAL = 0x04034b50;
const SIG_CENTRAL = 0x02014b50;
const SIG_EOCD = 0x06054b50;
const SIG_ZIP64_EOCD = 0x06064b50;

/**
 * Deflate level for rewritten entries. The dominant cost of saving a project
 * is re-deflating vbaProject.bin, and on a large project level 6 spends about
 * 18 ms to level 4's 10 ms while producing an entry only ~2.5% smaller - well
 * under a percent of the finished project. Ctrl+S happens far more often than
 * anyone counts those bytes, so buy the latency.
 */
const DEFLATE_LEVEL = 4;

export class ZipError extends Error {}

interface ZipEntry {
	name: string;
	/** Raw bytes as stored in the archive (still compressed when method=8). */
	compressed: Buffer;
	method: number;
	crc32: number;
	uncompressedSize: number;
	flags: number;
	modTime: number;
	modDate: number;
	externalAttrs: number;
	internalAttrs: number;
	versionMadeBy: number;
	extra: Buffer;
	comment: Buffer;
}

const CRC_TABLE = (() => {
	const table = new Int32Array(256);
	for (let n = 0; n < 256; n++) {
		let c = n;
		for (let k = 0; k < 8; k++) {
			c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
		}
		table[n] = c;
	}
	return table;
})();

function crc32(buf: Buffer): number {
	let c = 0xffffffff;
	for (let i = 0; i < buf.length; i++) {
		c = CRC_TABLE[(c ^ buf[i]) & 0xff] ^ (c >>> 8);
	}
	return (c ^ 0xffffffff) >>> 0;
}

export class ZipArchive {
	private readonly entries: ZipEntry[] = [];
	private readonly byName = new Map<string, number>();

	static read(data: Buffer): ZipArchive {
		const zip = new ZipArchive();
		const eocd = findEocd(data);
		let centralOffset = data.readUInt32LE(eocd + 16);
		let entryCount = data.readUInt16LE(eocd + 10);

		// ZIP64: the 32-bit fields saturate; read the real values from the
		// ZIP64 end-of-central-directory record when they do.
		if (centralOffset === 0xffffffff || entryCount === 0xffff) {
			const z64 = data.lastIndexOf(int32le(SIG_ZIP64_EOCD), eocd);
			if (z64 < 0) {
				throw new ZipError('ZIP64 archive without a ZIP64 EOCD record.');
			}
			entryCount = Number(data.readBigUInt64LE(z64 + 32));
			centralOffset = Number(data.readBigUInt64LE(z64 + 48));
		}

		let pos = centralOffset;
		for (let i = 0; i < entryCount; i++) {
			if (data.readUInt32LE(pos) !== SIG_CENTRAL) {
				throw new ZipError(`Bad central directory signature at ${pos}.`);
			}
			const versionMadeBy = data.readUInt16LE(pos + 4);
			const flags = data.readUInt16LE(pos + 8);
			const method = data.readUInt16LE(pos + 10);
			const modTime = data.readUInt16LE(pos + 12);
			const modDate = data.readUInt16LE(pos + 14);
			const crc = data.readUInt32LE(pos + 16);
			let compressedSize = data.readUInt32LE(pos + 20);
			let uncompressedSize = data.readUInt32LE(pos + 24);
			const nameLen = data.readUInt16LE(pos + 28);
			const extraLen = data.readUInt16LE(pos + 30);
			const commentLen = data.readUInt16LE(pos + 32);
			const internalAttrs = data.readUInt16LE(pos + 36);
			const externalAttrs = data.readUInt32LE(pos + 38);
			let localOffset = data.readUInt32LE(pos + 42);
			const name = data.subarray(pos + 46, pos + 46 + nameLen).toString('utf8');
			const extra = Buffer.from(data.subarray(pos + 46 + nameLen, pos + 46 + nameLen + extraLen));
			const comment = Buffer.from(
				data.subarray(pos + 46 + nameLen + extraLen, pos + 46 + nameLen + extraLen + commentLen),
			);

			// ZIP64 extended information extra field (0x0001) overrides saturated sizes.
			if (compressedSize === 0xffffffff || uncompressedSize === 0xffffffff || localOffset === 0xffffffff) {
				let ep = 0;
				while (ep + 4 <= extra.length) {
					const headerId = extra.readUInt16LE(ep);
					const dataSize = extra.readUInt16LE(ep + 2);
					if (headerId === 0x0001) {
						let q = ep + 4;
						if (uncompressedSize === 0xffffffff) { uncompressedSize = Number(extra.readBigUInt64LE(q)); q += 8; }
						if (compressedSize === 0xffffffff) { compressedSize = Number(extra.readBigUInt64LE(q)); q += 8; }
						if (localOffset === 0xffffffff) { localOffset = Number(extra.readBigUInt64LE(q)); }
						break;
					}
					ep += 4 + dataSize;
				}
			}

			// The local header repeats name/extra with its own lengths.
			if (data.readUInt32LE(localOffset) !== SIG_LOCAL) {
				throw new ZipError(`Bad local header signature for ${name}.`);
			}
			const localNameLen = data.readUInt16LE(localOffset + 26);
			const localExtraLen = data.readUInt16LE(localOffset + 28);
			const dataStart = localOffset + 30 + localNameLen + localExtraLen;
			const compressed = Buffer.from(data.subarray(dataStart, dataStart + compressedSize));

			zip.byName.set(name, zip.entries.length);
			zip.entries.push({
				name, compressed, method, crc32: crc, uncompressedSize,
				flags, modTime, modDate, externalAttrs, internalAttrs, versionMadeBy, extra, comment,
			});
			pos += 46 + nameLen + extraLen + commentLen;
		}
		return zip;
	}

	names(): string[] {
		return this.entries.map((e) => e.name);
	}

	has(name: string): boolean {
		return this.byName.has(name);
	}

	read(name: string): Buffer {
		const idx = this.byName.get(name);
		if (idx === undefined) {
			throw new ZipError(`Entry not found: ${name}`);
		}
		const entry = this.entries[idx];
		if (entry.method === 0) {
			return entry.compressed.subarray(0, entry.uncompressedSize);
		}
		if (entry.method === 8) {
			return zlib.inflateRawSync(entry.compressed);
		}
		throw new ZipError(`Unsupported compression method ${entry.method} for ${name}.`);
	}

	/** Replace or create an entry. Untouched entries keep their original bytes. */
	write(name: string, data: Buffer): void {
		const compressed = zlib.deflateRawSync(data, { level: DEFLATE_LEVEL });
		const idx = this.byName.get(name);
		const next: ZipEntry = {
			name,
			compressed,
			method: 8,
			crc32: crc32(data),
			uncompressedSize: data.length,
			flags: 0,
			modTime: idx !== undefined ? this.entries[idx].modTime : 0,
			modDate: idx !== undefined ? this.entries[idx].modDate : 0x21, // 1980-01-01
			externalAttrs: idx !== undefined ? this.entries[idx].externalAttrs : 0,
			internalAttrs: idx !== undefined ? this.entries[idx].internalAttrs : 0,
			versionMadeBy: idx !== undefined ? this.entries[idx].versionMadeBy : 20,
			extra: Buffer.alloc(0),
			comment: idx !== undefined ? this.entries[idx].comment : Buffer.alloc(0),
		};
		if (idx === undefined) {
			this.byName.set(name, this.entries.length);
			this.entries.push(next);
		} else {
			this.entries[idx] = next;
		}
	}

	toBytes(): Buffer {
		const localParts: Buffer[] = [];
		const centralParts: Buffer[] = [];
		let offset = 0;
		for (const entry of this.entries) {
			const nameBuf = Buffer.from(entry.name, 'utf8');
			const local = Buffer.alloc(30);
			local.writeUInt32LE(SIG_LOCAL, 0);
			local.writeUInt16LE(20, 4);                       // version needed
			local.writeUInt16LE(entry.flags & ~0x08, 6);      // no data descriptor
			local.writeUInt16LE(entry.method, 8);
			local.writeUInt16LE(entry.modTime, 10);
			local.writeUInt16LE(entry.modDate, 12);
			local.writeUInt32LE(entry.crc32, 14);
			local.writeUInt32LE(entry.compressed.length, 18);
			local.writeUInt32LE(entry.uncompressedSize, 22);
			local.writeUInt16LE(nameBuf.length, 26);
			local.writeUInt16LE(0, 28);                       // extra dropped (sizes are 32-bit here)
			localParts.push(local, nameBuf, entry.compressed);

			const central = Buffer.alloc(46);
			central.writeUInt32LE(SIG_CENTRAL, 0);
			central.writeUInt16LE(entry.versionMadeBy, 4);
			central.writeUInt16LE(20, 6);
			central.writeUInt16LE(entry.flags & ~0x08, 8);
			central.writeUInt16LE(entry.method, 10);
			central.writeUInt16LE(entry.modTime, 12);
			central.writeUInt16LE(entry.modDate, 14);
			central.writeUInt32LE(entry.crc32, 16);
			central.writeUInt32LE(entry.compressed.length, 20);
			central.writeUInt32LE(entry.uncompressedSize, 24);
			central.writeUInt16LE(nameBuf.length, 28);
			central.writeUInt16LE(0, 30);
			central.writeUInt16LE(entry.comment.length, 32);
			central.writeUInt16LE(0, 34);                     // disk number
			central.writeUInt16LE(entry.internalAttrs, 36);
			central.writeUInt32LE(entry.externalAttrs, 38);
			central.writeUInt32LE(offset, 42);
			centralParts.push(central, nameBuf, entry.comment);

			offset += local.length + nameBuf.length + entry.compressed.length;
		}
		const centralBuf = Buffer.concat(centralParts);
		const eocd = Buffer.alloc(22);
		eocd.writeUInt32LE(SIG_EOCD, 0);
		eocd.writeUInt16LE(0, 4);
		eocd.writeUInt16LE(0, 6);
		eocd.writeUInt16LE(this.entries.length, 8);
		eocd.writeUInt16LE(this.entries.length, 10);
		eocd.writeUInt32LE(centralBuf.length, 12);
		eocd.writeUInt32LE(offset, 16);
		eocd.writeUInt16LE(0, 20);
		return Buffer.concat([...localParts, centralBuf, eocd]);
	}
}

function int32le(value: number): Buffer {
	const buf = Buffer.alloc(4);
	buf.writeUInt32LE(value >>> 0, 0);
	return buf;
}

function findEocd(data: Buffer): number {
	const sig = int32le(SIG_EOCD);
	// The EOCD is last, followed only by an optional <=64KB comment.
	const from = Math.max(0, data.length - 22 - 0xffff);
	const idx = data.lastIndexOf(sig, data.length - 22);
	if (idx < from) {
		throw new ZipError('Not a ZIP archive (no end-of-central-directory record).');
	}
	return idx;
}
