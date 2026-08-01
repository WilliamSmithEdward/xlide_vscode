// Compound File Binary (CFB / OLE2) reader and writer - [MS-CFB].
//
// vbaProject.bin is a CFB container: a 512-byte header, then fixed-size
// sectors chained by a FAT (and a mini-FAT for streams under 4096 bytes),
// with a red-black directory tree naming the storages and streams.
//
// The writer does a full canonical rebuild rather than patching sectors in
// place: every stream is re-laid-out from scratch in a fresh
// [mini-stream][streams][directory][mini-FAT][FAT] layout. Rebuilding is much
// easier to get provably right than incremental sector surgery, and Excel only
// requires a spec-valid file, not a byte-identical one. Per-entry metadata
// (CLSID, state, colour, timestamps) is preserved verbatim from the original
// entry bytes so nothing outside our control changes.

const MAGIC = Buffer.from([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]);

export const FREESECT = 0xffffffff;
export const ENDOFCHAIN = 0xfffffffe;
const FATSECT = 0xfffffffd;
const NOSTREAM = 0xffffffff;

const OBJTYPE_EMPTY = 0;
const OBJTYPE_STORAGE = 1;
const OBJTYPE_STREAM = 2;
const OBJTYPE_ROOT = 5;

const HEADER_SIZE = 512;
const DIR_ENTRY_SIZE = 128;

/** Output geometry: v3 (512-byte sectors, 64-byte mini-sectors). */
const SECTOR = 512;
const MINI = 64;
const CUTOFF = 4096;
const ENTRIES_PER_SECTOR = SECTOR / 4;
const DIR_ENTRIES_PER_SECTOR = SECTOR / DIR_ENTRY_SIZE;

export class CfbError extends Error {}

interface DirEntry {
	name: string;
	objType: number;
	childId: number;
	leftSiblingId: number;
	rightSiblingId: number;
	startSector: number;
	size: number;
	/** Original 128-byte entry, preserved for round-trip fidelity. */
	raw: Buffer;
}

export class Cfb {
	private sectorSize = 512;
	private miniSectorSize = 64;
	private miniStreamCutoff = 4096;
	private fat: number[] = [];
	private minifat: number[] = [];
	private directory: DirEntry[] = [];
	private miniStream: Buffer = Buffer.alloc(0);
	private readonly overrides = new Map<number, Buffer>();

	private constructor(private readonly data: Buffer) {}

	static fromBytes(data: Buffer): Cfb {
		const cfb = new Cfb(data);
		cfb.parse();
		return cfb;
	}

	// ---------------------------------------------------------------- read

	listStreams(): string[] {
		return this.directory.filter((e) => e.objType === OBJTYPE_STREAM).map((e) => e.name);
	}

	listStorages(): string[] {
		return this.directory.filter((e) => e.objType === OBJTYPE_STORAGE).map((e) => e.name);
	}

	hasStream(name: string): boolean {
		return this.findStreamIndex(name) !== undefined;
	}

	getStream(name: string): Buffer {
		const idx = this.findStreamIndex(name);
		if (idx === undefined) {
			throw new CfbError(`Stream not found: ${name}`);
		}
		return this.readStream(idx);
	}

	/** Stream names that are children of the given storage, in directory order. */
	listStreamsInStorage(storage: string): string[] {
		const parent = this.findStorageIndex(storage);
		return this.collectSubtree(this.directory[parent].childId)
			.filter((i) => this.directory[i].objType === OBJTYPE_STREAM)
			.sort((a, b) => a - b)
			.map((i) => this.directory[i].name);
	}

	getStreamInStorage(storage: string, name: string): Buffer {
		const parent = this.findStorageIndex(storage);
		const idx = this.findChildStreamIndex(parent, name);
		if (idx === undefined) {
			throw new CfbError(`Stream ${name} not found in storage ${storage}`);
		}
		return this.readStream(idx);
	}

	hasStreamInStorage(storage: string, name: string): boolean {
		const parent = this.findStorageIndexOrUndefined(storage);
		return parent !== undefined && this.findChildStreamIndex(parent, name) !== undefined;
	}

	// --------------------------------------------------------------- write

	writeStream(name: string, data: Buffer): void {
		const idx = this.findStreamIndex(name);
		if (idx === undefined) {
			throw new CfbError(`Stream not found: ${name}`);
		}
		this.overrides.set(idx, Buffer.from(data));
	}

	writeStreamInStorage(storage: string, name: string, data: Buffer): void {
		const parent = this.findStorageIndex(storage);
		const idx = this.findChildStreamIndex(parent, name);
		if (idx === undefined) {
			throw new CfbError(`Stream ${name} not found in storage ${storage}`);
		}
		this.overrides.set(idx, Buffer.from(data));
	}

	addStreamToStorage(storage: string, name: string, data: Buffer): void {
		if (!name) {
			throw new CfbError('stream name must be non-empty');
		}
		const parent = this.findStorageIndex(storage);
		if (this.findChildStreamIndex(parent, name) !== undefined) {
			throw new CfbError(`Stream ${name} already exists in storage ${storage}`);
		}
		const target = this.claimSlot(name, OBJTYPE_STREAM);
		this.overrides.set(target, Buffer.from(data));
		const siblings = this.collectSubtree(this.directory[parent].childId);
		siblings.push(target);
		this.directory[parent].childId = this.rebuildBalancedSubtree(siblings);
	}

	/** Remove every stream child of `storage` whose name satisfies `predicate`. */
	dropStreamsInStorage(storage: string, predicate: (name: string) => boolean): string[] {
		const parent = this.findStorageIndexOrUndefined(storage);
		if (parent === undefined) {
			return [];
		}
		const removed = this.collectSubtree(this.directory[parent].childId)
			.filter((i) => this.directory[i].objType === OBJTYPE_STREAM && predicate(this.directory[i].name))
			.map((i) => this.directory[i].name);
		for (const name of removed) {
			this.removeStreamInStorage(storage, name);
		}
		return removed;
	}

	removeStreamInStorage(storage: string, name: string): void {
		const parent = this.findStorageIndex(storage);
		const target = this.findChildStreamIndex(parent, name);
		if (target === undefined) {
			throw new CfbError(`Stream ${name} not found in storage ${storage}`);
		}
		this.unlinkAndClear(parent, target);
	}

	renameStreamInStorage(storage: string, oldName: string, newName: string): void {
		if (!newName) {
			throw new CfbError('new stream name must be non-empty');
		}
		const parent = this.findStorageIndex(storage);
		const target = this.findChildStreamIndex(parent, oldName);
		if (target === undefined) {
			throw new CfbError(`Stream ${oldName} not found in storage ${storage}`);
		}
		if (oldName.toLowerCase() === newName.toLowerCase()) {
			this.directory[target].name = newName;
			return;
		}
		if (this.findChildStreamIndex(parent, newName) !== undefined) {
			throw new CfbError(`Stream ${newName} already exists in storage ${storage}`);
		}
		this.directory[target].name = newName;
		const siblings = this.collectSubtree(this.directory[parent].childId);
		this.directory[parent].childId = this.rebuildBalancedSubtree(siblings);
	}

	// ------------------------------------------------------------ internal

	private parse(): void {
		const data = this.data;
		if (data.length < HEADER_SIZE) {
			throw new CfbError('File too small to be a valid CFB.');
		}
		if (!data.subarray(0, 8).equals(MAGIC)) {
			throw new CfbError('Not a Compound File Binary (magic bytes mismatch).');
		}
		const majorVer = data.readUInt16LE(26);
		if (majorVer !== 3 && majorVer !== 4) {
			throw new CfbError(`Unsupported CFB major version: ${majorVer}`);
		}
		this.sectorSize = 1 << data.readUInt16LE(30);
		this.miniSectorSize = 1 << data.readUInt16LE(32);
		const rootDirStart = data.readUInt32LE(48);
		this.miniStreamCutoff = data.readUInt32LE(56);
		const minifatStart = data.readUInt32LE(60);
		const difatStart = data.readUInt32LE(68);
		const numDifatSectors = data.readUInt32LE(72);

		// DIFAT: 109 entries in the header, then a chain of DIFAT sectors.
		const difat: number[] = [];
		for (let i = 0; i < 109; i++) {
			difat.push(data.readUInt32LE(76 + i * 4));
		}
		if (difatStart !== ENDOFCHAIN && numDifatSectors > 0) {
			let sector = difatStart;
			for (let n = 0; n < numDifatSectors; n++) {
				if (sector === ENDOFCHAIN || sector === FREESECT) {
					break;
				}
				const sectorData = this.sector(sector);
				const perSector = this.sectorSize / 4 - 1;
				for (let i = 0; i < perSector; i++) {
					difat.push(sectorData.readUInt32LE(i * 4));
				}
				sector = sectorData.readUInt32LE(this.sectorSize - 4);
			}
		}

		const fatParts: Buffer[] = [];
		for (const sect of difat) {
			if (sect === FREESECT || sect === ENDOFCHAIN || sect === FATSECT || sect === 0xfffffffc) {
				break;
			}
			fatParts.push(this.sector(sect));
		}
		this.fat = readUint32Array(Buffer.concat(fatParts));

		if (minifatStart !== ENDOFCHAIN) {
			const parts = this.chain(minifatStart).map((s) => this.sector(s));
			this.minifat = readUint32Array(Buffer.concat(parts));
		}

		const dirRaw = Buffer.concat(this.chain(rootDirStart).map((s) => this.sector(s)));
		this.directory = [];
		for (let i = 0; i < Math.floor(dirRaw.length / DIR_ENTRY_SIZE); i++) {
			this.directory.push(parseDirEntry(dirRaw, i));
		}

		const root = this.directory[0];
		if (root && root.startSector !== ENDOFCHAIN && root.startSector !== FREESECT) {
			this.miniStream = Buffer.concat(this.chain(root.startSector).map((s) => this.sector(s)));
		}
	}

	private sector(index: number): Buffer {
		const offset = HEADER_SIZE + index * this.sectorSize;
		const out = this.data.subarray(offset, offset + this.sectorSize);
		// A truncated final sector still has to yield a full-size block.
		return out.length === this.sectorSize
			? out
			: Buffer.concat([out, Buffer.alloc(this.sectorSize - out.length)]);
	}

	private chain(start: number): number[] {
		const sectors: number[] = [];
		const seen = new Set<number>();
		let current = start;
		while (current !== ENDOFCHAIN && current !== FREESECT) {
			if (seen.has(current) || current >= this.fat.length) {
				throw new CfbError(`Cycle or out-of-range sector in FAT chain at ${current}.`);
			}
			seen.add(current);
			sectors.push(current);
			current = this.fat[current];
		}
		return sectors;
	}

	private miniChain(start: number): number[] {
		const sectors: number[] = [];
		const seen = new Set<number>();
		let current = start;
		while (current !== ENDOFCHAIN && current !== FREESECT) {
			if (seen.has(current) || current >= this.minifat.length) {
				throw new CfbError(`Cycle or out-of-range sector in mini-FAT chain at ${current}.`);
			}
			seen.add(current);
			sectors.push(current);
			current = this.minifat[current];
		}
		return sectors;
	}

	private readStream(index: number): Buffer {
		const override = this.overrides.get(index);
		if (override) {
			return override;
		}
		const entry = this.directory[index];
		if (entry.size < this.miniStreamCutoff && entry.objType !== OBJTYPE_ROOT) {
			const parts = this.miniChain(entry.startSector).map((s) =>
				this.miniStream.subarray(s * this.miniSectorSize, (s + 1) * this.miniSectorSize));
			return Buffer.concat(parts).subarray(0, entry.size);
		}
		const parts = this.chain(entry.startSector).map((s) => this.sector(s));
		return Buffer.concat(parts).subarray(0, entry.size);
	}

	private findStreamIndex(name: string): number | undefined {
		const needle = name.toLowerCase();
		for (let i = 0; i < this.directory.length; i++) {
			const e = this.directory[i];
			if (e.objType === OBJTYPE_STREAM && e.name.toLowerCase() === needle) {
				return i;
			}
		}
		return undefined;
	}

	private findStorageIndexOrUndefined(storage: string): number | undefined {
		const needle = storage.toLowerCase();
		for (let i = 0; i < this.directory.length; i++) {
			const e = this.directory[i];
			if (e.objType === OBJTYPE_STORAGE && e.name.toLowerCase() === needle) {
				return i;
			}
		}
		return undefined;
	}

	private findStorageIndex(storage: string): number {
		const idx = this.findStorageIndexOrUndefined(storage);
		if (idx === undefined) {
			throw new CfbError(`Storage not found: ${storage}`);
		}
		return idx;
	}

	private findChildStreamIndex(parentIdx: number, name: string): number | undefined {
		const needle = name.toLowerCase();
		for (const childIdx of this.collectSubtree(this.directory[parentIdx].childId)) {
			const e = this.directory[childIdx];
			if (e.objType === OBJTYPE_STREAM && e.name.toLowerCase() === needle) {
				return childIdx;
			}
		}
		return undefined;
	}

	private collectSubtree(rootId: number): number[] {
		const out: number[] = [];
		const seen = new Set<number>();
		const stack: number[] = [];
		if (rootId !== NOSTREAM && rootId < this.directory.length) {
			stack.push(rootId);
		}
		while (stack.length > 0) {
			const node = stack.pop()!;
			if (seen.has(node) || node === NOSTREAM || node >= this.directory.length) {
				continue;
			}
			seen.add(node);
			out.push(node);
			const entry = this.directory[node];
			if (entry.leftSiblingId !== NOSTREAM) { stack.push(entry.leftSiblingId); }
			if (entry.rightSiblingId !== NOSTREAM) { stack.push(entry.rightSiblingId); }
		}
		return out;
	}

	/** Reuse an EMPTY directory slot when one exists, else append. */
	private claimSlot(name: string, objType: number): number {
		const blank: DirEntry = {
			name,
			objType,
			childId: NOSTREAM,
			leftSiblingId: NOSTREAM,
			rightSiblingId: NOSTREAM,
			startSector: ENDOFCHAIN,
			size: 0,
			raw: Buffer.alloc(0),
		};
		for (let i = 0; i < this.directory.length; i++) {
			if (this.directory[i].objType === OBJTYPE_EMPTY) {
				this.directory[i] = blank;
				return i;
			}
		}
		this.directory.push(blank);
		return this.directory.length - 1;
	}

	/**
	 * Rebuild a sibling subtree as a balanced BST ordered by ([MS-CFB] 2.6.4)
	 * (name length, then upper-cased name).
	 */
	private rebuildBalancedSubtree(indices: number[]): number {
		if (indices.length === 0) {
			return NOSTREAM;
		}
		const ordered = [...indices].sort((a, b) => {
			const na = this.directory[a].name;
			const nb = this.directory[b].name;
			if (na.length !== nb.length) {
				return na.length - nb.length;
			}
			const ua = na.toUpperCase();
			const ub = nb.toUpperCase();
			return ua < ub ? -1 : ua > ub ? 1 : 0;
		});
		return this.buildFromSorted(ordered);
	}

	private buildFromSorted(ordered: number[]): number {
		if (ordered.length === 0) {
			return NOSTREAM;
		}
		const mid = Math.floor(ordered.length / 2);
		const root = ordered[mid];
		this.directory[root].leftSiblingId = this.buildFromSorted(ordered.slice(0, mid));
		this.directory[root].rightSiblingId = this.buildFromSorted(ordered.slice(mid + 1));
		return root;
	}

	private unlinkAndClear(parentIdx: number, targetIdx: number): void {
		const parent = this.directory[parentIdx];
		const siblings = this.collectSubtree(parent.childId).filter((i) => i !== targetIdx);
		parent.childId = this.rebuildBalancedSubtree(siblings);
		this.directory[targetIdx] = {
			name: '',
			objType: OBJTYPE_EMPTY,
			childId: NOSTREAM,
			leftSiblingId: NOSTREAM,
			rightSiblingId: NOSTREAM,
			startSector: FREESECT,
			size: 0,
			raw: Buffer.alloc(0),
		};
		this.overrides.delete(targetIdx);
	}

	// ---------------------------------------------------------- serializer

	/**
	 * Serialize to a fresh, canonical v3 CFB honouring pending overrides.
	 * Layout: [mini-stream][regular streams][directory][mini-FAT][FAT].
	 */
	toBytes(): Buffer {
		// 1. Snapshot every entry's stream bytes.
		const streamBytes: Buffer[] = this.directory.map((entry, i) =>
			entry.objType === OBJTYPE_STREAM ? this.readStream(i) : Buffer.alloc(0));

		// 2. Pack sub-cutoff streams into the mini-stream + mini-FAT.
		const minifat: number[] = [];
		const miniParts: Buffer[] = [];
		const miniStarts = new Map<number, number>();
		let miniLen = 0;
		for (let i = 0; i < this.directory.length; i++) {
			if (this.directory[i].objType !== OBJTYPE_STREAM) { continue; }
			const data = streamBytes[i];
			if (data.length === 0 || data.length >= CUTOFF) { continue; }
			const n = Math.ceil(data.length / MINI);
			const first = minifat.length;
			for (let k = 0; k < n; k++) {
				minifat.push(k + 1 < n ? first + k + 1 : ENDOFCHAIN);
			}
			miniParts.push(data);
			miniLen += data.length;
			const rem = data.length % MINI;
			if (rem !== 0) {
				miniParts.push(Buffer.alloc(MINI - rem));
				miniLen += MINI - rem;
			}
			miniStarts.set(i, first);
		}
		const miniStreamUsedBytes = minifat.length * MINI;
		if (miniLen % SECTOR !== 0) {
			miniParts.push(Buffer.alloc(SECTOR - (miniLen % SECTOR)));
			miniLen += SECTOR - (miniLen % SECTOR);
		}
		const miniStream = Buffer.concat(miniParts, miniLen);
		while (minifat.length % ENTRIES_PER_SECTOR !== 0) {
			minifat.push(FREESECT);
		}

		// 3. Assign sectors.
		const sectors: Buffer[] = [];
		const nMiniStreamSectors = miniStream.length / SECTOR;
		let miniStreamFirst = ENDOFCHAIN;
		if (nMiniStreamSectors > 0) {
			miniStreamFirst = sectors.length;
			for (let k = 0; k < nMiniStreamSectors; k++) {
				sectors.push(miniStream.subarray(k * SECTOR, (k + 1) * SECTOR));
			}
		}

		const regStarts = new Map<number, number>();
		for (let i = 0; i < this.directory.length; i++) {
			if (this.directory[i].objType !== OBJTYPE_STREAM) { continue; }
			const data = streamBytes[i];
			if (data.length < CUTOFF) { continue; }
			const n = Math.ceil(data.length / SECTOR);
			regStarts.set(i, sectors.length);
			const padded = Buffer.concat([data], n * SECTOR);
			for (let k = 0; k < n; k++) {
				sectors.push(padded.subarray(k * SECTOR, (k + 1) * SECTOR));
			}
		}

		const nDirSectors = Math.max(1, Math.ceil(this.directory.length / DIR_ENTRIES_PER_SECTOR));
		const dirParts: Buffer[] = this.directory.map((entry, i) => this.buildDirEntryBytes(
			entry, streamBytes[i].length, miniStarts, regStarts, miniStreamFirst, miniStreamUsedBytes, i));
		const dirBuf = Buffer.concat([
			...dirParts,
			...Array.from(
				{ length: nDirSectors * DIR_ENTRIES_PER_SECTOR - dirParts.length },
				() => emptyDirEntryBytes(),
			),
		], nDirSectors * SECTOR);
		const dirFirst = sectors.length;
		for (let k = 0; k < nDirSectors; k++) {
			sectors.push(dirBuf.subarray(k * SECTOR, (k + 1) * SECTOR));
		}

		let minifatFirst = ENDOFCHAIN;
		let nMinifatSectors = 0;
		if (minifat.length > 0) {
			minifatFirst = sectors.length;
			const minifatBytes = writeUint32Array(minifat);
			nMinifatSectors = minifatBytes.length / SECTOR;
			for (let k = 0; k < nMinifatSectors; k++) {
				sectors.push(minifatBytes.subarray(k * SECTOR, (k + 1) * SECTOR));
			}
		}

		// 4. Size the FAT (fixed point: FAT sectors are themselves in the FAT).
		const nData = sectors.length;
		let nFat = 1;
		for (;;) {
			const needed = Math.ceil((nData + nFat) / ENTRIES_PER_SECTOR);
			if (needed <= nFat) { break; }
			nFat = needed;
		}
		if (nFat > 109) {
			throw new CfbError('CFB writer requires DIFAT chain support for files this large.');
		}
		const fatFirst = nData;

		// 5. Build the FAT.
		const fat = new Array<number>(nFat * ENTRIES_PER_SECTOR).fill(FREESECT);
		const chain = (start: number, n: number): void => {
			for (let k = 0; k < n; k++) {
				fat[start + k] = k + 1 < n ? start + k + 1 : ENDOFCHAIN;
			}
		};
		if (nMiniStreamSectors > 0) { chain(miniStreamFirst, nMiniStreamSectors); }
		for (let i = 0; i < this.directory.length; i++) {
			if (this.directory[i].objType !== OBJTYPE_STREAM) { continue; }
			const data = streamBytes[i];
			if (data.length >= CUTOFF) {
				chain(regStarts.get(i)!, Math.ceil(data.length / SECTOR));
			}
		}
		chain(dirFirst, nDirSectors);
		if (nMinifatSectors > 0) { chain(minifatFirst, nMinifatSectors); }
		for (let k = 0; k < nFat; k++) { fat[fatFirst + k] = FATSECT; }
		const fatBytes = writeUint32Array(fat);
		for (let k = 0; k < nFat; k++) {
			sectors.push(fatBytes.subarray(k * SECTOR, (k + 1) * SECTOR));
		}

		// 6. Header.
		const header = Buffer.alloc(HEADER_SIZE);
		MAGIC.copy(header, 0);
		header.writeUInt16LE(0x003e, 24); // minor version
		header.writeUInt16LE(3, 26);      // major version
		header.writeUInt16LE(0xfffe, 28); // little-endian BOM
		header.writeUInt16LE(9, 30);      // sector shift (512)
		header.writeUInt16LE(6, 32);      // mini-sector shift (64)
		header.writeUInt32LE(0, 40);      // num dir sectors (0 for v3)
		header.writeUInt32LE(nFat, 44);
		header.writeUInt32LE(dirFirst, 48);
		header.writeUInt32LE(0, 52);      // transaction signature
		header.writeUInt32LE(CUTOFF, 56);
		header.writeUInt32LE(minifatFirst, 60);
		header.writeUInt32LE(nMinifatSectors, 64);
		header.writeUInt32LE(ENDOFCHAIN, 68); // first DIFAT sector
		header.writeUInt32LE(0, 72);          // num DIFAT sectors
		for (let i = 0; i < 109; i++) {
			header.writeUInt32LE(i < nFat ? fatFirst + i : FREESECT, 76 + i * 4);
		}

		return Buffer.concat([header, ...sectors]);
	}

	private buildDirEntryBytes(
		entry: DirEntry,
		dataSize: number,
		miniStarts: Map<number, number>,
		regStarts: Map<number, number>,
		miniStreamFirst: number,
		miniStreamUsedBytes: number,
		idx: number,
	): Buffer {
		if (entry.objType === OBJTYPE_EMPTY) {
			return emptyDirEntryBytes();
		}
		let startSector: number;
		let size: number;
		if (entry.objType === OBJTYPE_ROOT) {
			startSector = miniStreamFirst;
			size = miniStreamUsedBytes;
		} else if (entry.objType === OBJTYPE_STREAM) {
			if (dataSize === 0) {
				startSector = ENDOFCHAIN;
				size = 0;
			} else if (dataSize < CUTOFF) {
				startSector = miniStarts.get(idx)!;
				size = dataSize;
			} else {
				startSector = regStarts.get(idx)!;
				size = dataSize;
			}
		} else {
			startSector = ENDOFCHAIN;
			size = 0;
		}

		const buf = entry.raw.length === DIR_ENTRY_SIZE
			? Buffer.from(entry.raw)
			: Buffer.alloc(DIR_ENTRY_SIZE);
		if (entry.raw.length !== DIR_ENTRY_SIZE) {
			buf.writeUInt8(0, 67); // colour: red
		}
		const nameUtf16 = Buffer.concat([Buffer.from(entry.name, 'utf16le'), Buffer.alloc(2)]);
		if (nameUtf16.length > 64) {
			throw new CfbError(`Directory entry name too long: ${entry.name}`);
		}
		buf.fill(0, 0, 64);
		nameUtf16.copy(buf, 0);
		buf.writeUInt16LE(nameUtf16.length, 64);
		buf.writeUInt8(entry.objType, 66);
		buf.writeUInt32LE(entry.leftSiblingId >>> 0, 68);
		buf.writeUInt32LE(entry.rightSiblingId >>> 0, 72);
		buf.writeUInt32LE(entry.childId >>> 0, 76);
		buf.writeUInt32LE(startSector >>> 0, 116);
		buf.writeUInt32LE(size >>> 0, 120);
		buf.writeUInt32LE(Math.floor(size / 0x100000000) >>> 0, 124);
		return buf;
	}
}

function parseDirEntry(raw: Buffer, index: number): DirEntry {
	const offset = index * DIR_ENTRY_SIZE;
	const entryRaw = Buffer.from(raw.subarray(offset, offset + DIR_ENTRY_SIZE));
	const nameLen = entryRaw.readUInt16LE(64);
	const name = entryRaw.subarray(0, Math.max(0, nameLen - 2)).toString('utf16le');
	return {
		name,
		objType: entryRaw.readUInt8(66),
		leftSiblingId: entryRaw.readUInt32LE(68),
		rightSiblingId: entryRaw.readUInt32LE(72),
		childId: entryRaw.readUInt32LE(76),
		startSector: entryRaw.readUInt32LE(116),
		size: entryRaw.readUInt32LE(120),
		raw: entryRaw,
	};
}

function emptyDirEntryBytes(): Buffer {
	const buf = Buffer.alloc(DIR_ENTRY_SIZE);
	buf.writeUInt8(OBJTYPE_EMPTY, 66);
	buf.writeUInt32LE(NOSTREAM, 68);
	buf.writeUInt32LE(NOSTREAM, 72);
	buf.writeUInt32LE(NOSTREAM, 76);
	buf.writeUInt32LE(FREESECT, 116);
	return buf;
}

function readUint32Array(buf: Buffer): number[] {
	const out = new Array<number>(Math.floor(buf.length / 4));
	for (let i = 0; i < out.length; i++) {
		out[i] = buf.readUInt32LE(i * 4);
	}
	return out;
}

function writeUint32Array(values: readonly number[]): Buffer {
	const buf = Buffer.alloc(values.length * 4);
	for (let i = 0; i < values.length; i++) {
		buf.writeUInt32LE(values[i] >>> 0, i * 4);
	}
	return buf;
}
