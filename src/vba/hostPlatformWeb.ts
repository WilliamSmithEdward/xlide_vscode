// The web extension host's platform: an in-memory filesystem the extension
// layer fills and drains around each engine call.
//
// A browser has no synchronous file access, and the engine is synchronous
// throughout. The two are reconciled here rather than by making 170 files
// async: because every engine operation reads a whole container, works in
// memory and writes it back, the extension layer can read the bytes through
// vscode.workspace.fs first, prime them here, call in, and flush whatever the
// engine wrote. Priming a file it turns out not to need costs one read.
//
// A path that was never primed throws rather than reading as absent: an
// engine that silently saw an empty project would write a workbook with the
// user's code missing, which is the one failure worth being loud about.

import type { HostFileStat, HostPlatform } from './hostPlatform';

interface CachedFile {
	data: Buffer;
	mtimeMs: number;
}

/** Primed contents, by path. A null entry means "primed, and absent". */
const files = new Map<string, CachedFile | null>();

/** Writes the engine has made since the last drain, by path. */
const pendingWrites = new Map<string, Buffer>();

export class HostFileNotPrimedError extends Error {
	constructor(filePath: string) {
		super(
			`XLIDE has not loaded ${filePath}. This is an internal error: the file should have been read before the operation started.`,
		);
	}
}

/**
 * Supplies a file's bytes, after the extension layer has read them through
 * vscode.workspace.fs. `mtimeMs` must be the workspace's own stamp, because
 * the engine pairs it with the size to decide whether a parsed project is
 * still current.
 */
export function primeHostFile(filePath: string, data: Buffer, mtimeMs: number): void {
	files.set(filePath, { data, mtimeMs });
}

/** Records that a path has been checked and does not exist. */
export function primeHostFileAbsent(filePath: string): void {
	files.set(filePath, null);
}

/** Whether a path has been primed, either way. */
export function isHostFilePrimed(filePath: string): boolean {
	return files.has(filePath);
}

/**
 * The stamp of what is currently primed, or undefined when the path has not
 * been primed or is primed as absent. The priming step compares this against
 * the workspace's own stat so an unchanged container is not re-read on every
 * call - which for a large workbook is the difference between a responsive
 * editor and a stalled one.
 */
export function primedStamp(filePath: string): HostFileStat | undefined {
	const entry = files.get(filePath);
	return entry ? { mtimeMs: entry.mtimeMs, size: entry.data.length } : undefined;
}

/**
 * Hands back everything the engine wrote and clears the pending set, for the
 * extension layer to flush through vscode.workspace.fs. The bytes stay in the
 * cache, so a follow-up read in the same operation sees them.
 */
export function takeHostFileWrites(): Map<string, Buffer> {
	const writes = new Map(pendingWrites);
	pendingWrites.clear();
	return writes;
}

/**
 * Drops primed contents so the next operation reads the workspace again.
 * Pending writes are kept: losing them would lose the user's edit.
 */
export function forgetHostFiles(): void {
	files.clear();
}

function lookup(filePath: string): CachedFile | null {
	const entry = files.get(filePath);
	if (entry === undefined) {
		throw new HostFileNotPrimedError(filePath);
	}
	return entry;
}

export const platformHost: HostPlatform = {
	name: 'web',

	readFile(filePath: string): Buffer {
		const entry = lookup(filePath);
		if (!entry) {
			throw new Error(`File not found: ${filePath}`);
		}
		return entry.data;
	},

	readFileIfPresent(filePath: string): Buffer | undefined {
		return lookup(filePath)?.data;
	},

	stat(filePath: string): HostFileStat {
		const entry = lookup(filePath);
		if (!entry) {
			throw new Error(`File not found: ${filePath}`);
		}
		return { mtimeMs: entry.mtimeMs, size: entry.data.length };
	},

	statIfPresent(filePath: string): HostFileStat | undefined {
		const entry = lookup(filePath);
		return entry ? { mtimeMs: entry.mtimeMs, size: entry.data.length } : undefined;
	},

	exists(filePath: string): boolean {
		return lookup(filePath) !== null;
	},

	/**
	 * Atomicity is the workspace's problem here: vscode.workspace.fs.writeFile
	 * replaces a file in one operation, so the engine's temp-and-rename dance
	 * has nothing to add. The new bytes land in the cache immediately and are
	 * queued for the flush.
	 */
	writeFile(filePath: string, data: Buffer): void {
		files.set(filePath, { data, mtimeMs: Date.now() });
		pendingWrites.set(filePath, data);
	},

	randomBytes(count: number): Buffer {
		const out = new Uint8Array(count);
		crypto.getRandomValues(out);
		return Buffer.from(out);
	},
};
