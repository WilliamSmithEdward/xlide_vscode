// The one place the container engine touches compression.
//
// Desktop builds get node:zlib. The web extension host has no zlib, so the
// browser build swaps the leaf module underneath this one (see the alias in
// esbuild.js) for a pure-TypeScript implementation. Everything above this
// file - ZipArchive, the CFB reader, every container format - is unchanged
// and stays synchronous either way.

import type { InflateOptions } from './inflate';
import { platformCodec } from './containerCodecNode';

export type { InflateOptions };

/** What a ZIP entry's bytes are, and which method field describes them. */
export interface ZipPayload {
	/** ZIP compression method: 0 stored, 8 deflate. */
	method: number;
	bytes: Buffer;
}

export interface ContainerCodec {
	/** A name for diagnostics and the support bundle. */
	readonly name: string;

	/** Expands a raw DEFLATE stream: a ZIP entry with method 8. */
	inflateRaw(data: Buffer, options?: InflateOptions): Buffer;

	/** Expands an RFC 1950 zlib stream: a PowerPoint compressed storage. */
	inflate(data: Buffer, options?: InflateOptions): Buffer;

	/**
	 * Compresses an entry being written into a ZIP, choosing the method.
	 *
	 * A build with no deflate returns the bytes stored (method 0), which is
	 * valid ZIP and which Office reads. Only the handful of entries an edit
	 * rewrites are affected: ZipArchive carries every untouched entry over
	 * with its original compressed bytes.
	 */
	compressForZip(data: Buffer): ZipPayload;

	/**
	 * Produces an RFC 1950 zlib stream. Unlike a ZIP entry there is no stored
	 * alternative here, so a build without deflate throws instead.
	 */
	deflate(data: Buffer): Buffer;
}

let current: ContainerCodec = platformCodec;

/** The codec in force. Hot path: called once per part read or written. */
export function containerCodec(): ContainerCodec {
	return current;
}

/**
 * Replaces the codec. The web entry point does not need this - its build
 * aliases the leaf module - but tests use it to run the engine through the
 * browser codec on a desktop, which is how the two are held to the same
 * bytes.
 */
export function setContainerCodec(codec: ContainerCodec): void {
	current = codec;
}
