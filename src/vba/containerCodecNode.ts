// zlib-backed codec: what every desktop build and the test suite use.
//
// The browser build never reaches this module - esbuild.js aliases it to
// containerCodecWeb.ts - which is what keeps node:zlib out of the web bundle.

import * as zlib from 'zlib';
import type { ContainerCodec, ZipPayload } from './containerCodec';
import type { InflateOptions } from './inflate';

/**
 * Deflate level for rewritten entries. The dominant cost of saving a project
 * is re-deflating vbaProject.bin, and on a large project level 6 spends about
 * 18 ms to level 4's 10 ms while producing an entry only ~2.5% smaller - well
 * under a percent of the finished project. Ctrl+S happens far more often than
 * anyone counts those bytes, so buy the latency.
 */
const DEFLATE_LEVEL = 4;

export const platformCodec: ContainerCodec = {
	name: 'zlib',

	inflateRaw(data: Buffer, options: InflateOptions = {}): Buffer {
		return zlib.inflateRawSync(data, finish(options));
	},

	inflate(data: Buffer, options: InflateOptions = {}): Buffer {
		return zlib.inflateSync(data, finish(options));
	},

	compressForZip(data: Buffer): ZipPayload {
		return { method: 8, bytes: zlib.deflateRawSync(data, { level: DEFLATE_LEVEL }) };
	},

	deflate(data: Buffer): Buffer {
		return zlib.deflateSync(data, { level: 6 });
	},
};

/**
 * A sync-flush finish accepts a stream that stops without a final block,
 * which is how PowerPoint writes its compressed storages.
 */
function finish(options: InflateOptions): zlib.ZlibOptions {
	return options.allowTruncated ? { finishFlush: zlib.constants.Z_SYNC_FLUSH } : {};
}
