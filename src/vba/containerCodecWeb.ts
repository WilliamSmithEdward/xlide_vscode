// Pure-TypeScript codec for the web extension host, where node:zlib does not
// exist and the platform's own DecompressionStream is asynchronous.
//
// Reading is complete: inflate.ts is held byte-identical to zlib by
// tests/inflate.test.ts. Writing stores entries rather than deflating them,
// which is valid ZIP and which Office reads. Only the parts an edit rewrites
// grow, because ZipArchive carries untouched entries over compressed.

import type { ContainerCodec, ZipPayload } from './containerCodec';
import { inflate, inflateRaw, type InflateOptions } from './inflate';

export const platformCodec: ContainerCodec = {
	name: 'pure-ts',

	inflateRaw(data: Buffer, options: InflateOptions = {}): Buffer {
		return inflateRaw(data, options);
	},

	inflate(data: Buffer, options: InflateOptions = {}): Buffer {
		return inflate(data, options);
	},

	compressForZip(data: Buffer): ZipPayload {
		return { method: 0, bytes: data };
	},

	deflate(): Buffer {
		// Reached only by the PowerPoint writer, whose compressed storage has
		// no stored alternative. Failing here is better than writing a
		// presentation PowerPoint cannot open.
		throw new Error(
			'Saving a PowerPoint presentation is not supported in the browser. Open the file in the desktop editor.',
		);
	},
};
